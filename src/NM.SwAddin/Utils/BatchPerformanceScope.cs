using System;
using NM.Core;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.Utils
{
    /// <summary>
    /// RAII wrapper that suppresses SolidWorks UI overhead during batch operations.
    /// Automatically restores all state on dispose.
    ///
    /// Usage:
    ///   using (new BatchPerformanceScope(swApp, doc)) { /* batch work */ }
    ///   using (new BatchPerformanceScope(swApp, doc, debugMode: true)) { /* visible for AI debugging */ }
    ///
    /// Performance optimizations applied (non-debug mode):
    /// - CommandInProgress=true:       Prevents undo record consolidation (~significant)
    /// - EnableGraphicsUpdate=false:   Stops viewport redraws (~3.2x speedup per CAD Booster benchmark)
    /// - EnableFeatureTree=false:      Stops feature tree data updates (~30% reduction)
    /// - EnableFeatureTreeWindow=false: Stops feature tree window rendering (additional ~10%)
    ///
    /// In debug mode, CommandInProgress is still set (no visual effect, just undo batching),
    /// but all UI suppression is skipped so the operator/AI can observe what SolidWorks is doing.
    /// All logging and performance tracking remain active in both modes.
    /// </summary>
    public sealed class BatchPerformanceScope : IDisposable
    {
        private readonly ISldWorks _swApp;
        private readonly IModelDoc2 _doc;
        private readonly bool _debugMode;
        private bool _disposed;

        // Saved states for restore
        private bool _featureTreeWasEnabled = true;
        private bool _featureTreeWindowWasEnabled = true;
        private bool _graphicsUpdateWasEnabled = true;

        // Track which optimizations were actually applied (for logging)
        private bool _appliedGraphicsUpdate;
        private bool _appliedFeatureTree;
        private bool _appliedFeatureTreeWindow;

        /// <summary>
        /// Creates scope that suppresses UI overhead for batch operations.
        /// </summary>
        /// <param name="swApp">SolidWorks application</param>
        /// <param name="doc">Optional: specific document (null = no document-level optimization)</param>
        /// <param name="debugMode">
        /// When true, skip UI suppression so the model/tree remain visible for debugging.
        /// CommandInProgress and all logging/timing still apply.
        /// Defaults to the global Configuration.Logging.EnableDebugMode setting.
        /// </param>
        public BatchPerformanceScope(ISldWorks swApp, IModelDoc2 doc = null, bool? debugMode = null)
        {
            _swApp = swApp;
            _doc = doc;
            _debugMode = debugMode ?? Configuration.Logging.EnableDebugMode;

            if (_swApp == null)
            {
                ErrorHandler.LogInfo("[PERF] BatchPerformanceScope: swApp is null, skipping optimization");
                return;
            }

            try
            {
                PerformanceTracker.Instance.StartTimer("BatchPerformanceScope");

                // CommandInProgress: always set regardless of debug mode.
                // This only affects undo consolidation, not visibility.
                _swApp.CommandInProgress = true;

                if (_doc != null && !_debugMode)
                {
                    // --- Graphics Update (IModelView) ---
                    // This is the single biggest perf win (~3.2x). Lives on IModelView, not ModelDocExtension.
                    try
                    {
                        var view = _doc.ActiveView as IModelView;
                        if (view != null)
                        {
                            _graphicsUpdateWasEnabled = view.EnableGraphicsUpdate;
                            view.EnableGraphicsUpdate = false;
                            _appliedGraphicsUpdate = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorHandler.LogInfo($"[PERF] BatchPerformanceScope: EnableGraphicsUpdate not available: {ex.Message}");
                    }

                    // --- Feature Tree ---
                    var fm = _doc.FeatureManager;
                    if (fm != null)
                    {
                        try
                        {
                            _featureTreeWasEnabled = fm.EnableFeatureTree;
                            fm.EnableFeatureTree = false;
                            _appliedFeatureTree = true;
                        }
                        catch (Exception ex)
                        {
                            ErrorHandler.LogInfo($"[PERF] BatchPerformanceScope: EnableFeatureTree failed: {ex.Message}");
                        }

                        // --- Feature Tree Window (companion to EnableFeatureTree) ---
                        try
                        {
                            _featureTreeWindowWasEnabled = fm.EnableFeatureTreeWindow;
                            fm.EnableFeatureTreeWindow = false;
                            _appliedFeatureTreeWindow = true;
                        }
                        catch (Exception ex)
                        {
                            ErrorHandler.LogInfo($"[PERF] BatchPerformanceScope: EnableFeatureTreeWindow failed: {ex.Message}");
                        }
                    }
                }

                ErrorHandler.LogInfo($"[PERF] BatchPerformanceScope: Entered " +
                    $"(debug={_debugMode}, graphics={_appliedGraphicsUpdate}, " +
                    $"tree={_appliedFeatureTree}, treeWindow={_appliedFeatureTreeWindow})");
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("BatchPerformanceScope.ctor", "Failed to enter batch scope", ex, ErrorHandler.LogLevel.Warning);
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                // Restore in reverse order of application
                if (_doc != null)
                {
                    // Feature tree window first (companion property)
                    if (_appliedFeatureTreeWindow)
                    {
                        try { _doc.FeatureManager.EnableFeatureTreeWindow = _featureTreeWindowWasEnabled; }
                        catch { }
                    }

                    // Feature tree data
                    if (_appliedFeatureTree)
                    {
                        try { _doc.FeatureManager.EnableFeatureTree = _featureTreeWasEnabled; }
                        catch { }
                    }

                    // Graphics update (restore last so tree redraws first)
                    if (_appliedGraphicsUpdate)
                    {
                        try
                        {
                            var view = _doc.ActiveView as IModelView;
                            if (view != null) view.EnableGraphicsUpdate = _graphicsUpdateWasEnabled;
                        }
                        catch { }
                    }
                }

                if (_swApp != null)
                {
                    try { _swApp.CommandInProgress = false; }
                    catch { }
                }

                PerformanceTracker.Instance.StopTimer("BatchPerformanceScope");
                ErrorHandler.LogInfo("[PERF] BatchPerformanceScope: Exited, all state restored");
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("BatchPerformanceScope.Dispose", "Failed to restore batch scope", ex, ErrorHandler.LogLevel.Warning);
            }
        }
    }
}
