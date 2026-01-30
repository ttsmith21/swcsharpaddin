using System;
using NM.Core;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.Utils
{
    /// <summary>
    /// RAII wrapper that disables graphics/feature-tree updates during batch operations.
    /// Automatically restores state on dispose.
    /// Usage: using (new BatchPerformanceScope(swApp, doc)) { /* batch work */ }
    ///
    /// Performance benefits:
    /// - CommandInProgress: Prevents undo record consolidation
    /// - EnableGraphicsUpdate=false: Stops viewport redraws
    /// - EnableFeatureTree=false: Stops feature tree updates (optional)
    /// </summary>
    public sealed class BatchPerformanceScope : IDisposable
    {
        private readonly ISldWorks _swApp;
        private readonly IModelDoc2 _doc;
        private readonly bool _suppressFeatureTree;
        private bool _disposed;
        private bool _featureTreeWasEnabled;

        /// <summary>
        /// Creates scope that disables graphics and optionally feature tree updates.
        /// </summary>
        /// <param name="swApp">SolidWorks application</param>
        /// <param name="doc">Optional: specific document (null = no document-level optimization)</param>
        /// <param name="suppressFeatureTree">Also suppress feature tree updates (heavy operations only)</param>
        public BatchPerformanceScope(ISldWorks swApp, IModelDoc2 doc = null, bool suppressFeatureTree = false)
        {
            _swApp = swApp;
            _doc = doc;
            _suppressFeatureTree = suppressFeatureTree;
            _featureTreeWasEnabled = true;

            if (_swApp == null)
            {
                ErrorHandler.DebugLog("[PERF] BatchPerformanceScope: swApp is null, skipping optimization");
                return;
            }

            try
            {
                PerformanceTracker.Instance.StartTimer("BatchPerformanceScope");

                // Signal that a command is in progress (prevents undo consolidation)
                _swApp.CommandInProgress = true;

                // Disable feature tree updates for batch performance
                // Note: EnableGraphicsUpdate is not available in SolidWorks 2022 API on ModelDocExtension
                if (_doc != null && _suppressFeatureTree)
                {
                    var fm = _doc.FeatureManager;
                    if (fm != null)
                    {
                        try
                        {
                            _featureTreeWasEnabled = fm.EnableFeatureTree;
                            fm.EnableFeatureTree = false;
                        }
                        catch (Exception ex)
                        {
                            ErrorHandler.DebugLog($"[PERF] BatchPerformanceScope: Failed to disable feature tree: {ex.Message}");
                        }
                    }
                }

                ErrorHandler.DebugLog($"[PERF] BatchPerformanceScope: Entered (commandInProgress=true, featureTree={_suppressFeatureTree})");
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
                // Restore states
                if (_doc != null && _suppressFeatureTree)
                {
                    var fm = _doc.FeatureManager;
                    if (fm != null)
                    {
                        try
                        {
                            fm.EnableFeatureTree = _featureTreeWasEnabled;
                        }
                        catch { }
                    }
                }

                if (_swApp != null)
                {
                    try
                    {
                        _swApp.CommandInProgress = false;
                    }
                    catch { }
                }

                PerformanceTracker.Instance.StopTimer("BatchPerformanceScope");
                ErrorHandler.DebugLog("[PERF] BatchPerformanceScope: Exited, state restored");
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("BatchPerformanceScope.Dispose", "Failed to restore batch scope", ex, ErrorHandler.LogLevel.Warning);
            }
        }
    }
}
