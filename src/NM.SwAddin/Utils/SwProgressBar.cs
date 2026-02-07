using System;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Utils
{
    /// <summary>
    /// Lightweight wrapper around SolidWorks' built-in IUserProgressBar.
    /// Provides progress in the SolidWorks status bar AND automatically mirrors
    /// it onto the Windows taskbar icon (green overlay) -- no extra code needed.
    ///
    /// Also supports Escape-to-cancel via the UserCanceled property.
    ///
    /// Usage:
    ///   using (var progress = new SwProgressBar(swApp, fileCount, "Processing parts"))
    ///   {
    ///       for (int i = 0; i &lt; fileCount; i++)
    ///       {
    ///           progress.Update(i, $"Part {i+1}/{fileCount}: {fileName}");
    ///           if (progress.UserCanceled) break;
    ///           // ... do work ...
    ///       }
    ///   }
    /// </summary>
    public sealed class SwProgressBar : IDisposable
    {
        private UserProgressBar _prgBar;
        private bool _disposed;
        private readonly int _total;

        /// <summary>
        /// True if the user pressed Escape during the last UpdateProgress call.
        /// Check this after each Update() call to support cancellation.
        /// </summary>
        public bool UserCanceled { get; private set; }

        /// <summary>
        /// Creates and starts a SolidWorks native progress bar.
        /// </summary>
        /// <param name="swApp">SolidWorks application instance</param>
        /// <param name="total">Total number of steps (upper bound)</param>
        /// <param name="title">Initial title text shown in the status bar</param>
        public SwProgressBar(ISldWorks swApp, int total, string title = "Processing...")
        {
            _total = total > 0 ? total : 1;

            if (swApp == null) return;

            try
            {
                swApp.GetUserProgressBar(out _prgBar);
                if (_prgBar != null)
                {
                    _prgBar.Start(0, _total, title ?? "Processing...");
                }
            }
            catch (Exception ex)
            {
                // Non-fatal: progress bar is a convenience, not required
                ErrorHandler.LogInfo($"[PERF] SwProgressBar: Could not create native progress bar: {ex.Message}");
                _prgBar = null;
            }
        }

        /// <summary>
        /// Update the progress position and optionally the title text.
        /// Sets UserCanceled=true if the user pressed Escape.
        /// </summary>
        /// <param name="position">Current step (0-based, clamped to total)</param>
        /// <param name="title">Optional new title text. Pass null to keep the current title.</param>
        public void Update(int position, string title = null)
        {
            if (_prgBar == null) return;

            try
            {
                if (title != null)
                {
                    _prgBar.UpdateTitle(title);
                }

                int result = _prgBar.UpdateProgress(position);
                if (result == (int)swUpdateProgressError_e.swUpdateProgressError_UserCancel)
                {
                    UserCanceled = true;
                }
            }
            catch
            {
                // Swallow -- progress bar may have been destroyed externally
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                _prgBar?.End();
            }
            catch { }

            _prgBar = null;
        }
    }
}
