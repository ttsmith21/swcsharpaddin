using System;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.Utils
{
    /// <summary>
    /// Temporarily sets SolidWorks application visibility, restoring the previous state on dispose.
    /// </summary>
    public sealed class VisibilityScope : IDisposable
    {
        private readonly ISldWorks _sw;
        private readonly bool _restore;
        private readonly bool _prev;

        public VisibilityScope(ISldWorks sw, bool visible)
        {
            _sw = sw;
            try
            {
                _prev = _sw != null && _sw.Visible;
                if (_sw != null) _sw.Visible = visible;
                _restore = true;
            }
            catch
            {
                _restore = false;
            }
        }

        public void Dispose()
        {
            if (!_restore) return;
            try { if (_sw != null) _sw.Visible = _prev; } catch { }
        }
    }
}
