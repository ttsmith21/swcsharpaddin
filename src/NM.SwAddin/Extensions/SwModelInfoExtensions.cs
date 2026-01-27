using NM.Core.Models;
using NM.SwAddin.Services;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.Extensions
{
    /// <summary>
    /// Add-in side helpers for SwModelInfo that interact with SolidWorks COM via DocumentService.
    /// Keeps NM.Core free of COM references.
    /// </summary>
    public static class SwModelInfoExtensions
    {
        /// <summary>
        /// Open the model in SolidWorks if not already open. Returns true if open.
        /// </summary>
        public static bool OpenInSolidWorks(this SwModelInfo info, ISldWorks swApp, bool silent = true)
        {
            if (info == null || swApp == null) return false;
            if (info.ModelDoc != null) return true;
            var svc = new DocumentService(swApp);
            return svc.Open(info, silent);
        }

        /// <summary>
        /// Close the model in SolidWorks if open. Optionally save if dirty.
        /// </summary>
        public static bool CloseInSolidWorks(this SwModelInfo info, ISldWorks swApp, bool save = false)
        {
            if (info == null || swApp == null) return false;
            if (info.ModelDoc == null) return true;
            var svc = new DocumentService(swApp);
            return svc.Close(info, save);
        }
    }
}
