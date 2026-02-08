using System;
using NM.Core;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin
{
    /// <summary>
    /// Sketch creation and management operations for SolidWorks models.
    /// Extracted from SolidWorksApiWrapper for single-responsibility.
    /// </summary>
    public static class SwSketchHelper
    {
        /// <summary>
        /// Validates that a SketchManager reference is not null.
        /// </summary>
        public static bool ValidateSketchManager(ISketchManager swSkMgr, string procedureName)
        {
            if (swSkMgr == null)
            {
                ErrorHandler.HandleError(procedureName, "Invalid SketchManager reference");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Retrieves the SketchManager from a SolidWorks model.
        /// </summary>
        public static ISketchManager GetSketchManager(IModelDoc2 swModel)
        {
            const string procName = "GetSketchManager";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return null;
                return swModel.SketchManager;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Failed to get sketch manager", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Creates a line in the active sketch between two points.
        /// </summary>
        public static ISketchSegment CreateSketchLine(IModelDoc2 swModel, double x1, double y1, double x2, double y2)
        {
            const string procName = "CreateSketchLine";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return null;
                if (x1 == x2 && y1 == y2)
                {
                    ErrorHandler.HandleError(procName, "Zero-length line not allowed", null, ErrorHandler.LogLevel.Warning);
                    return null;
                }

                var swSkMgr = swModel.SketchManager;
                if (!ValidateSketchManager(swSkMgr, procName)) return null;

                var seg = swSkMgr.CreateLine(x1, y1, 0.0, x2, y2, 0.0) as ISketchSegment;
                if (seg == null)
                {
                    ErrorHandler.HandleError(procName, "Failed to create line", null, ErrorHandler.LogLevel.Warning);
                }
                return seg;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Exception creating line", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Creates a new sketch on the specified plane in a SolidWorks model.
        /// </summary>
        public static bool StartSketchOnPlane(IModelDoc2 swModel, string planeName)
        {
            const string procName = "StartSketchOnPlane";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return false;
                if (!SolidWorksApiWrapper.ValidateString(planeName, procName, "plane name")) return false;

                if (!SolidWorksApiWrapper.ClearSelection(swModel)) return false;

                if (!SolidWorksApiWrapper.SelectByName(swModel, planeName, "PLANE", 0, 0, 0))
                {
                    ErrorHandler.HandleError(procName, $"Failed to select plane: {planeName}", null, ErrorHandler.LogLevel.Warning);
                    return false;
                }

                var swSkMgr = swModel.SketchManager;
                swSkMgr.InsertSketch(true);
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception creating sketch on: {planeName}", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Exits sketch editing mode in a SolidWorks model.
        /// </summary>
        public static bool EndSketch(IModelDoc2 swModel)
        {
            const string procName = "EndSketch";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return false;

                var swSkMgr = swModel.SketchManager;
                if (!ValidateSketchManager(swSkMgr, procName)) return false;

                swSkMgr.InsertSketch(true);
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Failed to exit sketch mode", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }
    }
}
