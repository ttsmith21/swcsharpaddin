using System;
using NM.Core;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin
{
    /// <summary>
    /// Mass, volume, and center-of-mass queries for SolidWorks models.
    /// Extracted from SolidWorksApiWrapper for single-responsibility.
    /// </summary>
    public static class SwMassPropertiesHelper
    {
        /// <summary>
        /// Retrieves mass, volume, and center of mass from a model.
        /// </summary>
        public static bool GetAllMassProperties(IModelDoc2 swModel, out double mass, out double volume, out double comX, out double comY, out double comZ)
        {
            const string procName = "GetAllMassProperties";
            mass = 0; volume = 0; comX = 0; comY = 0; comZ = 0;

            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return false;

                var swExt = swModel.Extension;
                var swMassProp = swExt.CreateMassProperty() as IMassProperty;
                if (swMassProp == null)
                {
                    ErrorHandler.HandleError(procName, "Failed to create mass property object");
                    return false;
                }

                mass = swMassProp.Mass;
                volume = swMassProp.Volume;
                var arrCOM = swMassProp.CenterOfMass as object[];
                if (arrCOM == null || arrCOM.Length != 3)
                {
                    ErrorHandler.HandleError(procName, "Invalid center of mass data", null, ErrorHandler.LogLevel.Warning);
                    return false;
                }

                comX = Convert.ToDouble(arrCOM[0]);
                comY = Convert.ToDouble(arrCOM[1]);
                comZ = Convert.ToDouble(arrCOM[2]);
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Error getting mass properties", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Gets current model volume using mass properties.
        /// </summary>
        public static double GetModelVolume(IModelDoc2 swModel)
        {
            const string procName = "GetModelVolume";
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return -1;

                var swMassProp = swModel.Extension.CreateMassProperty() as IMassProperty;
                if (swMassProp == null)
                {
                    ErrorHandler.HandleError(procName, "Failed to create mass property", null, ErrorHandler.LogLevel.Warning);
                    return -1;
                }

                return Math.Round(swMassProp.Volume, 6);
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Exception getting model volume", ex);
                return -1;
            }
        }

        /// <summary>
        /// Gets current model mass using mass properties.
        /// </summary>
        public static double GetModelMass(IModelDoc2 swModel)
        {
            const string procName = "GetModelMass";
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return -1;

                var swMassProp = swModel.Extension.CreateMassProperty() as IMassProperty;
                if (swMassProp == null)
                {
                    ErrorHandler.HandleError(procName, "Failed to create mass property", null, ErrorHandler.LogLevel.Warning);
                    return -1;
                }

                return Math.Round(swMassProp.Mass, 6);
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Exception getting model mass", ex);
                return -1;
            }
        }
    }
}
