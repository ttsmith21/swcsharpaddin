using System;
using System.IO;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.SheetMetal
{
    /// <summary>
    /// Validates bend allowance settings and logs warnings for non-standard configurations.
    /// Ported from VBA sheetmetal1.bas Process_CustomBendAllowance().
    /// </summary>
    public static class BendAllowanceValidator
    {
        private const double M_TO_IN = 39.37007874015748;

        /// <summary>
        /// Validates a sheet metal feature's bend allowance setting and logs a warning
        /// if it's not using a bend table (direct allowance, deduction, or K-factor).
        /// </summary>
        /// <param name="feat">The sheet metal feature to validate</param>
        /// <param name="model">The model document (for file name in log)</param>
        public static void ValidateAndLogWarnings(IFeature feat, IModelDoc2 model)
        {
            if (feat == null || model == null) return;

            try
            {
                var custBend = GetCustomBendAllowance(feat);
                if (custBend == null) return;

                var allowanceType = (swBendAllowanceTypes_e)custBend.Type;
                string fileName = Path.GetFileNameWithoutExtension(model.GetTitle() ?? "Unknown");
                string featName = feat.Name ?? "Unknown";

                switch (allowanceType)
                {
                    case swBendAllowanceTypes_e.swBendAllowanceDirect:
                        double allowanceIn = custBend.BendAllowance * M_TO_IN;
                        ErrorHandler.DebugLog($"[SM-WARN] {fileName}: {featName} uses Direct BendAllowance = {allowanceIn:F4} in");
                        break;

                    case swBendAllowanceTypes_e.swBendAllowanceDeduction:
                        double deductionIn = custBend.BendDeduction * M_TO_IN;
                        ErrorHandler.DebugLog($"[SM-WARN] {fileName}: {featName} uses BendDeduction = {deductionIn:F4} in");
                        break;

                    case swBendAllowanceTypes_e.swBendAllowanceKFactor:
                        double kFactor = custBend.KFactor;
                        ErrorHandler.DebugLog($"[SM-WARN] {fileName}: {featName} uses KFactor = {kFactor:F3}");
                        break;

                    case swBendAllowanceTypes_e.swBendAllowanceBendTable:
                        // Using bend table - this is the expected/preferred setting, no warning
                        break;

                    default:
                        // Unknown type - log for debugging
                        ErrorHandler.DebugLog($"[SM-WARN] {fileName}: {featName} has unknown bend allowance type = {(int)allowanceType}");
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[SM-WARN] Error validating bend allowance for {feat?.Name}: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets the CustomBendAllowance from various sheet metal feature types.
        /// Different feature types have different data interfaces.
        /// </summary>
        private static ICustomBendAllowance GetCustomBendAllowance(IFeature feat)
        {
            var def = feat.GetDefinition();
            if (def == null) return null;

            // Try each sheet metal feature type that has CustomBendAllowance
            // SW API returns object, explicit cast required
            if (def is ISheetMetalFeatureData smData)
                return (ICustomBendAllowance)smData.GetCustomBendAllowance();

            if (def is IBaseFlangeFeatureData bfData)
                return (ICustomBendAllowance)bfData.GetCustomBendAllowance();

            if (def is IEdgeFlangeFeatureData efData)
                return (ICustomBendAllowance)efData.GetCustomBendAllowance();

            if (def is IHemFeatureData hemData)
                return (ICustomBendAllowance)hemData.GetCustomBendAllowance();

            if (def is IJogFeatureData jogData)
                return (ICustomBendAllowance)jogData.GetCustomBendAllowance();

            if (def is IOneBendFeatureData obData)
                return (ICustomBendAllowance)obData.GetCustomBendAllowance();

            if (def is ISketchedBendFeatureData sbData)
                return (ICustomBendAllowance)sbData.GetCustomBendAllowance();

            if (def is IBendsFeatureData bendsData)
                return (ICustomBendAllowance)bendsData.GetCustomBendAllowance();

            return null;
        }

        /// <summary>
        /// Traverses all features in a model and validates bend allowance settings.
        /// Call this after sheet metal classification succeeds.
        /// </summary>
        /// <param name="model">The model document to validate</param>
        public static void ValidateAllFeatures(IModelDoc2 model)
        {
            if (model == null) return;

            var feat = (IFeature)model.FirstFeature();
            while (feat != null)
            {
                string typeName = feat.GetTypeName2();

                // Check sheet metal features that can have custom bend allowance
                if (IsSheetMetalFeatureWithBendAllowance(typeName))
                {
                    ValidateAndLogWarnings(feat, model);
                }

                // Also check sub-features (OneBend, SketchBend are often sub-features)
                var subFeat = (IFeature)feat.GetFirstSubFeature();
                while (subFeat != null)
                {
                    string subTypeName = subFeat.GetTypeName2();
                    if (subTypeName == "OneBend" || subTypeName == "SketchBend")
                    {
                        ValidateAndLogWarnings(subFeat, model);
                    }
                    subFeat = (IFeature)subFeat.GetNextSubFeature();
                }

                feat = (IFeature)feat.GetNextFeature();
            }
        }

        /// <summary>
        /// Returns true if the feature type is a sheet metal feature that can have
        /// custom bend allowance settings.
        /// </summary>
        private static bool IsSheetMetalFeatureWithBendAllowance(string typeName)
        {
            switch (typeName)
            {
                case "SheetMetal":      // Main sheet metal feature
                case "SMBaseFlange":    // Base flange
                case "EdgeFlange":      // Edge flange
                case "Hem":             // Hem
                case "Jog":             // Jog
                case "ProcessBends":    // Process bends (InsertBends result)
                case "FlattenBends":    // Flatten bends
                case "SM3dBend":        // 3D bend (rare)
                    return true;
                default:
                    return false;
            }
        }
    }
}
