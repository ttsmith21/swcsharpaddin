using System;
using NM.Core;
using NM.Core.Processing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Processing
{
    /// <summary>
    /// Part classification pipeline using VBA-style trial-and-validate approach.
    /// Sequence: Try Sheet Metal → Try Tube → Fallback to Other
    /// </summary>
    public sealed class ClassificationPipeline
    {
        private const string LogPrefix = "[CLASSIFY]";
        private readonly ISldWorks _swApp;

        // Minimum wall thickness for valid tube: 0.015" (0.381mm)
        private const double MIN_TUBE_WALL_IN = 0.015;

        public ClassificationPipeline(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Classifies a part model using trial-and-validate approach.
        /// </summary>
        /// <param name="model">The SolidWorks part model to classify.</param>
        /// <param name="info">Optional ModelInfo for property updates.</param>
        /// <param name="options">Optional processing options.</param>
        /// <returns>Classification result with type and extracted metrics.</returns>
        public ClassificationResult Classify(IModelDoc2 model, ModelInfo info = null, ProcessingOptions options = null)
        {
            ErrorHandler.PushCallStack("ClassificationPipeline.Classify");
            try
            {
                if (model == null)
                    return ClassificationResult.Failed("Model is null");

                if (model.GetType() != (int)swDocumentTypes_e.swDocPART)
                    return ClassificationResult.Failed("Not a part document");

                options = options ?? new ProcessingOptions();
                info = info ?? new ModelInfo();

                // Check if already has sheet metal features - fast path
                if (SolidWorksApiWrapper.HasSheetMetalFeature(model))
                {
                    ErrorHandler.DebugLog($"{LogPrefix} Already has sheet metal features - classifying as SheetMetal");
                    return new ClassificationResult(PartClassification.SheetMetal, "Existing sheet metal features");
                }

                // STEP 1: Try sheet metal conversion (with validation)
                ErrorHandler.DebugLog($"{LogPrefix} Step 1: Attempting sheet metal conversion...");
                var smProcessor = new SimpleSheetMetalProcessor(_swApp);
                bool smSuccess = smProcessor.ConvertToSheetMetalAndOptionallyFlatten(info, model, true, options);

                if (smSuccess && info.IsSheetMetal)
                {
                    ErrorHandler.DebugLog($"{LogPrefix} Sheet metal conversion succeeded");
                    return new ClassificationResult(PartClassification.SheetMetal, "Converted to sheet metal")
                    {
                        Thickness = info.ThicknessInMeters,
                        BendRadius = info.BendRadius,
                        KFactor = info.KFactor
                    };
                }

                // Sheet metal failed - ensure we're back to original state
                ErrorHandler.DebugLog($"{LogPrefix} Sheet metal conversion failed: {info.ProblemDescription}");

                // STEP 2: Try tube detection
                ErrorHandler.DebugLog($"{LogPrefix} Step 2: Attempting tube detection...");
                var tubeProcessor = new SimpleTubeProcessor(_swApp);
                var tubeGeom = tubeProcessor.TryGetGeometry(model);

                if (tubeGeom != null &&
                    tubeGeom.OuterDiameter > 0 &&
                    tubeGeom.WallThickness >= MIN_TUBE_WALL_IN &&
                    tubeGeom.Length > 0)
                {
                    ErrorHandler.DebugLog($"{LogPrefix} Tube detected: OD={tubeGeom.OuterDiameter:F3}in, Wall={tubeGeom.WallThickness:F3}in, L={tubeGeom.Length:F3}in");

                    // Update info if provided
                    if (info.CustomProperties != null)
                    {
                        info.CustomProperties.IsTube = true;
                    }

                    return new ClassificationResult(PartClassification.Tube, "Tube geometry detected")
                    {
                        TubeOD = tubeGeom.OuterDiameter,
                        TubeWall = tubeGeom.WallThickness,
                        TubeLength = tubeGeom.Length,
                        TubeAxis = tubeGeom.Axis
                    };
                }

                if (tubeGeom != null)
                {
                    ErrorHandler.DebugLog($"{LogPrefix} Tube geometry found but invalid: OD={tubeGeom.OuterDiameter:F3}, Wall={tubeGeom.WallThickness:F3} (min={MIN_TUBE_WALL_IN})");
                }
                else
                {
                    ErrorHandler.DebugLog($"{LogPrefix} No tube geometry detected");
                }

                // STEP 3: Fallback to Other
                ErrorHandler.DebugLog($"{LogPrefix} Step 3: Classifying as Other");
                return new ClassificationResult(PartClassification.Other, "Not sheet metal or tube");
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("ClassificationPipeline.Classify", "Classification failed", ex, ErrorHandler.LogLevel.Error);
                return ClassificationResult.Failed($"Exception: {ex.Message}");
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }
    }

    /// <summary>
    /// Part classification types.
    /// </summary>
    public enum PartClassification
    {
        Unknown,
        SheetMetal,
        Tube,
        Other,
        Failed
    }

    /// <summary>
    /// Result of part classification with extracted metrics.
    /// </summary>
    public sealed class ClassificationResult
    {
        public PartClassification Classification { get; }
        public string Reason { get; }
        public bool Success => Classification != PartClassification.Failed && Classification != PartClassification.Unknown;

        // Sheet metal properties
        public double Thickness { get; set; }
        public double BendRadius { get; set; }
        public double KFactor { get; set; }

        // Tube properties
        public double TubeOD { get; set; }
        public double TubeWall { get; set; }
        public double TubeLength { get; set; }
        public double[] TubeAxis { get; set; }

        public ClassificationResult(PartClassification classification, string reason)
        {
            Classification = classification;
            Reason = reason ?? string.Empty;
        }

        public static ClassificationResult Failed(string reason)
        {
            return new ClassificationResult(PartClassification.Failed, reason);
        }
    }
}
