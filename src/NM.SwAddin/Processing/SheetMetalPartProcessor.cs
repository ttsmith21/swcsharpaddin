using System;
using NM.Core;
using NM.Core.ProblemParts;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Processing
{
    /// <summary>
    /// IPartProcessor adapter for SimpleSheetMetalProcessor.
    /// Handles parts that already have sheet metal features or can be converted.
    /// </summary>
    public sealed class SheetMetalPartProcessor : IPartProcessor
    {
        private readonly ISldWorks _swApp;

        public SheetMetalPartProcessor(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        public ProcessorType Type => ProcessorType.SheetMetal;

        /// <summary>
        /// Returns true if the model already has sheet metal features.
        /// Note: This does NOT check if a part CAN be converted - only if it already IS sheet metal.
        /// </summary>
        public bool CanProcess(IModelDoc2 model)
        {
            if (model == null) return false;
            if (model.GetType() != (int)swDocumentTypes_e.swDocPART) return false;

            // Check for existing sheet metal features
            return SwGeometryHelper.HasSheetMetalFeature(model);
        }

        /// <summary>
        /// Process sheet metal part: convert if needed, flatten, extract metrics.
        /// </summary>
        public ProcessingResult Process(IModelDoc2 model, ModelInfo info, ProcessingOptions options)
        {
            const string proc = "SheetMetalPartProcessor.Process";
            ErrorHandler.PushCallStack(proc);
            ErrorHandler.DebugLog("[SMDBG] >>> SheetMetalPartProcessor.Process() ENTER");
            try
            {
                if (model == null || info == null)
                {
                    ErrorHandler.DebugLog($"[SMDBG] FAIL: Invalid inputs (model={model != null}, info={info != null})");
                    return ProcessingResult.Fail("Invalid inputs", ProblemPartManager.ProblemCategory.Fatal);
                }

                // Check if already has sheet metal features
                bool hasExistingFeatures = SwGeometryHelper.HasSheetMetalFeature(model);
                ErrorHandler.DebugLog($"[SMDBG] HasSheetMetalFeature={hasExistingFeatures}");

                ErrorHandler.DebugLog("[SMDBG] Creating SimpleSheetMetalProcessor...");
                var processor = new SimpleSheetMetalProcessor(_swApp);

                ErrorHandler.DebugLog("[SMDBG] Calling ConvertToSheetMetalAndOptionallyFlatten()...");
                bool success = processor.ConvertToSheetMetalAndOptionallyFlatten(info, model, true, options);
                ErrorHandler.DebugLog($"[SMDBG] ConvertToSheetMetalAndOptionallyFlatten returned: {success}");

                if (success)
                {
                    ErrorHandler.DebugLog("[SMDBG] SUCCESS - marking as SheetMetal");
                    // Mark as sheet metal in custom properties
                    info.CustomProperties.SetPropertyValue("PartType", "SheetMetal", CustomPropertyType.Text);
                    return ProcessingResult.Ok(Type.ToString());
                }
                else
                {
                    string reason = info.ProblemDescription ?? "Sheet metal processing failed";
                    ErrorHandler.DebugLog($"[SMDBG] FAIL - reason: {reason}");
                    return ProcessingResult.Fail(reason, ProblemPartManager.ProblemCategory.ProcessingError);
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[SMDBG] EXCEPTION: {ex.Message}");
                ErrorHandler.HandleError(proc, "Sheet metal processing exception", ex, ErrorHandler.LogLevel.Error);
                return ProcessingResult.Fail("Exception: " + ex.Message, ProblemPartManager.ProblemCategory.ProcessingError);
            }
            finally
            {
                ErrorHandler.DebugLog("[SMDBG] <<< SheetMetalPartProcessor.Process() EXIT");
                ErrorHandler.PopCallStack();
            }
        }
    }
}
