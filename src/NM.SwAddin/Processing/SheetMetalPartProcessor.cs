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
            return SolidWorksApiWrapper.HasSheetMetalFeature(model);
        }

        /// <summary>
        /// Process sheet metal part: convert if needed, flatten, extract metrics.
        /// </summary>
        public ProcessingResult Process(IModelDoc2 model, ModelInfo info, ProcessingOptions options)
        {
            const string proc = "SheetMetalPartProcessor.Process";
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (model == null || info == null)
                {
                    return ProcessingResult.Fail("Invalid inputs", ProblemPartManager.ProblemCategory.Fatal);
                }

                var processor = new SimpleSheetMetalProcessor(_swApp);
                bool success = processor.ConvertToSheetMetalAndOptionallyFlatten(info, model, true, options);

                if (success)
                {
                    // Mark as sheet metal in custom properties
                    info.CustomProperties.SetPropertyValue("PartType", "SheetMetal", CustomPropertyType.Text);
                    return ProcessingResult.Ok(Type.ToString());
                }
                else
                {
                    string reason = info.ProblemDescription ?? "Sheet metal processing failed";
                    return ProcessingResult.Fail(reason, ProblemPartManager.ProblemCategory.ProcessingError);
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Sheet metal processing exception", ex, ErrorHandler.LogLevel.Error);
                return ProcessingResult.Fail("Exception: " + ex.Message, ProblemPartManager.ProblemCategory.ProcessingError);
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }
    }
}
