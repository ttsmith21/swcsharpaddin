using System;
using NM.Core;
using NM.Core.DataModel;
using NM.Core.Processing;
using NM.SwAddin.Processing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Pipeline
{
    /// <summary>
    /// Minimal single-part orchestrator used by AutoWorkflow and FolderProcessor.
    /// Validates the model, runs a generic processor, and writes back custom properties.
    /// This can be extended to route to sheet metal or tube processors.
    /// </summary>
    public static class MainRunner
    {
        public sealed class RunResult
        {
            public bool Success { get; set; }
            public string Message { get; set; }
        }

        public static RunResult RunSinglePart(ISldWorks swApp, IModelDoc2 doc, ProcessingOptions options)
        {
            const string proc = nameof(MainRunner) + ".RunSinglePart";
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (swApp == null || doc == null)
                {
                    return new RunResult { Success = false, Message = "Null app or document" };
                }

                var type = (swDocumentTypes_e)doc.GetType();
                if (type != swDocumentTypes_e.swDocPART)
                {
                    return new RunResult { Success = false, Message = "Active document is not a part" };
                }

                // Basic geometry sanity check
                var body = SolidWorksApiWrapper.GetMainBody(doc);
                if (body == null)
                {
                    return new RunResult { Success = false, Message = "No solid body detected" };
                }

                // Build model info and attach SolidWorks model wrapper
                var cfg = doc.ConfigurationManager?.ActiveConfiguration?.Name ?? string.Empty;
                var pathOrTitle = doc.GetPathName();
                if (string.IsNullOrWhiteSpace(pathOrTitle)) pathOrTitle = doc.GetTitle() ?? "UnsavedModel";

                var info = new NM.Core.ModelInfo();
                info.Initialize(pathOrTitle, cfg);

                var swModel = new SolidWorksModel(info, swApp);
                swModel.Attach(doc, cfg);

                // Run a simple processor now (sheet/tube routing can be added later)
                var processor = new GenericPartProcessor();
                if (!processor.CanProcess(doc))
                {
                    return new RunResult { Success = false, Message = "Unsupported document type" };
                }

                var pres = processor.Process(doc, info, options ?? new ProcessingOptions());
                if (!pres.Success)
                {
                    return new RunResult { Success = false, Message = pres.ErrorMessage ?? "Processing failed" };
                }

                // Batched custom property writeback
                if (info.CustomProperties.IsDirty)
                {
                    if (!swModel.SavePropertiesToSolidWorks())
                    {
                        return new RunResult { Success = false, Message = "Property writeback failed" };
                    }
                }

                return new RunResult { Success = true, Message = pres.ProcessorType + " OK" };
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Runner exception", ex, ErrorHandler.LogLevel.Error);
                return new RunResult { Success = false, Message = "Exception: " + ex.Message };
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Overload that returns a PartData DTO populated with basic identity and metrics.
        /// Also maps the DTO back into custom properties in one batched write.
        /// </summary>
        public static PartData RunSinglePartData(ISldWorks swApp, IModelDoc2 doc, ProcessingOptions options)
        {
            const string proc = nameof(MainRunner) + ".RunSinglePartData";
            ErrorHandler.PushCallStack(proc);
            try
            {
                var pd = new PartData();
                if (swApp == null || doc == null)
                {
                    pd.Status = ProcessingStatus.Failed;
                    pd.FailureReason = "Null app or document";
                    return pd;
                }

                var type = (swDocumentTypes_e)doc.GetType();
                if (type != swDocumentTypes_e.swDocPART)
                {
                    pd.Status = ProcessingStatus.Skipped;
                    pd.FailureReason = "Active document is not a part";
                    return pd;
                }

                var cfg = doc.ConfigurationManager?.ActiveConfiguration?.Name ?? string.Empty;
                var pathOrTitle = doc.GetPathName();
                if (string.IsNullOrWhiteSpace(pathOrTitle)) pathOrTitle = doc.GetTitle() ?? "UnsavedModel";

                pd.FilePath = doc.GetPathName();
                pd.PartName = doc.GetTitle();
                pd.Configuration = cfg;
                pd.Material = SolidWorksApiWrapper.GetMaterialName(doc);

                // Basic geometry check
                var body = SolidWorksApiWrapper.GetMainBody(doc);
                if (body == null)
                {
                    pd.Status = ProcessingStatus.Failed;
                    pd.FailureReason = "No solid body detected";
                    return pd;
                }

                // Build core wrappers
                var info = new NM.Core.ModelInfo();
                info.Initialize(pathOrTitle, cfg);
                var swModel = new SolidWorksModel(info, swApp);
                swModel.Attach(doc, cfg);

                // Detect tube geometry (inches ? meters)
                try
                {
                    var tube = new SimpleTubeProcessor().TryGetGeometry(doc);
                    if (tube != null && tube.OuterDiameter > 0 && tube.WallThickness > 0 && tube.Length > 0)
                    {
                        const double IN_TO_M = 0.0254;
                        pd.Tube.IsTube = true;
                        pd.Classification = PartType.Tube;
                        pd.Tube.OD_m = tube.OuterDiameter * IN_TO_M;
                        pd.Tube.Wall_m = tube.WallThickness * IN_TO_M;
                        pd.Tube.ID_m = Math.Max(0.0, (tube.OuterDiameter - 2 * tube.WallThickness) * IN_TO_M);
                        pd.Tube.Length_m = tube.Length * IN_TO_M;
                    }
                }
                catch { }

                // Run generic processor for baseline metrics/properties
                var processor = new GenericPartProcessor();
                if (!processor.CanProcess(doc))
                {
                    pd.Status = ProcessingStatus.Skipped;
                    pd.FailureReason = "Unsupported document type";
                    return pd;
                }

                var pres = processor.Process(doc, info, options ?? new ProcessingOptions());
                if (!pres.Success)
                {
                    pd.Status = ProcessingStatus.Failed;
                    pd.FailureReason = pres.ErrorMessage ?? "Processing failed";
                    return pd;
                }

                // Populate minimal DTO metrics
                pd.Status = ProcessingStatus.Success;
                if (pd.Classification == PartType.Unknown) pd.Classification = PartType.Generic;

                // Mass (kg)
                var mass = SolidWorksApiWrapper.GetModelMass(doc);
                if (mass >= 0) pd.Mass_kg = mass;

                // Thickness (inches) to meters if present in cache
                var thicknessIn = info.CustomProperties.Thickness; // inches
                if (thicknessIn > 0)
                {
                    pd.Thickness_m = thicknessIn * NM.Core.Configuration.Materials.InchesToMeters;
                }

                // Sheet percent if written by processors
                pd.SheetPercent = info.CustomProperties.SheetPercent;

                // Map DTO -> properties and batch save
                var mapped = PartDataPropertyMap.ToProperties(pd);
                foreach (var kv in mapped)
                {
                    // best-effort numeric detection for typing
                    if (double.TryParse(kv.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out _))
                        info.CustomProperties.SetPropertyValue(kv.Key, kv.Value, CustomPropertyType.Number);
                    else
                        info.CustomProperties.SetPropertyValue(kv.Key, kv.Value, CustomPropertyType.Text);
                }

                if (info.CustomProperties.IsDirty)
                {
                    if (!swModel.SavePropertiesToSolidWorks())
                    {
                        pd.Status = ProcessingStatus.Failed;
                        pd.FailureReason = "Property writeback failed";
                        return pd;
                    }
                }

                return pd;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Runner exception", ex, ErrorHandler.LogLevel.Error);
                var pd = new PartData { Status = ProcessingStatus.Failed, FailureReason = ex.Message };
                return pd;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }
    }
}
