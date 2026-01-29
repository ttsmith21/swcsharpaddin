using System;
using NM.Core;
using NM.Core.DataModel;
using NM.Core.Manufacturing;
using NM.Core.Processing;
using NM.Core.Tubes;
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

                // Use ProcessorFactory to detect part type and route to correct processor
                var factory = new ProcessorFactory(swApp);
                var processor = factory.DetectFor(doc);

                // Processor will be SheetMetal, Tube, or Generic based on part characteristics
                var pres = processor.Process(doc, info, options ?? new ProcessingOptions());
                if (!pres.Success)
                {
                    return new RunResult { Success = false, Message = pres.ErrorMessage ?? "Processing failed" };
                }

                // Batched custom property writeback (only if SaveChanges is enabled)
                var opts = options ?? new ProcessingOptions();
                if (opts.SaveChanges && info.CustomProperties.IsDirty)
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

                // Basic geometry check - get all solid bodies
                var partDoc = doc as IPartDoc;
                if (partDoc == null)
                {
                    pd.Status = ProcessingStatus.Failed;
                    pd.FailureReason = "Could not cast to IPartDoc";
                    return pd;
                }

                var bodiesObj = partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
                if (bodiesObj == null)
                {
                    pd.Status = ProcessingStatus.Failed;
                    pd.FailureReason = "No solid body detected";
                    return pd;
                }

                var bodies = (object[])bodiesObj;
                if (bodies.Length == 0)
                {
                    pd.Status = ProcessingStatus.Failed;
                    pd.FailureReason = "No solid body detected";
                    return pd;
                }

                // Multi-body check - FAIL validation for multi-body parts
                if (bodies.Length > 1)
                {
                    pd.Status = ProcessingStatus.Failed;
                    pd.FailureReason = $"Multi-body part ({bodies.Length} bodies)";
                    return pd;
                }

                var body = (IBody2)bodies[0];

                // Build core wrappers
                var info = new NM.Core.ModelInfo();
                info.Initialize(pathOrTitle, cfg);
                var swModel = new SolidWorksModel(info, swApp);
                swModel.Attach(doc, cfg);

                // VBA Logic: Try sheet metal FIRST, only fall back to tube if sheet metal fails
                ErrorHandler.DebugLog("[SMDBG] === CLASSIFICATION START ===");
                ErrorHandler.DebugLog($"[SMDBG] File: {pathOrTitle}");

                var factory = new ProcessorFactory(swApp);
                var sheetProcessor = factory.Get(ProcessorType.SheetMetal);
                ErrorHandler.DebugLog($"[SMDBG] SheetProcessor obtained: {(sheetProcessor != null ? sheetProcessor.GetType().Name : "NULL")}");

                // Step 1: Try sheet metal processing first (like VBA's SMInsertBends)
                bool isSheetMetal = false;
                if (sheetProcessor != null)
                {
                    ErrorHandler.DebugLog("[SMDBG] Step 1: Calling sheetProcessor.Process()...");
                    var sheetResult = sheetProcessor.Process(doc, info, options ?? new ProcessingOptions());
                    ErrorHandler.DebugLog($"[SMDBG] Step 1 Result: Success={sheetResult.Success}, Error={sheetResult.ErrorMessage ?? "none"}");
                    if (sheetResult.Success)
                    {
                        isSheetMetal = true;
                        pd.Classification = PartType.SheetMetal;
                        ErrorHandler.DebugLog("[SMDBG] Step 1: SHEET METAL DETECTED - Classification set to SheetMetal");
                    }
                    else
                    {
                        ErrorHandler.DebugLog("[SMDBG] Step 1: Sheet metal processing FAILED - will try tube");
                    }
                }
                else
                {
                    ErrorHandler.DebugLog("[SMDBG] Step 1: SKIPPED - sheetProcessor is NULL");
                }

                // Step 2: Only try tube detection if sheet metal failed (like VBA)
                ErrorHandler.DebugLog($"[SMDBG] Step 2: isSheetMetal={isSheetMetal}, will try tube={!isSheetMetal}");
                if (!isSheetMetal)
                {
                    try
                    {
                        ErrorHandler.DebugLog("[SMDBG] Step 2: Calling SimpleTubeProcessor.TryGetGeometry...");
                        var tube = new SimpleTubeProcessor(swApp).TryGetGeometry(doc);
                        if (tube != null)
                        {
                            ErrorHandler.DebugLog($"[SMDBG] Step 2: Tube result: Shape={tube.ShapeName}, Wall={tube.WallThickness:F3}, Length={tube.Length:F3}");
                        }
                        else
                        {
                            ErrorHandler.DebugLog("[SMDBG] Step 2: Tube detection returned null");
                        }

                        if (tube != null && tube.Shape != TubeShape.None && tube.WallThickness > 0 && tube.Length > 0)
                        {
                            const double IN_TO_M = 0.0254;
                            pd.Tube.IsTube = true;
                            pd.Classification = PartType.Tube;
                            pd.Tube.OD_m = tube.OuterDiameter * IN_TO_M;
                            pd.Tube.Wall_m = tube.WallThickness * IN_TO_M;
                            pd.Tube.ID_m = Math.Max(0.0, (tube.OuterDiameter - 2 * tube.WallThickness) * IN_TO_M);
                            pd.Tube.Length_m = tube.Length * IN_TO_M;
                            pd.Tube.TubeShape = tube.ShapeName; // Round, Square, Rectangle, Angle, Channel
                            pd.Tube.CrossSection = tube.CrossSection;
                            pd.Tube.CutLength_m = tube.CutLength * IN_TO_M;
                            pd.Tube.NumberOfHoles = tube.NumberOfHoles;
                            ErrorHandler.DebugLog($"[SMDBG] Step 2: TUBE DETECTED - Classification set to Tube ({tube.ShapeName})");

                            // Resolve pipe schedule (NPS and schedule code)
                            var pipeService = new PipeScheduleService();
                            string materialCategory = info.CustomProperties.MaterialCategory;
                            if (pipeService.TryResolveByOdAndWall(tube.OuterDiameter, tube.WallThickness, materialCategory, out string npsText, out string scheduleCode))
                            {
                                pd.Tube.NpsText = npsText;
                                pd.Tube.ScheduleCode = scheduleCode;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorHandler.DebugLog($"[SMDBG] Step 2: Tube detection EXCEPTION: {ex.Message}");
                    }
                }

                // Step 3: If neither sheet metal nor tube, use generic processor
                IPartProcessor processor;
                ProcessingResult pres;
                if (isSheetMetal)
                {
                    // Already processed as sheet metal above
                    processor = sheetProcessor;
                    pres = ProcessingResult.Ok("SheetMetal");
                }
                else if (pd.Classification == PartType.Tube)
                {
                    processor = factory.Get(ProcessorType.Tube);
                    pres = processor.Process(doc, info, options ?? new ProcessingOptions());
                }
                else
                {
                    processor = factory.Get(ProcessorType.Generic);
                    pres = processor.Process(doc, info, options ?? new ProcessingOptions());
                }
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

                // Read sheet metal data from custom properties if available
                if (info.CustomProperties.BendCount > 0)
                {
                    pd.Sheet.IsSheetMetal = true;
                    pd.Sheet.BendCount = info.CustomProperties.BendCount;
                }

                // ====== COST CALCULATIONS ======
                CalculateCosts(pd, info, options ?? new ProcessingOptions());

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

                // Only save if SaveChanges is enabled (default true, but QA disables it)
                var opts2 = options ?? new ProcessingOptions();
                if (opts2.SaveChanges && info.CustomProperties.IsDirty)
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

        /// <summary>
        /// Calculates all work center costs and populates PartData.Cost.
        /// Routes to tube or sheet metal cost calculations based on part type.
        /// </summary>
        private static void CalculateCosts(PartData pd, ModelInfo info, ProcessingOptions options)
        {
            const double KG_TO_LB = 2.20462;

            try
            {
                double rawWeightLb = pd.Mass_kg * KG_TO_LB;
                pd.Cost.MaterialWeight_lb = rawWeightLb;
                int quantity = Math.Max(1, options.Quantity > 0 ? options.Quantity : pd.QuoteQty);

                // Route to appropriate cost calculator based on part type
                if (pd.Classification == PartType.Tube && pd.Tube.IsTube)
                {
                    CalculateTubeCosts(pd, rawWeightLb, quantity);
                }
                else
                {
                    CalculateSheetMetalCosts(pd, info, rawWeightLb, quantity);
                }

                // Material Cost (applies to both tube and sheet metal)
                if (rawWeightLb > 0 && !string.IsNullOrWhiteSpace(pd.Material))
                {
                    var matInput = new MaterialCostCalculator.MaterialCostInput
                    {
                        WeightLb = rawWeightLb,
                        MaterialCode = pd.Material,
                        Quantity = quantity,
                        NestEfficiency = options.NestEfficiency > 0 ? options.NestEfficiency : 0.85
                    };
                    var matResult = MaterialCostCalculator.Calculate(matInput);
                    pd.Cost.MaterialCost = matResult.CostPerPiece;
                    pd.Cost.TotalMaterialCost = matResult.TotalMaterialCost;
                    pd.MaterialCostPerLB = matResult.CostPerLb;
                }

                // Total Cost Rollup
                double processingCost = pd.Cost.F115_Price + pd.Cost.F210_Price +
                                        pd.Cost.F140_Price + pd.Cost.F220_Price +
                                        pd.Cost.F325_Price;
                pd.Cost.TotalProcessingCost = processingCost;
                pd.Cost.TotalCost = pd.Cost.TotalMaterialCost + processingCost;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("CalculateCosts", "Cost calculation failed", ex, ErrorHandler.LogLevel.Warning);
            }
        }

        /// <summary>
        /// Calculates tube-specific work center costs using TubeWorkCenterRules.
        /// F325 (Roll Form), F140 (Press Brake if required), F210 (Deburr).
        /// </summary>
        private static void CalculateTubeCosts(PartData pd, double rawWeightLb, int quantity)
        {
            const double M_TO_IN = 39.3701;

            // Convert tube dimensions to inches
            double wallIn = pd.Tube.Wall_m * M_TO_IN;
            double lengthIn = pd.Tube.Length_m * M_TO_IN;

            // F325 Roll Form - always applies to tubes
            var f325Result = TubeWorkCenterRules.ComputeF325(rawWeightLb, wallIn);
            pd.Cost.F325_S_min = f325Result.SetupHours * 60.0;
            pd.Cost.F325_R_min = f325Result.RunHours * 60.0;
            pd.Cost.F325_Price = (f325Result.SetupHours + f325Result.RunHours * quantity) * CostConstants.F325_COST;

            // F140 Press Brake - only if F325 requires it (heavy tube with thick wall)
            if (f325Result.RequiresPressBrake)
            {
                var f140Result = TubeWorkCenterRules.ComputeF140(rawWeightLb, wallIn);
                pd.Cost.F140_S_min = f140Result.SetupHours * 60.0;
                pd.Cost.F140_R_min = f140Result.RunHours * 60.0;
                pd.Cost.F140_Price = (f140Result.SetupHours + f140Result.RunHours * quantity) * CostConstants.F140_COST;
            }

            // F210 Deburr - based on tube length
            if (lengthIn > 0)
            {
                var f210Result = TubeWorkCenterRules.ComputeF210(lengthIn);
                pd.Cost.F210_S_min = f210Result.SetupHours * 60.0;
                pd.Cost.F210_R_min = f210Result.RunHours * 60.0;
                pd.Cost.F210_Price = (f210Result.SetupHours + f210Result.RunHours * quantity) * CostConstants.F210_COST;
            }
        }

        /// <summary>
        /// Calculates sheet metal work center costs.
        /// F210 (Deburr), F140 (Press Brake), F220 (Tapping), F325 (Roll Forming if large radius).
        /// </summary>
        private static void CalculateSheetMetalCosts(PartData pd, ModelInfo info, double rawWeightLb, int quantity)
        {
            const double M_TO_IN = 39.3701;

            // F210 Deburr - based on cut perimeter
            if (pd.Sheet.TotalCutLength_m > 0)
            {
                double cutPerimeterIn = pd.Sheet.TotalCutLength_m * M_TO_IN;
                double f210Hours = F210Calculator.ComputeHours(cutPerimeterIn);
                pd.Cost.F210_R_min = f210Hours * 60.0;
                pd.Cost.F210_Price = F210Calculator.ComputeCost(cutPerimeterIn, quantity);
            }

            // F140 Press Brake - based on bend info
            if (pd.Sheet.BendCount > 0)
            {
                var bendInfo = new BendInfo
                {
                    Count = pd.Sheet.BendCount,
                    LongestBendIn = info.CustomProperties.LongestBendIn > 0
                        ? info.CustomProperties.LongestBendIn
                        : pd.BBoxWidth_m * M_TO_IN, // Estimate from bounding box
                    NeedsFlip = pd.Sheet.BendsBothDirections
                };

                var f140Result = F140Calculator.Compute(bendInfo, rawWeightLb, quantity);
                pd.Cost.F140_S_min = f140Result.SetupHours * 60.0;
                pd.Cost.F140_R_min = f140Result.RunHours * 60.0;
                pd.Cost.F140_Price = f140Result.Price(quantity);
            }

            // F220 Tapping - based on tapped hole count
            int tappedHoles = info.CustomProperties.TappedHoleCount;
            if (tappedHoles > 0)
            {
                var f220Input = new F220Input { Setups = 1, Holes = tappedHoles };
                var f220Result = F220Calculator.Compute(f220Input);
                pd.Cost.F220_S_min = f220Result.SetupHours * 60.0;
                pd.Cost.F220_R_min = f220Result.RunHours * 60.0;
                pd.Cost.F220_RN = tappedHoles;
                pd.Cost.F220_Price = (f220Result.SetupHours + f220Result.RunHours * quantity) * CostConstants.F220_COST;
            }

            // F325 Roll Forming - based on max bend radius (only if radius > 2 inches)
            double maxRadiusIn = info.CustomProperties.MaxBendRadiusIn;
            if (maxRadiusIn > 2.0)
            {
                var f325Calc = new F325Calculator();
                double arcLengthIn = info.CustomProperties.ArcLengthIn > 0
                    ? info.CustomProperties.ArcLengthIn
                    : maxRadiusIn * 3.14159; // Rough estimate: half circle
                var f325Result = f325Calc.CalculateRollForming(maxRadiusIn, arcLengthIn, quantity);
                if (f325Result.RequiresRollForming)
                {
                    pd.Cost.F325_S_min = f325Result.SetupHours * 60.0;
                    pd.Cost.F325_R_min = f325Result.RunHours * 60.0;
                    pd.Cost.F325_Price = f325Result.TotalCost;
                }
            }
        }
    }
}
