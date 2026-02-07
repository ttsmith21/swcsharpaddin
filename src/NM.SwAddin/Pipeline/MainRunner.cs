using System;
using NM.Core;
using NM.Core.DataModel;
using NM.Core.Manufacturing;
using NM.Core.Manufacturing.Laser;
using NM.Core.Materials;
using NM.Core.Processing;
using NM.Core.Tubes;
using NM.SwAddin.Geometry;
using NM.SwAddin.Manufacturing;
using NM.SwAddin.Processing;
using NM.SwAddin.SheetMetal;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using static NM.Core.Constants.UnitConversions;

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
            PerformanceTracker.Instance.StartTimer("RunSinglePartData");
            try
            {
                var pd = new PartData();

                // ====== VALIDATION ======
                PerformanceTracker.Instance.StartTimer("Validation");
                if (swApp == null || doc == null)
                {
                    PerformanceTracker.Instance.StopTimer("Validation");
                    pd.Status = ProcessingStatus.Failed;
                    pd.FailureReason = "Null app or document";
                    return pd;
                }

                var type = (swDocumentTypes_e)doc.GetType();
                if (type != swDocumentTypes_e.swDocPART)
                {
                    PerformanceTracker.Instance.StopTimer("Validation");
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
                    PerformanceTracker.Instance.StopTimer("Validation");
                    pd.Status = ProcessingStatus.Failed;
                    pd.FailureReason = "Could not cast to IPartDoc";
                    return pd;
                }

                var bodiesObj = partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
                if (bodiesObj == null)
                {
                    PerformanceTracker.Instance.StopTimer("Validation");
                    pd.Status = ProcessingStatus.Failed;
                    pd.FailureReason = "No solid body detected";
                    return pd;
                }

                var bodies = (object[])bodiesObj;
                if (bodies.Length == 0)
                {
                    PerformanceTracker.Instance.StopTimer("Validation");
                    pd.Status = ProcessingStatus.Failed;
                    pd.FailureReason = "No solid body detected";
                    return pd;
                }

                // Multi-body check - FAIL validation for multi-body parts
                if (bodies.Length > 1)
                {
                    PerformanceTracker.Instance.StopTimer("Validation");
                    pd.Status = ProcessingStatus.Failed;
                    pd.FailureReason = $"Multi-body part ({bodies.Length} bodies)";
                    return pd;
                }

                var body = (IBody2)bodies[0];

                // Material check - FAIL validation if no material assigned
                if (string.IsNullOrWhiteSpace(pd.Material))
                {
                    PerformanceTracker.Instance.StopTimer("Validation");
                    pd.Status = ProcessingStatus.Failed;
                    pd.FailureReason = "No material assigned";
                    return pd;
                }
                PerformanceTracker.Instance.StopTimer("Validation");

                // Build core wrappers
                var info = new NM.Core.ModelInfo();
                info.Initialize(pathOrTitle, cfg);
                var swModel = new SolidWorksModel(info, swApp);
                swModel.Attach(doc, cfg);

                // ====== PURCHASED/OVERRIDE EARLY-OUT ======
                // If part has rbPartType=1 (set by user via PUR/MACH/CUST buttons or pre-existing),
                // skip the entire classification pipeline. These are known non-fabricated parts.
                string rbPartTypeEarly = SolidWorksApiWrapper.GetCustomPropertyValue(doc, "rbPartType");
                if (rbPartTypeEarly == "1")
                {
                    string rbSubVal = SolidWorksApiWrapper.GetCustomPropertyValue(doc, "rbPartTypeSub");
                    pd.Classification = PartType.Purchased;
                    pd.IsPurchased = true;
                    pd.Extra["rbPartType"] = "1";
                    if (!string.IsNullOrEmpty(rbSubVal)) pd.Extra["rbPartTypeSub"] = rbSubVal;

                    // Set work center based on sub-type
                    pd.Cost.OP20_WorkCenter = (rbSubVal == "2") ? "CUST" : "NPUR";
                    pd.Cost.OP20_S_min = 0;
                    pd.Cost.OP20_R_min = 0;

                    ErrorHandler.DebugLog($"[SMDBG] rbPartType=1 detected (sub={rbSubVal}) - skipping classification, work center={pd.Cost.OP20_WorkCenter}");

                    // Still collect mass for material costing
                    var purchMass = SolidWorksApiWrapper.GetModelMass(doc);
                    if (purchMass >= 0) pd.Mass_kg = purchMass;
                    pd.Material = SolidWorksApiWrapper.GetMaterialName(doc);
                    pd.Status = ProcessingStatus.Success;

                    // Skip to ERP property copy (bypass classification + processing)
                    goto ErpPropertyCopy;
                }

                // ====== CLASSIFICATION ======
                // VBA Logic: Try sheet metal FIRST, only fall back to tube if sheet metal fails
                PerformanceTracker.Instance.StartTimer("Classification");
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
                    PerformanceTracker.Instance.StartTimer("Classification_SheetMetal");
                    var sheetResult = sheetProcessor.Process(doc, info, options ?? new ProcessingOptions());
                    PerformanceTracker.Instance.StopTimer("Classification_SheetMetal");
                    ErrorHandler.DebugLog($"[SMDBG] Step 1 Result: Success={sheetResult.Success}, Error={sheetResult.ErrorMessage ?? "none"}");
                    if (sheetResult.Success)
                    {
                        isSheetMetal = true;
                        pd.Classification = PartType.SheetMetal;
                        ErrorHandler.DebugLog("[SMDBG] Step 1: SHEET METAL DETECTED - Classification set to SheetMetal");

                        // Validate bend allowance settings (logs warnings for non-table settings)
                        // Ported from VBA sheetmetal1.bas Process_CustomBendAllowance()
                        BendAllowanceValidator.ValidateAllFeatures(doc);
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
                        PerformanceTracker.Instance.StartTimer("Classification_Tube");
                        var tube = new SimpleTubeProcessor(swApp).TryGetGeometry(doc);
                        if (tube != null)
                        {
                            ErrorHandler.DebugLog($"[SMDBG] Step 2: Tube result: Shape={tube.ShapeName}, Wall={tube.WallThickness:F3}, Length={tube.Length:F3}");
                        }
                        else
                        {
                            ErrorHandler.DebugLog("[SMDBG] Step 2: Tube detection returned null");
                        }

                        // Tube detection validation:
                        // - Must have a minimum length (0.5" = 12.7mm) to avoid classifying machined blocks as tubes
                        // - Must have proper aspect ratio (length > 2x wall thickness) for true extrusions
                        const double MIN_TUBE_LENGTH_IN = 0.5;

                        // Round bars are solid cylinders (no inner face → wall=0) but still valid tube stock
                        bool isRoundBar = tube != null &&
                                          tube.Shape == TubeShape.Round &&
                                          tube.WallThickness == 0 &&
                                          tube.OuterDiameter > 0 &&
                                          tube.Length >= MIN_TUBE_LENGTH_IN;

                        bool isValidTube = (tube != null &&
                                           tube.Shape != TubeShape.None &&
                                           tube.WallThickness > 0 &&
                                           tube.Length >= MIN_TUBE_LENGTH_IN &&
                                           tube.Length > tube.WallThickness * 2) || isRoundBar;

                        if (!isValidTube && tube != null)
                        {
                            ErrorHandler.DebugLog($"[SMDBG] Step 2: Tube detection REJECTED - Length={tube.Length:F3}in (min={MIN_TUBE_LENGTH_IN}), Wall={tube.WallThickness:F3}in, AspectRatio={tube.Length / Math.Max(tube.WallThickness, 0.001):F1}");
                        }

                        if (isValidTube)
                        {
                            const double IN_TO_M = 0.0254;
                            pd.Tube.IsTube = true;
                            pd.Classification = PartType.Tube;
                            pd.Tube.OD_m = tube.OuterDiameter * IN_TO_M;
                            pd.Tube.Wall_m = tube.WallThickness * IN_TO_M;
                            pd.Tube.ID_m = tube.InnerDiameter > 0
                                ? tube.InnerDiameter * IN_TO_M
                                : Math.Max(0.0, (tube.OuterDiameter - 2 * tube.WallThickness) * IN_TO_M);
                            pd.Tube.Length_m = tube.Length * IN_TO_M;
                            pd.Tube.TubeShape = isRoundBar ? "Round Bar" : tube.ShapeName;
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
                    finally
                    {
                        PerformanceTracker.Instance.StopTimer("Classification_Tube");
                    }
                }

                // Step 3: If neither sheet metal nor tube, check for invalid geometry
                // Parts that attempted sheet metal but failed with no thickness are likely
                // invalid geometry (knife edges, degenerate faces) and should fail validation.
                bool attemptedSheetMetal = (sheetProcessor != null);
                bool sheetMetalFailed = attemptedSheetMetal && !isSheetMetal;
                bool isTube = (pd.Classification == PartType.Tube);

                if (sheetMetalFailed && !isTube)
                {
                    // Check if we can measure thickness - if not, this MIGHT be invalid geometry
                    double probeThickness = info.CustomProperties.Thickness;
                    if (probeThickness <= 0)
                    {
                        // Also check if part has very low sheet percentage (< 10%)
                        double sheetPercent = info.CustomProperties.SheetPercent;
                        if (sheetPercent < 0.10)
                        {
                            // Before failing, check if part has real mass.
                            // Solid parts (round bars, blocks) have no sheet thickness but are valid Generic parts.
                            // Only truly invalid geometry (knife edges, degenerate faces) has essentially no mass.
                            var massProps = doc.Extension?.CreateMassProperty() as IMassProperty;
                            double massKg = massProps?.Mass ?? 0.0;
                            const double MIN_MASS_KG = 0.001; // ~0.002 lb (1 gram) - only truly degenerate geometry is below this

                            if (massKg < MIN_MASS_KG)
                            {
                                ErrorHandler.DebugLog($"[SMDBG] Step 3: Part has no measurable thickness, low sheet% ({sheetPercent:P0}), and negligible mass ({massKg:F4} kg) - failing as invalid geometry");
                                PerformanceTracker.Instance.StopTimer("Classification");
                                pd.Status = ProcessingStatus.Failed;
                                pd.FailureReason = "Invalid geometry - no measurable thickness";
                                return pd;
                            }
                            else
                            {
                                ErrorHandler.DebugLog($"[SMDBG] Step 3: Part has real mass ({massKg:F4} kg) - treating as valid Generic part, not invalid geometry");
                            }
                        }
                    }
                }

                PerformanceTracker.Instance.StopTimer("Classification");

                // ====== PROCESSING ======
                // Proceed with appropriate processor
                PerformanceTracker.Instance.StartTimer("Processing");
                IPartProcessor processor;
                ProcessingResult pres;
                if (isSheetMetal)
                {
                    // Already processed as sheet metal above
                    processor = sheetProcessor;
                    pres = ProcessingResult.Ok("SheetMetal");
                }
                else if (isTube)
                {
                    processor = factory.Get(ProcessorType.Tube);
                    pres = processor.Process(doc, info, options ?? new ProcessingOptions());
                }
                else
                {
                    processor = factory.Get(ProcessorType.Generic);
                    pres = processor.Process(doc, info, options ?? new ProcessingOptions());
                }
                PerformanceTracker.Instance.StopTimer("Processing");
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

                // Bounding box (inches -> meters)
                try
                {
                    var bboxExtractor = new BoundingBoxExtractor();
                    var blankSize = bboxExtractor.GetBlankSize(doc);
                    if (blankSize.length > 0)
                    {
                        pd.BBoxLength_m = blankSize.length * InchesToMeters;
                        pd.BBoxWidth_m = blankSize.width * InchesToMeters;
                    }
                }
                catch (Exception bboxEx)
                {
                    ErrorHandler.HandleError("MainRunner", "BBox extraction failed", bboxEx, ErrorHandler.LogLevel.Warning);
                }

                // Thickness (inches) to meters if present in cache
                var thicknessIn = info.CustomProperties.Thickness; // inches
                if (thicknessIn > 0)
                {
                    pd.Thickness_m = thicknessIn * NM.Core.Configuration.Materials.InchesToMeters;
                }

                // Sheet percent if written by processors
                pd.SheetPercent = info.CustomProperties.SheetPercent;

                // Extract bend data directly from the model feature tree
                // (custom properties may not have BendCount after InsertBends2)
                // Note: countSuppressed=true because the model is flattened at this point,
                // so bend features are suppressed but still valid.
                if (isSheetMetal)
                {
                    var bends = BendAnalyzer.AnalyzeBends(doc, countSuppressed: true);
                    if (bends != null && bends.Count > 0)
                    {
                        pd.Sheet.IsSheetMetal = true;
                        pd.Sheet.BendCount = bends.Count;
                        pd.Sheet.BendsBothDirections = bends.NeedsFlip;
                        // Store bend geometry for F325/F140 calculations
                        if (bends.MaxRadiusIn > 0)
                            pd.Extra["MaxBendRadiusIn"] = bends.MaxRadiusIn.ToString("F4", System.Globalization.CultureInfo.InvariantCulture);
                        if (bends.LongestBendIn > 0)
                            pd.Extra["LongestBendIn"] = bends.LongestBendIn.ToString("F4", System.Globalization.CultureInfo.InvariantCulture);
                    }
                }
                else if (info.CustomProperties.BendCount > 0)
                {
                    pd.Sheet.IsSheetMetal = true;
                    pd.Sheet.BendCount = info.CustomProperties.BendCount;
                }

                // Extract flat pattern metrics (cut lengths, pierce count) for sheet metal
                if (isSheetMetal)
                {
                    try
                    {
                        var faceAnalyzer = new FaceAnalyzer();
                        var flatFace = faceAnalyzer.GetProcessingFace(doc);
                        if (flatFace != null)
                        {
                            var cutMetrics = FlatPatternAnalyzer.Extract(doc, flatFace);
                            if (cutMetrics != null && cutMetrics.TotalCutLengthIn > 0)
                            {
                                pd.Sheet.TotalCutLength_m = cutMetrics.TotalCutLengthIn * InchesToMeters;
                                pd.Extra["CutMetrics_PerimeterIn"] = cutMetrics.PerimeterLengthIn.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                pd.Extra["CutMetrics_InternalIn"] = cutMetrics.InternalCutLengthIn.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                pd.Extra["CutMetrics_PierceCount"] = cutMetrics.PierceCount.ToString();
                                pd.Extra["CutMetrics_HoleCount"] = cutMetrics.HoleCount.ToString();
                            }
                        }
                    }
                    catch (Exception flatEx)
                    {
                        ErrorHandler.HandleError("MainRunner", "FlatPattern extraction failed", flatEx, ErrorHandler.LogLevel.Warning);
                    }
                }

                ErpPropertyCopy:
                // ====== COPY ERP-RELEVANT CUSTOM PROPERTIES ======
                // These properties are used by ErpExportDataBuilder for VBA-parity export
                string[] erpProps = { "rbPartType", "rbPartTypeSub", "OS_WC", "OS_WC_A",
                                      "CustPartNumber", "PurchasedPartNumber", "Print",
                                      "Description", "Revision", "OP20" };
                foreach (var prop in erpProps)
                {
                    var val = info.CustomProperties.GetPropertyValue(prop);
                    if (val != null && !string.IsNullOrEmpty(val.ToString()))
                        pd.Extra[prop] = val.ToString();
                }

                // ====== DESCRIPTION GENERATION ======
                var description = DescriptionGenerator.Generate(pd);
                if (!string.IsNullOrEmpty(description))
                    pd.Extra["Description"] = description;

                // ====== OPTIMATERIAL RESOLUTION ======
                // Use static service as fallback when Excel data unavailable
                if (string.IsNullOrEmpty(pd.OptiMaterial))
                {
                    var optiCode = StaticOptiMaterialService.Resolve(pd);
                    if (!string.IsNullOrEmpty(optiCode))
                        pd.OptiMaterial = optiCode;
                }

                // ====== RAW WEIGHT CALCULATION ======
                // Use ManufacturingCalculator with custom properties (rbWeightCalc, NestEfficiency, Length, Width)
                // to compute raw blank weight with nest efficiency and thickness scrap multiplier
                try
                {
                    var partMetrics = MetricsExtractor.FromModel(doc, info);
                    // Override with our already-computed values (more reliable than re-reading)
                    if (pd.Mass_kg > 0) partMetrics.MassKg = pd.Mass_kg;
                    if (pd.Thickness_m > 0) partMetrics.ThicknessIn = pd.Thickness_m * MetersToInches;
                    if (!string.IsNullOrEmpty(pd.Material)) partMetrics.MaterialCode = pd.Material;

                    var calcResult = ManufacturingCalculator.Compute(partMetrics, new CalcOptions());
                    if (calcResult.RawWeightLb > 0)
                    {
                        pd.Cost.MaterialWeight_lb = calcResult.RawWeightLb;
                        pd.SheetPercent = calcResult.SheetPercent;
                    }
                }
                catch (Exception mfgEx)
                {
                    ErrorHandler.HandleError("MainRunner", "Raw weight calculation failed", mfgEx, ErrorHandler.LogLevel.Warning);
                }

                // ====== COST CALCULATIONS ======
                PerformanceTracker.Instance.StartTimer("CostCalculation");
                CalculateCosts(pd, info, options ?? new ProcessingOptions());
                PerformanceTracker.Instance.StopTimer("CostCalculation");

                // ====== PROPERTY WRITE ======
                // Map DTO -> properties and batch save
                PerformanceTracker.Instance.StartTimer("PropertyWrite");
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
                        PerformanceTracker.Instance.StopTimer("PropertyWrite");
                        pd.Status = ProcessingStatus.Failed;
                        pd.FailureReason = "Property writeback failed";
                        return pd;
                    }
                }
                PerformanceTracker.Instance.StopTimer("PropertyWrite");

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
                PerformanceTracker.Instance.StopTimer("RunSinglePartData");
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Calculates all work center costs and populates PartData.Cost.
        /// Routes to tube or sheet metal cost calculations based on part type.
        /// </summary>
        private static void CalculateCosts(PartData pd, ModelInfo info, ProcessingOptions options)
        {
            try
            {
                // Use ManufacturingCalculator-computed rawWeight if available, else fall back to finished mass
                double rawWeightLb = pd.Cost.MaterialWeight_lb > 0
                    ? pd.Cost.MaterialWeight_lb
                    : pd.Mass_kg * KgToLbs;
                int quantity = Math.Max(1, options.Quantity > 0 ? options.Quantity : pd.QuoteQty);

                // Check for purchased/customer-supplied parts (VBA: rbPartType=1)
                string rbPartType;
                if (pd.Extra.TryGetValue("rbPartType", out rbPartType) && rbPartType == "1")
                {
                    string rbPartTypeSub;
                    pd.Extra.TryGetValue("rbPartTypeSub", out rbPartTypeSub);
                    pd.Cost.OP20_WorkCenter = (rbPartTypeSub == "2") ? "CUST" : "NPUR";
                    pd.Cost.OP20_S_min = 0;
                    pd.Cost.OP20_R_min = 0;
                    // No processing costs for purchased parts - skip to material cost
                }
                // Route to appropriate cost calculator based on part type
                else if (pd.Classification == PartType.Tube && pd.Tube.IsTube)
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
            // Convert tube dimensions to inches
            double wallIn = pd.Tube.Wall_m * MetersToInches;
            double lengthIn = pd.Tube.Length_m * MetersToInches;
            double odIn = pd.Tube.OD_m * MetersToInches;

            // OP20 Routing: Work center assignment based on tube geometry
            // Solid bar: no wall thickness, or ID same as OD (wall=0 causes ID=OD in calculation)
            bool isSolidBar = pd.Tube.Wall_m < 0.001 || pd.Tube.ID_m < 0.001;
            if (isSolidBar)
            {
                pd.Cost.OP20_WorkCenter = "F300";
                // F300 saw: setup is fixed 0.05 hrs, run = ((diameter * 90) + 15) / 3600 hrs
                // Formula from VBA SP.bas RoundBar() function
                pd.Cost.OP20_S_min = 0.05 * 60; // 3 minutes
                pd.Cost.OP20_R_min = ((odIn * 90.0) + 15.0) / 60.0; // convert seconds to minutes
            }
            else
            {
                string shape = pd.Tube.TubeShape ?? "Round";
                if (shape == "Round")
                {
                    // OD thresholds use small epsilon (0.05") to handle floating point
                    // from metric-to-imperial conversion (e.g., 6.0" stores as 6.000039")
                    if (odIn > 10.80)       { pd.Cost.OP20_WorkCenter = "N145"; pd.Cost.OP20_S_min = 0.25 * 60; }
                    else if (odIn > 10.05)  { pd.Cost.OP20_WorkCenter = "F110"; pd.Cost.OP20_S_min = 1.0 * 60; }
                    else if (odIn > 6.05)   { pd.Cost.OP20_WorkCenter = "F110"; pd.Cost.OP20_S_min = 0.5 * 60; }
                    else                    { pd.Cost.OP20_WorkCenter = "F110"; pd.Cost.OP20_S_min = 0.15 * 60; }
                }
                else if (shape == "Angle" || shape == "Channel")
                {
                    pd.Cost.OP20_WorkCenter = "F110";
                    pd.Cost.OP20_S_min = 0.25 * 60;
                }
                else // Rectangle, Square
                {
                    pd.Cost.OP20_WorkCenter = "F110";
                    pd.Cost.OP20_S_min = 0.15 * 60;
                }
                // Run time from Mazak ExternalStart library - cannot replicate, remains 0
            }

            // F325 Roll Form - always applies to tubes
            PerformanceTracker.Instance.StartTimer("Cost_F325_RollForm");
            var f325Result = TubeWorkCenterRules.ComputeF325(rawWeightLb, wallIn);
            pd.Cost.F325_S_min = f325Result.SetupHours * 60.0;
            pd.Cost.F325_R_min = f325Result.RunHours * 60.0;
            pd.Cost.F325_Price = (f325Result.SetupHours + f325Result.RunHours * quantity) * CostConstants.F325_COST;
            PerformanceTracker.Instance.StopTimer("Cost_F325_RollForm");

            // F140 Press Brake - only if F325 requires it (heavy tube with thick wall)
            if (f325Result.RequiresPressBrake)
            {
                PerformanceTracker.Instance.StartTimer("Cost_F140_Brake");
                var f140Result = TubeWorkCenterRules.ComputeF140(rawWeightLb, wallIn);
                pd.Cost.F140_S_min = f140Result.SetupHours * 60.0;
                pd.Cost.F140_R_min = f140Result.RunHours * 60.0;
                pd.Cost.F140_Price = (f140Result.SetupHours + f140Result.RunHours * quantity) * CostConstants.F140_COST;
                PerformanceTracker.Instance.StopTimer("Cost_F140_Brake");
            }

            // F210 Deburr - based on tube length
            if (lengthIn > 0)
            {
                PerformanceTracker.Instance.StartTimer("Cost_F210_Deburr");
                var f210Result = TubeWorkCenterRules.ComputeF210(lengthIn);
                pd.Cost.F210_S_min = f210Result.SetupHours * 60.0;
                pd.Cost.F210_R_min = f210Result.RunHours * 60.0;
                pd.Cost.F210_Price = (f210Result.SetupHours + f210Result.RunHours * quantity) * CostConstants.F210_COST;
                PerformanceTracker.Instance.StopTimer("Cost_F210_Deburr");
            }
        }

        /// <summary>
        /// Calculates sheet metal work center costs.
        /// F210 (Deburr), F140 (Press Brake), F220 (Tapping), F325 (Roll Forming if large radius).
        /// </summary>
        private static void CalculateSheetMetalCosts(PartData pd, ModelInfo info, double rawWeightLb, int quantity)
        {

            // F115 Laser Cutting - based on cut length and pierce count
            if (pd.Sheet.TotalCutLength_m > 0 && pd.Thickness_m > 0)
            {
                PerformanceTracker.Instance.StartTimer("Cost_F115_Laser");
                try
                {
                    int pierceCount = 0;
                    string pcStr;
                    if (pd.Extra.TryGetValue("CutMetrics_PierceCount", out pcStr))
                        int.TryParse(pcStr, out pierceCount);

                    var partMetrics = new PartMetrics
                    {
                        ApproxCutLengthIn = pd.Sheet.TotalCutLength_m * MetersToInches,
                        PierceCount = pierceCount,
                        ThicknessIn = pd.Thickness_m * MetersToInches,
                        MaterialCode = pd.Material ?? "304L",
                        MassKg = pd.Mass_kg
                    };
                    ILaserSpeedProvider speedProvider = new StaticLaserSpeedProvider();
                    var laserResult = LaserCalculator.Compute(partMetrics, speedProvider, isWaterjet: false, rawWeightLb: rawWeightLb);
                    pd.Cost.OP20_S_min = laserResult.SetupHours * 60.0;
                    pd.Cost.OP20_R_min = laserResult.RunHours * 60.0;
                    pd.Cost.OP20_WorkCenter = "F115";
                    pd.Cost.F115_Price = laserResult.Cost;
                }
                catch (Exception laserEx)
                {
                    ErrorHandler.HandleError("MainRunner", "F115 laser calc failed", laserEx, ErrorHandler.LogLevel.Warning);
                }
                finally
                {
                    PerformanceTracker.Instance.StopTimer("Cost_F115_Laser");
                }
            }

            // F210 Deburr - based on cut perimeter
            if (pd.Sheet.TotalCutLength_m > 0)
            {
                PerformanceTracker.Instance.StartTimer("Cost_F210_Deburr");
                double cutPerimeterIn = pd.Sheet.TotalCutLength_m * MetersToInches;
                double f210Hours = F210Calculator.ComputeHours(cutPerimeterIn);
                pd.Cost.F210_R_min = f210Hours * 60.0;
                pd.Cost.F210_Price = F210Calculator.ComputeCost(cutPerimeterIn, quantity);
                PerformanceTracker.Instance.StopTimer("Cost_F210_Deburr");
            }

            // F140 Press Brake - based on bend info
            if (pd.Sheet.BendCount > 0)
            {
                PerformanceTracker.Instance.StartTimer("Cost_F140_Brake");
                // Longest bend line from BendAnalyzer (extracted from "Bend-Lines" sketch)
                double longestBendIn = 0.0;
                {
                    string extraBend;
                    if (pd.Extra.TryGetValue("LongestBendIn", out extraBend))
                        double.TryParse(extraBend, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out longestBendIn);
                }

                var bendInfo = new BendInfo
                {
                    Count = pd.Sheet.BendCount,
                    LongestBendIn = longestBendIn,
                    NeedsFlip = pd.Sheet.BendsBothDirections
                };

                // VBA: dblLength1 from LengthWidth(objFace) — longest flat face dimension for FindRate
                double partLengthIn = pd.BBoxLength_m * MetersToInches;

                var f140Result = F140Calculator.Compute(bendInfo, rawWeightLb, partLengthIn, quantity);
                pd.Cost.F140_S_min = f140Result.SetupHours * 60.0;
                pd.Cost.F140_R_min = f140Result.RunHours * 60.0;
                pd.Cost.F140_Price = f140Result.Price(quantity);
                PerformanceTracker.Instance.StopTimer("Cost_F140_Brake");
            }

            // F220 Tapping - based on tapped hole count
            int tappedHoles = info.CustomProperties.TappedHoleCount;
            if (tappedHoles > 0)
            {
                PerformanceTracker.Instance.StartTimer("Cost_F220_Tapping");
                var f220Input = new F220Input { Setups = 1, Holes = tappedHoles };
                var f220Result = F220Calculator.Compute(f220Input);
                pd.Cost.F220_S_min = f220Result.SetupHours * 60.0;
                pd.Cost.F220_R_min = f220Result.RunHours * 60.0;
                pd.Cost.F220_RN = tappedHoles;
                pd.Cost.F220_Price = (f220Result.SetupHours + f220Result.RunHours * quantity) * CostConstants.F220_COST;
                PerformanceTracker.Instance.StopTimer("Cost_F220_Tapping");
            }

            // F325 Roll Forming - based on max bend radius (only if radius > 2 inches)
            double maxRadiusIn = info.CustomProperties.MaxBendRadiusIn;
            // Fallback: use BendAnalyzer result stored in Extra
            if (maxRadiusIn <= 0)
            {
                string extraRadius;
                if (pd.Extra.TryGetValue("MaxBendRadiusIn", out extraRadius))
                    double.TryParse(extraRadius, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out maxRadiusIn);
            }
            if (maxRadiusIn > 2.0)
            {
                PerformanceTracker.Instance.StartTimer("Cost_F325_RollForm");
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
                PerformanceTracker.Instance.StopTimer("Cost_F325_RollForm");
            }
        }
    }
}
