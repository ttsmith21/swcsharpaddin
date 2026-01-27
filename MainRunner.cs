using System;
using NM.Core;
using NM.Core.DataModel;
using NM.Core.Processing;
using NM.SwAddin.Validation;
using SolidWorks.Interop.sldworks;
using NM.Core.Manufacturing;
using NM.SwAddin.Manufacturing;
using NM.Core.Manufacturing.Laser;
using NM.SwAddin.Data;
using System.Globalization;

namespace NM.SwAddin
{
    public static class MainRunner
    {
        public sealed class RunResult
        {
            public bool Success { get; set; }
            public string Message { get; set; }
            public bool WasAlreadySheet { get; set; }
            public string ProblemDescription { get; set; }
        }

        /// <summary>
        /// Runs single part processing and returns PartData for batch operations.
        /// </summary>
        public static PartData RunSinglePartData(ISldWorks swApp, IModelDoc2 doc, ProcessingOptions options)
        {
            var pd = new PartData
            {
                FilePath = doc?.GetPathName() ?? string.Empty,
                PartName = System.IO.Path.GetFileNameWithoutExtension(doc?.GetPathName() ?? doc?.GetTitle() ?? "Unknown"),
                Status = ProcessingStatus.Pending
            };

            try
            {
                var result = RunSinglePart(swApp, doc, options ?? new ProcessingOptions());
                if (result.Success)
                {
                    pd.Status = ProcessingStatus.Success;
                }
                else
                {
                    pd.Status = ProcessingStatus.Failed;
                    pd.FailureReason = result.ProblemDescription ?? result.Message;
                }
            }
            catch (Exception ex)
            {
                pd.Status = ProcessingStatus.Failed;
                pd.FailureReason = ex.Message;
            }

            return pd;
        }

        // Local helper: interpret a property value as "tube" indicator
        private static bool LooksLikeTubeLocal(string value)
        {
            try
            {
                var s = (value ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(s)) return false;
                var u = s.ToUpperInvariant();
                if (u == "YES" || u == "TRUE" || u == "1") return true;
                return u.Contains("TUBE") || u.Contains("TUBING") || u.Contains("PIPE");
            }
            catch { return false; }
        }

        private static bool ShouldRouteToTube(NM.Core.ModelInfo info)
        {
            try
            {
                var cp = info?.CustomProperties;
                string[] keys = { "Shape", "IsTube", "Profile", "Section", "Type", "CrossSection" };
                foreach (var k in keys)
                {
                    var v = cp?.GetPropertyValue(k)?.ToString();
                    if (string.IsNullOrWhiteSpace(v)) continue;
                    if (k.Equals("Shape", System.StringComparison.OrdinalIgnoreCase))
                    {
                        // Non-empty shape is sufficient
                        return true;
                    }
                    if (LooksLikeTubeLocal(v)) return true;
                }
            }
            catch { }
            return false;
        }

        public static RunResult RunSinglePart(ISldWorks swApp, IModelDoc2 doc, ProcessingOptions options)
        {
            const string proc = nameof(MainRunner) + "." + nameof(RunSinglePart);
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (swApp == null || doc == null)
                {
                    return new RunResult { Success = false, Message = "No active document.", ProblemDescription = "Null app/doc" };
                }

                // Preflight validation (single solid body, etc.)
                try
                {
                    var validator = new PartValidationAdapter();
                    var swInfo = new NM.Core.Models.SwModelInfo(doc.GetPathName() ?? (doc.GetTitle() ?? "UnsavedModel"));
                    var vr = validator.Validate(swInfo, doc);
                    if (!vr.Success)
                    {
                        return new RunResult { Success = false, Message = "Validation failed: " + (vr.Summary ?? "Unknown"), ProblemDescription = vr.Summary };
                    }
                }
                catch { }

                // Load basic props to detect tube vs sheet
                var info = new NM.Core.ModelInfo();
                var cfg = (doc.ConfigurationManager != null && doc.ConfigurationManager.ActiveConfiguration != null)
                    ? doc.ConfigurationManager.ActiveConfiguration.Name : string.Empty;
                var pathOrTitle = doc.GetPathName();
                if (string.IsNullOrWhiteSpace(pathOrTitle)) pathOrTitle = doc.GetTitle() ?? "UnsavedModel";
                info.Initialize(pathOrTitle, cfg);

                var swModel = new SolidWorksModel(info, swApp);
                swModel.Attach(doc, cfg);

                // Ensure extractor populates Shape/CrossSection/etc. before routing
                try { ExternalStartAdapter.TryRunExternalStart(swApp, doc); } catch { }

                try { swModel.LoadPropertiesFromSolidWorks(); } catch { }

                bool treatAsTube = ShouldRouteToTube(info);
                ErrorHandler.DebugLog($"[Router] TubeRouting={treatAsTube}");

                if (treatAsTube)
                {
                    return RunTube(swApp, doc, options);
                }
                else
                {
                    return Run(swApp, doc, options);
                }
            }
            catch (System.Exception ex)
            {
                ErrorHandler.HandleError(proc, "Runner exception", ex);
                return new RunResult { Success = false, Message = "Exception: " + ex.Message };
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        public static RunResult Run(ISldWorks swApp, IModelDoc2 doc, ProcessingOptions options)
        {
            const string proc = nameof(MainRunner) + ".Run";
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (swApp == null || doc == null)
                {
                    return new RunResult { Success = false, Message = "No active document.", ProblemDescription = "Null app/doc" };
                }

                try
                {
                    var validator = new PartValidationAdapter();
                    var swInfo = new NM.Core.Models.SwModelInfo(doc.GetPathName() ?? (doc.GetTitle() ?? "UnsavedModel"));
                    var vr = validator.Validate(swInfo, doc);
                    if (!vr.Success)
                    {
                        return new RunResult { Success = false, Message = "Validation failed: " + (vr.Summary ?? "Unknown") };
                    }
                }
                catch
                {
                }

                bool wasAlreadySheet = false;
                try { wasAlreadySheet = SolidWorksApiWrapper.HasSheetMetalFeature(doc); } catch { }

                var info = new NM.Core.ModelInfo();
                var cfg = (doc.ConfigurationManager != null && doc.ConfigurationManager.ActiveConfiguration != null)
                    ? doc.ConfigurationManager.ActiveConfiguration.Name : string.Empty;
                var pathOrTitle = doc.GetPathName();
                if (string.IsNullOrWhiteSpace(pathOrTitle)) pathOrTitle = doc.GetTitle() ?? "UnsavedModel";
                info.Initialize(pathOrTitle, cfg);

                var swModel = new SolidWorksModel(info, swApp);
                swModel.Attach(doc, cfg);

                // NEW: If this is actually a tube, route to tube flow
                try { ExternalStartAdapter.TryRunExternalStart(swApp, doc); } catch { }

                try
                {
                    swModel.LoadPropertiesFromSolidWorks();
                }
                catch { }

                try
                {
                    if (ShouldRouteToTube(info))
                    {
                        ErrorHandler.DebugLog($"[Router] (from Run) Tube detected; routing to Tube.");
                        return RunTube(swApp, doc, options);
                    }
                }
                catch { }

                var simple = new SimpleSheetMetalProcessor(swApp);
                bool ok = simple.ConvertToSheetMetalAndOptionallyFlatten(info, doc, flatten: true, options: options);

                if (!ok)
                {
                    var reason = info?.ProblemDescription ?? "Unknown";
                    return new RunResult { Success = false, WasAlreadySheet = wasAlreadySheet, Message = "Conversion failed: " + reason, ProblemDescription = reason };
                }

                // Epic 6 v1: extract metrics and compute manufacturing values
                double rawWeightLb = 0.0;
                try
                {
                    var metrics = MetricsExtractor.FromModel(doc, info);
                    var calc = ManufacturingCalculator.Compute(metrics, new CalcOptions { UseMassIfAvailable = true, QuoteEnabled = options?.QuoteEnabled ?? false });
                    rawWeightLb = calc.RawWeightLb;
                    // Log legacy properties already set elsewhere
                }
                catch { }

                // F115 laser/waterjet: cut metrics + speeds + OP20 properties
                try
                {
                    // Ensure a flat face is selected/available
                    IFace2 flatFace = null;
                    try { flatFace = SolidWorksApiWrapper.GetFixedFace(doc); } catch { }
                    if (flatFace == null)
                    {
                        ErrorHandler.DebugLog("[F115] No flat face; skipping OP20.");
                    }
                    else
                    {
                        var cut = FlatPatternAnalyzer.Extract(doc, flatFace);
                        // Build speed provider
                        using (var loader = new ExcelDataLoader())
                        {
                            loader.LoadAllExcelData();
                            var provider = new NM.SwAddin.Manufacturing.Laser.LaserSpeedExcelProvider(loader);
                            var metrics = MetricsExtractor.FromModel(doc, info);
                            // Fill from cut metrics
                            metrics.ApproxCutLengthIn = cut.TotalCutLengthIn;
                            // Legacy pierce formula: loops + 2 + floor(length/30)
                            int piercesBase = System.Math.Max(0, cut.PierceCount) + 2;
                            int tabs = (int)System.Math.Floor(cut.TotalCutLengthIn / 30.0);
                            metrics.PierceCount = piercesBase + System.Math.Max(0, tabs);

                            // Determine process
                            string op20 = info.CustomProperties.GetPropertyValue("OP20") as string;
                            bool isWaterjet = false;
                            if (!string.IsNullOrWhiteSpace(op20))
                            {
                                var u = op20.ToUpperInvariant();
                                isWaterjet = u.Contains("155") || u.Contains("WATERJET");
                            }
                            // If OP20 missing, set default string
                            if (string.IsNullOrWhiteSpace(op20))
                            {
                                info.CustomProperties.SetPropertyValue("OP20", "F115 - LASER");
                            }

                            var op = LaserCalculator.Compute(metrics, provider, isWaterjet, rawWeightLb);

                            // Property write sequence: OP20_S, OP20_R, F115_Price
                            info.CustomProperties.SetPropertyValue("OP20_S", op.SetupHours.ToString("0.##", CultureInfo.InvariantCulture));
                            info.CustomProperties.SetPropertyValue("OP20_R", op.RunHours.ToString("0.####", CultureInfo.InvariantCulture));
                            var hoursTotal = op.SetupHours + (op.RunHours * System.Math.Max(1, metrics.Quantity));
                            var price = hoursTotal * CostConstants.F115_COST;
                            info.CustomProperties.SetPropertyValue("F115_Price", price.ToString("0.##", CultureInfo.InvariantCulture));
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ErrorHandler.HandleError(proc, "F115 compute failed", ex, "Warning");
                }

                // F140: press brake timing and price
                try
                {
                    var bends = BendAnalyzer.AnalyzeBends(doc);
                    if (bends != null && bends.Count > 0)
                    {
                        var f140 = F140Calculator.Compute(bends, rawWeightLb, 1);
                        info.CustomProperties.SetPropertyValue("F140_S", f140.SetupHours.ToString("0.##", CultureInfo.InvariantCulture));
                        info.CustomProperties.SetPropertyValue("F140_R", f140.RunHours.ToString("0.####", CultureInfo.InvariantCulture));
                        info.CustomProperties.SetPropertyValue("F140_S_Cost", f140.SetupHours.ToString("0.##", CultureInfo.InvariantCulture));
                        info.CustomProperties.SetPropertyValue("F140_Price", f140.Price(1).ToString("0.##", CultureInfo.InvariantCulture));
                    }
                }
                catch (System.Exception ex)
                {
                    ErrorHandler.HandleError(proc, "F140 compute failed", ex, "Warning");
                }

                // F220: tapped hole detection and timing
                try
                {
                    var th = TappedHoleAnalyzer.Analyze(doc, info?.CustomProperties?.GetPropertyValue("rbMaterialType") as string ?? string.Empty);
                    if (th != null && (th.Setups > 0 || th.Holes > 0))
                    {
                        var f220 = F220Calculator.Compute(new F220Input { Setups = th.Setups, Holes = th.Holes });
                        info.CustomProperties.SetPropertyValue("F220", "1");
                        info.CustomProperties.SetPropertyValue("F220_S", f220.SetupHours.ToString("0.##", CultureInfo.InvariantCulture));
                        info.CustomProperties.SetPropertyValue("F220_R", f220.RunHours.ToString("0.###", CultureInfo.InvariantCulture));
                        info.CustomProperties.SetPropertyValue("F220_RN", "TAP HOLES PER CAD");
                        if (th.StainlessNote)
                        {
                            info.CustomProperties.SetPropertyValue("F220_Note", "Verify drill size for stainless");
                        }
                        double f220Price = (f220.SetupHours + f220.RunHours) * 65.0; // fallback if constant missing
                        try { f220Price = (f220.SetupHours + f220.RunHours) * CostConstants.F220_COST; } catch { }
                        info.CustomProperties.SetPropertyValue("F220_Price", f220Price.ToString("0.##", CultureInfo.InvariantCulture));
                    }
                }
                catch (System.Exception ex)
                {
                    ErrorHandler.HandleError(proc, "F220 compute failed", ex, "Warning");
                }

                // Quoting / Total cost aggregation per legacy rules
                try
                {
                    if (options != null && options.QuoteEnabled)
                    {
                        // Quantity: UI overrides model property QuoteQty, default 1
                        int qty = 1;
                        if (options.Quantity > 0) qty = options.Quantity;
                        else
                        {
                            string propQty = info.CustomProperties.GetPropertyValue("QuoteQty") as string;
                            if (int.TryParse(propQty, out var q2) && q2 > 0) qty = q2;
                        }

                        // Material cost per LB: UI only
                        double costPerLb = options.CostPerLB > 0 ? options.CostPerLB : 0.0;

                        // Work center costs (prices) from properties
                        double f115 = ParseDouble(info.CustomProperties.GetPropertyValue("F115_Price"));
                        double f140p = ParseDouble(info.CustomProperties.GetPropertyValue("F140_Price"));
                        double f220p = ParseDouble(info.CustomProperties.GetPropertyValue("F220_Price"));
                        double f325 = ParseDouble(info.CustomProperties.GetPropertyValue("F325_Price"));

                        var tcIn = new TotalCostInputs
                        {
                            Quantity = qty,
                            RawWeightLb = rawWeightLb,
                            MaterialCostPerLb = costPerLb,
                            F115Price = f115,
                            F140Price = f140p,
                            F220Price = f220p,
                            F325Price = f325,
                            Difficulty = options.Difficulty
                        };
                        var tc = TotalCostCalculator.Compute(tcIn);

                        // Write properties (formats per legacy)
                        info.CustomProperties.SetPropertyValue("QuoteQty", tc.Quantity.ToString(CultureInfo.InvariantCulture));
                        info.CustomProperties.SetPropertyValue("MaterialCostPerLB", tc.MaterialCostPerLB.ToString("0.##", CultureInfo.InvariantCulture));
                        info.CustomProperties.SetPropertyValue("MaterailCostPerLB", tc.MaterialCostPerLB.ToString("0.##", CultureInfo.InvariantCulture));
                        info.CustomProperties.SetPropertyValue("TotalPrice", tc.TotalCost.ToString("0.##", CultureInfo.InvariantCulture));
                    }
                }
                catch (System.Exception ex)
                {
                    ErrorHandler.HandleError(proc, "Total cost aggregation failed", ex, "Warning");
                }

                try
                {
                    if (info.CustomProperties.IsDirty)
                    {
                        if (!swModel.SavePropertiesToSolidWorks())
                        {
                            var reason = info?.ProblemDescription ?? "Property writeback failed";
                            return new RunResult { Success = false, WasAlreadySheet = wasAlreadySheet, Message = "Property writeback failed.", ProblemDescription = reason };
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ErrorHandler.HandleError(proc, "Exception during property writeback", ex, "Warning");
                    var reason = info?.ProblemDescription ?? ex.Message;
                    return new RunResult { Success = false, WasAlreadySheet = wasAlreadySheet, Message = "Property writeback exception.", ProblemDescription = reason };
                }

                string msg = wasAlreadySheet ? "Already sheet metal: Flattened OK" : "Sheet metal conversion OK";
                return new RunResult { Success = true, WasAlreadySheet = wasAlreadySheet, Message = msg };
            }
            catch (System.Exception ex)
            {
                ErrorHandler.HandleError(proc, "Runner exception", ex);
                return new RunResult { Success = false, Message = "Exception: " + ex.Message };
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        private static double ParseDouble(object v)
        {
            try
            {
                if (v == null) return 0.0;
                if (v is double d) return d;
                if (double.TryParse(v.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var x)) return x;
                if (double.TryParse(v.ToString(), NumberStyles.Any, CultureInfo.CurrentCulture, out x)) return x;
                return 0.0;
            }
            catch { return 0.0; }
        }

        public static RunResult RunTube(ISldWorks swApp, IModelDoc2 doc, ProcessingOptions options)
        {
            const string proc = nameof(MainRunner) + ".RunTube";
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (swApp == null || doc == null)
                {
                    return new RunResult { Success = false, Message = "No active document.", ProblemDescription = "Null app/doc" };
                }

                var info = new NM.Core.ModelInfo();
                var cfg = (doc.ConfigurationManager != null && doc.ConfigurationManager.ActiveConfiguration != null)
                    ? doc.ConfigurationManager.ActiveConfiguration.Name : string.Empty;
                var pathOrTitle = doc.GetPathName();
                if (string.IsNullOrWhiteSpace(pathOrTitle)) pathOrTitle = doc.GetTitle() ?? "UnsavedModel";
                info.Initialize(pathOrTitle, cfg);

                var swModel = new SolidWorksModel(info, swApp);
                swModel.Attach(doc, cfg);

                // Try legacy extractor first
                ExternalStartAdapter.TryRunExternalStart(swApp, doc);

                try
                {
                    if (!swModel.LoadPropertiesFromSolidWorks())
                    {
                        ErrorHandler.HandleError(proc, "Failed to load custom properties; continuing with defaults", null, "Warning");
                    }
                }
                catch
                {
                    ErrorHandler.HandleError(proc, "Exception while loading properties; continuing with defaults", null, "Warning");
                }

                var tube = new SimpleTubeProcessor(swApp);
                bool ok = tube.Process(info, doc, options);
                if (!ok)
                {
                    var reason = info?.ProblemDescription ?? "Tube processing failed";
                    return new RunResult { Success = false, Message = "Tube processing failed: " + reason, ProblemDescription = reason };
                }

                if (info.CustomProperties.IsDirty)
                {
                    if (!swModel.SavePropertiesToSolidWorks())
                    {
                        return new RunResult { Success = false, Message = "Tube property writeback failed." };
                    }
                }

                return new RunResult { Success = true, Message = "Tube processing OK" };
            }
            catch (System.Exception ex)
            {
                ErrorHandler.HandleError(proc, "Runner exception", ex);
                return new RunResult { Success = false, Message = "Exception: " + ex.Message };
            }
            finally { ErrorHandler.PopCallStack(); }
        }
    }
}
