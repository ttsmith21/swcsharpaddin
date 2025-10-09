using System;
using SolidWorks.Interop.sldworks;
using NM.Core; // For ErrorHandler
using System.Threading;
using System.Reflection;
using SolidWorks.Interop.gtswutilities; // strong-typed Utilities interop
using System.IO;

namespace NM.StepClassifierAddin.Classification
{
    internal static class ThicknessAnalyzer
    {
        private static bool TrySetIntervals(object taObj, int intervals)
        {
            try
            {
                var t = taObj.GetType();
                string[] methodNames = { "SetIntervalCount", "SetIntervals", "SetNumIntervals", "SetRanges", "SetNumberOfIntervals" };
                foreach (var m in methodNames)
                {
                    try { t.InvokeMember(m, BindingFlags.InvokeMethod, null, taObj, new object[] { intervals }); ErrorHandler.DebugLog($"ThicknessAnalyzer: {m}({intervals}) invoked."); return true; } catch { }
                }
                string[] propNames = { "IntervalCount", "Intervals", "Ranges", "NumIntervals" };
                foreach (var p in propNames)
                {
                    try { t.InvokeMember(p, BindingFlags.SetProperty, null, taObj, new object[] { intervals }); ErrorHandler.DebugLog($"ThicknessAnalyzer: set {p}={intervals} via property."); return true; } catch { }
                }
            }
            catch { }
            ErrorHandler.DebugLog("ThicknessAnalyzer: Could not set intervals (API may not support it).");
            return false;
        }

        private static int RunThickWithRetry(IThicknessAnalysis ta, double minM, double maxM, int resOpt, int resultOpt, string reportPath)
        {
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                int rc = ta.RunThickAnalysis2(minM, maxM, false, resOpt, resultOpt, reportPath, false, false, false);
                ErrorHandler.DebugLog($"ThicknessAnalyzer: RunThickAnalysis2 attempt {attempt} rc={rc}");
                if (rc == 0) return 0;
                try { ta.Init(); } catch { }
                Thread.Sleep(150);
            }
            return -1;
        }

        private static int RunThinWithRetry(IThicknessAnalysis ta, double target, int resOpt, int resultOpt, string reportPath)
        {
            // IMPORTANT: Do NOT call Init() here. Re-initializing between thin retries can drop the analyzed body context
            // and lead to persistent rc=12004 errors. Just retry the run with a short delay.
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                int rc = ta.RunThinAnalysis2(target, resOpt, resultOpt, reportPath, false, false, false);
                ErrorHandler.DebugLog($"ThicknessAnalyzer: RunThinAnalysis2 target={target:F6} attempt {attempt} rc={rc}");
                if (rc == 0) return 0;
                Thread.Sleep(150);
            }
            return -1;
        }

        public static bool TryGetModeAndCoverage(ISldWorks app, IBody2 body, out double tMode, out double coverage)
        {
            tMode = 0.0;
            coverage = 0.0;
            if (app == null || body == null)
            {
                ErrorHandler.DebugLog("ThicknessAnalyzer: app or body is null.");
                return false;
            }

            try
            {
                // Pre-run: clear selections and rebuild to stabilize Utilities
                try
                {
                    var model = app.IActiveDoc2 as IModelDoc2;
                    model?.ClearSelection2(true);
                    model?.ForceRebuild3(false);
                }
                catch { }

                var utilApp = app.GetAddInObject("Utilities.UtilitiesApp") as IUtilities;
                if (utilApp != null)
                {
                    var ta = utilApp.ThicknessAnalysis;
                    if (ta != null)
                    {
                        ErrorHandler.DebugLog("ThicknessAnalyzer: Using typed Utilities IThicknessAnalysis.");
                        int initRc = ta.Init();
                        ErrorHandler.DebugLog($"ThicknessAnalyzer: Init rc={initRc}");
                        if (initRc == 0)
                        {
                            double IN2M = 0.0254;
                            int resOpt = (int)gttckResolutionOptions_e.gttckHighResolution;
                            int resultOpt = 0; // no report output
                            // Use a non-null report path to satisfy versions sensitive to null
                            string reportDir = Path.Combine(Path.GetTempPath(), "NM_TckReport");
                            try { Directory.CreateDirectory(reportDir); } catch { }
                            string reportPath = Path.Combine(reportDir, "report");

                            TrySetIntervals(ta, 8);

                            // Thick band to get coarse distribution
                            double minIn = 0.020, maxIn = 1.000;
                            double minM = minIn * IN2M, maxM = maxIn * IN2M;
                            int rcThick = RunThickWithRetry(ta, minM, maxM, resOpt, resultOpt, reportPath);
                            if (rcThick != 0)
                            {
                                ErrorHandler.DebugLog("ThicknessAnalyzer: Thick analysis failed after retries.");
                                try { ta.Close(); } catch { }
                                return false;
                            }

                            int err;
                            int bins = ta.GetIntervalCount(out err);
                            ErrorHandler.DebugLog($"ThicknessAnalyzer: Thick bins={bins}, err={err}");
                            if (err != 0 || bins <= 0) { try { ta.Close(); } catch { } return false; }

                            double maxShare = -1.0; int modeIdx = -1;
                            var lows = new double[bins + 1];
                            var highs = new double[bins + 1];
                            for (int i = 1; i <= bins; i++)
                            {
                                double low, high, area, perAnalArea; int numFaces;
                                int drc = ta.GetAnalysisDetails(i, out low, out high, out numFaces, out area, out perAnalArea);
                                if (drc != 0) continue;
                                double share = perAnalArea / 100.0;
                                lows[i] = low; highs[i] = high;
                                ErrorHandler.DebugLog($"  Thick Bin {i}: low={low:F4}, high={high:F4}, faces={numFaces}, share={share:P2}, area={area:F6}");
                                if (share > maxShare) { maxShare = share; modeIdx = i; }
                            }
                            if (modeIdx <= 0) { try { ta.Close(); } catch { } return false; }

                            // Use HIGH of thick mode bin as first thin target
                            double thinTarget = highs[modeIdx];
                            ErrorHandler.DebugLog($"ThicknessAnalyzer: Thick mode bin={modeIdx}; using bin HIGH as thin target={thinTarget:F4}m");

                            // Iteratively refine thin target if modal bin is still the last bin
                            double finalT = 0.0; double finalCov = 0.0;
                            int maxRefines = 3; int attempt = 0;
                            while (attempt < maxRefines)
                            {
                                attempt++;
                                TrySetIntervals(ta, 8);
                                int rcThin = RunThinWithRetry(ta, thinTarget, resOpt, resultOpt, reportPath);
                                if (rcThin != 0) break;

                                int binsT = ta.GetIntervalCount(out err);
                                ErrorHandler.DebugLog($"ThicknessAnalyzer: Thin bins={binsT}, err={err}");
                                if (err != 0 || binsT <= 0) break;

                                double maxShareT = -1.0; int modeIdxT = -1; double tLocal = 0.0;
                                var lowsT = new double[binsT + 1];
                                var highsT = new double[binsT + 1];
                                var sharesT = new double[binsT + 1];
                                for (int i = 1; i <= binsT; i++)
                                {
                                    double low, high, area, perAnalArea; int numFaces;
                                    int drc = ta.GetAnalysisDetails(i, out low, out high, out numFaces, out area, out perAnalArea);
                                    if (drc != 0) continue;
                                    double share = perAnalArea / 100.0;
                                    lowsT[i] = low; highsT[i] = high; sharesT[i] = share;
                                    ErrorHandler.DebugLog($"  Thin Bin {i}: low={low:F4}, high={high:F4}, share={share:P2}");
                                    if (share > maxShareT) { maxShareT = share; modeIdxT = i; tLocal = 0.5 * (low + high); }
                                }
                                if (modeIdxT <= 0) break;

                                // Compute ±5% coverage
                                double loBand = tLocal * 0.95, hiBand = tLocal * 1.05;
                                double covLocal = 0.0;
                                for (int i = 1; i <= binsT; i++) if (highsT[i] >= loBand && lowsT[i] <= hiBand) covLocal += sharesT[i];

                                // Store current results
                                finalT = tLocal; finalCov = covLocal;

                                // If mode is not the last bin, stop refining
                                if (modeIdxT < binsT) { ErrorHandler.DebugLog($"ThicknessAnalyzer: Strategy SUCCESS tMode={finalT:F4}m, cov±5%={finalCov:P2}"); break; }

                                // If mode is still the last bin, lower the target near that bin: new target = low*1.10
                                double lastLow = lowsT[modeIdxT];
                                double newTarget = Math.Max(lastLow * 1.10, lastLow + 1e-6);
                                ErrorHandler.DebugLog($"ThicknessAnalyzer: Refining thin target from {thinTarget:F4}m ? {newTarget:F4}m (last bin low * 1.10)");

                                if (Math.Abs(newTarget - thinTarget) / Math.Max(1e-9, thinTarget) < 0.01) { ErrorHandler.DebugLog("ThicknessAnalyzer: Target change <1%; stopping refine."); break; }
                                thinTarget = newTarget;
                            }

                            tMode = finalT; coverage = finalCov;
                            try { ta.Close(); } catch { }
                            if (tMode > 0) return true;

                            // Fallback refine and multi-seed paths unchanged...
                            double[] refineSeedsIn = new double[] { 0.030, 0.039, 0.048, 0.059, 0.075, 0.105 };
                            var best = new { cov = -1.0, t = 0.0, seedIdx = -1 };
                            var all = new System.Collections.Generic.List<(double cov, double t, int idx)>();

                            for (int si = 0; si < refineSeedsIn.Length; si++)
                            {
                                double targetM = refineSeedsIn[si] * IN2M;
                                int rcThin = RunThinWithRetry(ta, targetM, resOpt, resultOpt, reportPath);
                                if (rcThin != 0) { ErrorHandler.DebugLog($"  Refine seed {refineSeedsIn[si]:F3}in rc={rcThin}"); continue; }

                                int err2; int bins2 = ta.GetIntervalCount(out err2);
                                ErrorHandler.DebugLog($"ThicknessAnalyzer: Thin(refine {refineSeedsIn[si]:F3}in) bins={bins2}, err={err2}");
                                if (err2 != 0 || bins2 <= 0) continue;

                                double maxShare2 = -1.0; int modeIdx2 = -1; double tLocal2 = 0.0;
                                double[] lows2 = new double[bins2 + 1];
                                double[] highs2 = new double[bins2 + 1];
                                double[] shares2 = new double[bins2 + 1];
                                for (int i = 1; i <= bins2; i++)
                                {
                                    double low, high, area, perAnalArea; int numFaces;
                                    int detRc = ta.GetAnalysisDetails(i, out low, out high, out numFaces, out area, out perAnalArea);
                                    if (detRc != 0) continue;
                                    double share = perAnalArea / 100.0;
                                    lows2[i] = low; highs2[i] = high; shares2[i] = share;
                                    ErrorHandler.DebugLog($"    Thin(refine) Bin {i}: low={low:F4}, high={high:F4}, share={share:P2}");
                                    if (share > maxShare2) { maxShare2 = share; modeIdx2 = i; tLocal2 = 0.5 * (low + high); }
                                }
                                if (modeIdx2 <= 0) continue;

                                double loBand2 = tLocal2 * 0.95, hiBand2 = tLocal2 * 1.05;
                                double covLocal2 = 0.0;
                                for (int i = 1; i <= bins2; i++) if (highs2[i] >= loBand2 && lows2[i] <= hiBand2) covLocal2 += shares2[i];

                                ErrorHandler.DebugLog($"  Refine candidate seed={refineSeedsIn[si]:F3}in ? tMode={tLocal2:F4}m, cov±5%={covLocal2:P2}");
                                all.Add((covLocal2, tLocal2, si));
                                if (covLocal2 > best.cov) best = new { cov = covLocal2, t = tLocal2, seedIdx = si };
                            }

                            if (best.seedIdx >= 0)
                            {
                                double tol = Math.Max(0.0, best.cov - 0.04);
                                double chosenT = best.t; double chosenCov = best.cov; int chosenIdx = best.seedIdx;
                                foreach (var c in all) if (c.cov >= tol && c.t < chosenT) { chosenT = c.t; chosenCov = c.cov; chosenIdx = c.idx; }
                                tMode = chosenT; coverage = chosenCov;
                                try { ta.Close(); } catch { }
                                ErrorHandler.DebugLog($"ThicknessAnalyzer: Refine-picked seed[{chosenIdx}] {refineSeedsIn[chosenIdx]:F3}in. tMode={tMode:F4}m, cov±5%={coverage:P2}");
                                return true;
                            }

                            double[] seedsIn = new double[] { 0.118, 0.25, 0.375, 0.188, 0.177, 0.059, 0.075, 0.105, 0.135, 0.165, 0.048, 0.500, 1.000, 0.315, 0.149, 0.197, 0.157, 0.098, 0.039, 0.030, 0.024 };
                            double bestCov2 = -1.0; double bestT2 = 0.0; int bestSeed2 = -1;

                            for (int si = 0; si < seedsIn.Length; si++)
                            {
                                double targetM = seedsIn[si] * IN2M;
                                int rcThin = RunThinWithRetry(ta, targetM, resOpt, resultOpt, reportPath);
                                if (rcThin != 0) { ErrorHandler.DebugLog($"  Thin(seed {seedsIn[si]:F3}in) rc={rcThin}"); continue; }
                                int err3; int bins3 = ta.GetIntervalCount(out err3);
                                ErrorHandler.DebugLog($"ThicknessAnalyzer: Thin(seed {seedsIn[si]:F3}in) bins={bins3}, err={err3}");
                                if (err3 != 0 || bins3 <= 0) continue;
                                double maxShare3 = -1.0; int modeIdx3 = -1; double tLocal3 = 0.0;
                                double[] lows3 = new double[bins3 + 1];
                                double[] highs3 = new double[bins3 + 1];
                                double[] shares3 = new double[bins3 + 1];
                                for (int i = 1; i <= bins3; i++)
                                {
                                    double low, high, area, perAnalArea; int numFaces;
                                    int detRc = ta.GetAnalysisDetails(i, out low, out high, out numFaces, out area, out perAnalArea);
                                    if (detRc != 0) continue;
                                    double share = perAnalArea / 100.0;
                                    lows3[i] = low; highs3[i] = high; shares3[i] = share;
                                    ErrorHandler.DebugLog($"    Thin(seed) Bin {i}: low={low:F4}, high={high:F4}, share={share:P2}");
                                    if (share > maxShare3) { maxShare3 = share; modeIdx3 = i; tLocal3 = 0.5 * (low + high); }
                                }
                                if (modeIdx3 <= 0) continue;
                                double loBand3 = tLocal3 * 0.95, hiBand3 = tLocal3 * 1.05;
                                double covLocal3 = 0.0; for (int i = 1; i <= bins3; i++) if (highs3[i] >= loBand3 && lows3[i] <= hiBand3) covLocal3 += shares3[i];
                                if (covLocal3 > bestCov2) { bestCov2 = covLocal3; bestT2 = tLocal3; bestSeed2 = si; }
                            }
                            try { ta.Close(); } catch { }
                            if (bestSeed2 >= 0)
                            {
                                tMode = bestT2; coverage = bestCov2;
                                ErrorHandler.DebugLog($"ThicknessAnalyzer: Thin analysis SUCCESS best seed[{bestSeed2}] target={seedsIn[bestSeed2]:F3}in. tMode={tMode:F4}m, coverage±5%={coverage:P2}");
                                return true;
                            }
                        }
                        else
                        {
                            ErrorHandler.DebugLog($"ThicknessAnalyzer: Init failed with rc={initRc}");
                        }
                    }
                }
                else
                {
                    ErrorHandler.DebugLog("ThicknessAnalyzer: IUtilities cast failed; will try late-binding.");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"ThicknessAnalyzer: Typed Utilities path failed: {ex.Message}; falling back to late-binding.");
            }

            // Late-binding with retries (fallback) — unchanged
            const int maxRetries = 5;
            const int delayMs = 250;

            try
            {
                object utilObj = app.GetAddInObject("Utilities.UtilitiesApp");
                if (utilObj == null)
                {
                    ErrorHandler.DebugLog("ThicknessAnalyzer: GetAddInObject('Utilities.UtilitiesApp') returned null. Is the add-in enabled and licensed (Pro/Premium)?");
                    return false;
                }

                // Try to acquire the ThicknessAnalysis COM object via multiple names and strategies
                object anal = null;
                var utilsType = utilObj.GetType();
                string[] propNames = new[] { "IThicknessAnalysis", "ThicknessAnalysis", "ThicknessCheck", "Thickness" };
                string[] methodNames = new[] { "CreateThicknessAnalysis", "GetThicknessAnalysis", "CreateIThicknessAnalysis" };

                for (int attempt = 1; attempt <= maxRetries && anal == null; attempt++)
                {
                    foreach (var pn in propNames)
                    {
                        try
                        {
                            var val = utilsType.InvokeMember(pn, BindingFlags.GetProperty, null, utilObj, null);
                            if (val != null)
                            {
                                anal = val;
                                ErrorHandler.DebugLog($"ThicknessAnalyzer: Acquired property '{pn}' on attempt {attempt}.");
                                break;
                            }
                        }
                        catch { }
                    }

                    if (anal == null)
                    {
                        foreach (var mn in methodNames)
                        {
                            try
                            {
                                var val = utilsType.InvokeMember(mn, BindingFlags.InvokeMethod, null, utilObj, null);
                                if (val != null)
                                {
                                    anal = val;
                                    ErrorHandler.DebugLog($"ThicknessAnalyzer: Acquired via method '{mn}' on attempt {attempt}.");
                                    break;
                                }
                            }
                            catch { }
                        }
                    }

                    if (anal == null)
                    {
                        ErrorHandler.DebugLog($"ThicknessAnalyzer: ThicknessAnalysis object null (attempt {attempt}/{maxRetries}); sleeping {delayMs} ms...");
                        Thread.Sleep(delayMs);
                    }
                }
                if (anal == null)
                {
                    ErrorHandler.DebugLog("ThicknessAnalyzer: Could not get ThicknessAnalysis object after retries.");
                    return false;
                }

                // Minimal late-binding run: attempt Run/RunThinAnalysis2 then parse bins (names may differ across versions)
                var analType = anal.GetType();
                bool ranOk = false;
                string[] runNames = new[] { "RunThinAnalysis2", "RunAnalysis2", "Run" };
                foreach (var rn in runNames)
                {
                    if (ranOk) break;
                    try { var res = analType.InvokeMember(rn, BindingFlags.InvokeMethod, null, anal, null); ranOk = res is bool b ? b : true; } catch { }
                }
                if (!ranOk)
                {
                    ErrorHandler.DebugLog("ThicknessAnalyzer: Could not run analysis (no compatible run method).");
                    return false;
                }

                int binsFallback = -1;
                try { binsFallback = Convert.ToInt32(analType.InvokeMember("GetIntervalCount", BindingFlags.InvokeMethod, null, anal, null)); }
                catch { try { binsFallback = Convert.ToInt32(analType.InvokeMember("IntervalCount", BindingFlags.GetProperty, null, anal, null)); } catch { }
                }
                ErrorHandler.DebugLog($"ThicknessAnalyzer: Interval count: {binsFallback}");
                if (binsFallback <= 0) return false;

                double maxPctF = 0.0; int modeIdxF = -1; double tModeF = 0.0; double totalCovF = 0.0;
                for (int i = 0; i < binsFallback; i++)
                {
                    object[] detArgs = new object[] { i, 0.0, 0.0, 0.0, 0.0 };
                    try { analType.InvokeMember("GetAnalysisDetails", BindingFlags.InvokeMethod, null, anal, detArgs); } catch { continue; }
                    double low = Convert.ToDouble(detArgs[1]);
                    double high = Convert.ToDouble(detArgs[2]);
                    double pct = Convert.ToDouble(detArgs[3]);
                    ErrorHandler.DebugLog($"  LB Bin {i}: low={low:F4}, high={high:F4}, pct={pct:F2}");
                    if (pct > maxPctF) { maxPctF = pct; modeIdxF = i; tModeF = 0.5 * (low + high); }
                }
                for (int i = Math.Max(0, modeIdxF - 1); i <= Math.Min(binsFallback - 1, modeIdxF + 1); i++)
                {
                    object[] detArgs = new object[] { i, 0.0, 0.0, 0.0, 0.0 };
                    try { analType.InvokeMember("GetAnalysisDetails", BindingFlags.InvokeMethod, null, anal, detArgs); totalCovF += Convert.ToDouble(detArgs[3]) / 100.0; } catch { }
                }
                tMode = tModeF; coverage = totalCovF;
                ErrorHandler.DebugLog($"ThicknessAnalyzer: Late-binding SUCCESS tMode={tMode:F4}m, cov={coverage:P2}");
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("ThicknessAnalyzer", $"Exception during thickness analysis: {ex.Message}", ex);
                return false;
            }
        }

        // Returns the face count divided by volume (m^3) for a body. High values indicate complex hardware (threads/knurls).
        // Requires the owning IModelDoc2 to get MassProperty2.
        public static double GetComplexityRatio(IModelDoc2 model, IBody2 body)
        {
            if (body == null || model == null) return 0.0;
            int faceCount = body.GetFaceCount();
            // Use MassProperty2 for modern, robust mass/volume queries
            var massProp = model.Extension.CreateMassProperty2() as MassProperty2;
            double volume = (massProp != null) ? massProp.Volume : 0.0;
            if (volume <= 0.0) return 0.0;
            return faceCount / volume;
        }

        // Returns true if the body is likely complex hardware (e.g., modeled threads/knurls) based on ratio threshold.
        public static bool IsComplexHardware(IModelDoc2 model, IBody2 body, double threshold = 5000)
        {
            double ratio = GetComplexityRatio(model, body);
            return ratio > threshold;
        }
    }
}