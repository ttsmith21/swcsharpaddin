using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin
{
    public sealed class SheetMetalProcessor
    {
        private readonly ISldWorks _swApp;
        private readonly SolidWorksFileOperations _fileOps;
        private readonly ProblemPartTracker _problemTracker;

        // Clamp unrealistic sheet thickness values (meters). 1 inch max.
        private const double MaxSheetThicknessMeters = 0.0254;

        public SheetMetalProcessor(ISldWorks swApp, SolidWorksFileOperations fileOps, ProblemPartTracker problemTracker = null)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
            _fileOps = fileOps ?? throw new ArgumentNullException(nameof(fileOps));
            _problemTracker = problemTracker; // optional
        }

        // ------------------------ Diagnostics helpers ------------------------
        private static bool IsDiagnosticsEnabled
        {
            get
            {
                try
                {
                    var t = typeof(NM.Core.Configuration.Logging);
                    var pDiag = t.GetProperty("EnableDiagnostics");
                    if (pDiag != null)
                    {
                        var v = pDiag.GetValue(null, null);
                        return v is bool b && b;
                    }
                }
                catch { }
                try
                {
                    // Fallback to Debug mode
                    var t = typeof(NM.Core.Configuration.Logging);
                    var pDbg = t.GetProperty("EnableDebugMode");
                    if (pDbg != null)
                    {
                        var v = pDbg.GetValue(null, null);
                        return v is bool b && b;
                    }
                }
                catch { }
                return false;
            }
        }

        private static void DLog(string msg)
        {
            if (IsDiagnosticsEnabled && !string.IsNullOrWhiteSpace(msg))
                ErrorHandler.DebugLog(msg);
        }

        // Add small math wrappers to resolve references
        private static double[] Normalize(double[] v) => NM.StepClassifierAddin.Utils.Math3D.Normalize(v);
        private static double Dot(double[] a, double[] b) => NM.StepClassifierAddin.Utils.Math3D.Dot(a, b);

        // Provide conservative stubs if advanced analysis helpers are missing
        private static void ComputeCylinderShareAndWallThickness(Body2 body, out double cylShare, out double wall, out double[] axis)
        {
            cylShare = 0.0; wall = 0.0; axis = null;
            try
            {
                if (body == null) return;
                double total = 0.0, cyl = 0.0;
                var faces = body.GetFaces() as object[];
                if (faces != null)
                {
                    foreach (var fo in faces)
                    {
                        var f = fo as IFace2; if (f == null) continue;
                        double a = 0.0; try { a = f.GetArea(); } catch { }
                        total += a;
                        var s = f.IGetSurface() as ISurface; if (s == null) continue;
                        bool isCyl = false; try { isCyl = s.IsCylinder(); } catch { }
                        if (isCyl) cyl += a;
                    }
                }
                cylShare = (total > 0.0) ? (cyl / total) : 0.0;
            }
            catch { cylShare = 0.0; wall = 0.0; axis = null; }
        }

        private static bool IsRolledCylinder(Body2 body, out double diameter)
        {
            diameter = 0.0;
            try
            {
                if (body == null) return false;
                var faces = body.GetFaces() as object[];
                if (faces == null) return false;
                double radius = 0.0; double totalArea = 0.0; double cylArea = 0.0;
                foreach (var fo in faces)
                {
                    var f = fo as IFace2; if (f == null) continue;
                    double a = 0.0; try { a = f.GetArea(); } catch { }
                    totalArea += a;
                    var s = f.IGetSurface() as ISurface; if (s == null) continue;
                    bool isCyl = false; try { isCyl = s.IsCylinder(); } catch { }
                    if (!isCyl) continue;
                    cylArea += a;
                    try
                    {
                        var cp = s.CylinderParams as object; var arr = cp as double[];
                        if (arr != null && arr.Length >= 7)
                        {
                            var r = Math.Abs(arr[6]);
                            if (r > 0) radius = Math.Max(radius, r);
                        }
                    }
                    catch { }
                }
                if (totalArea <= 0) return false;
                double share = cylArea / totalArea;
                if (share >= 0.9 && radius > 0)
                {
                    diameter = radius * 2.0;
                    return true;
                }
            }
            catch { }
            return false;
        }

        private static void LogSelectionState(IModelDoc2 model, string label)
        {
            if (!IsDiagnosticsEnabled || model == null) return;
            try
            {
                var sel = model.ISelectionManager;
                int totalAny = sel?.GetSelectedObjectCount2(-1) ?? 0;
                int m0 = sel?.GetSelectedObjectCount2(0) ?? 0;
                int m1 = sel?.GetSelectedObjectCount2(1) ?? 0;
                int m2 = sel?.GetSelectedObjectCount2(2) ?? 0;
                DLog($"Sel[{label}]: total={totalAny}, m0={m0}, m1={m1}, m2={m2}");
                for (int mark = 0; mark <= 2; mark++)
                {
                    int cnt = sel?.GetSelectedObjectCount2(mark) ?? 0;
                    if (cnt <= 0) continue;
                    string types = string.Empty;
                    for (int i = 1; i <= cnt; i++)
                    {
                        int t = sel.GetSelectedObjectType3(i, mark);
                        if (i <= 5) types += ((swSelectType_e)t).ToString() + (i < cnt ? "," : string.Empty);
                    }
                    if (cnt > 5) types += ",...";
                    DLog($"  mark {mark}: count={cnt}, types=[{types}]");
                }
            }
            catch { }
        }

        private static List<string> SnapshotFeatures(IModelDoc2 model)
        {
            var list = new List<string>();
            try
            {
                var f = model?.FirstFeature() as IFeature;
                while (f != null)
                {
                    string name = string.Empty; string type = string.Empty;
                    try { name = f.Name; } catch { }
                    try { type = f.GetTypeName2(); } catch { }
                    list.Add((name ?? string.Empty) + "|" + (type ?? string.Empty));
                    f = f.GetNextFeature() as IFeature;
                }
            }
            catch { }
            return list;
        }

        private static void LogFeatureDiff(List<string> before, List<string> after, string label)
        {
            if (!IsDiagnosticsEnabled) return;
            try
            {
                int b = before?.Count ?? 0; int a = after?.Count ?? 0;
                if (a <= b)
                {
                    DLog($"Feat[{label}]: before={b}, after={a} (no additions)");
                    return;
                }
                var added = new List<string>();
                for (int i = b; i < a; i++)
                {
                    string s = after[i];
                    if (!string.IsNullOrWhiteSpace(s)) added.Add(s);
                }
                string joined = string.Join(", ", added.ToArray());
                if (joined.Length > 240) joined = joined.Substring(0, 240) + "...";
                DLog($"Feat[{label}]: before={b}, after={a}, added=[{joined}]");
            }
            catch { }
        }

        private static void LogMassVolume(IModelDoc2 model, string label)
        {
            if (!IsDiagnosticsEnabled || model == null) return;
            try
            {
                double m = SolidWorksApiWrapper.GetModelMass(model);
                double v = SolidWorksApiWrapper.GetModelVolume(model);
                DLog($"MV[{label}]: mass={m}, vol={v}");
            }
            catch { }
        }
        // ---------------------- End diagnostics helpers ----------------------

        public async Task<bool> ConvertToSheetMetalAndFlattenAsync(ModelInfo modelInfo, IModelDoc2 model = null, CancellationToken cancellationToken = default)
        {
            const string proc = nameof(ConvertToSheetMetalAndFlattenAsync);
            ErrorHandler.PushCallStack(proc);
            PerformanceTracker.Instance.StartTimer(proc);

            try
            {
                if (modelInfo == null)
                {
                    ErrorHandler.HandleError(proc, "ModelInfo is null");
                    return false;
                }

                IModelDoc2 swModel = model;

                if (swModel == null)
                {
                    bool silent = NM.Core.Configuration.Logging.IsProductionMode;
                    swModel = _fileOps.OpenSWDocument(modelInfo.FilePath, silent: silent, readOnly: false, configurationName: modelInfo.ConfigurationName ?? string.Empty);
                    if (swModel == null)
                    {
                        modelInfo.ProblemDescription = "Failed to open model";
                        ErrorHandler.HandleError(proc, $"Failed to open model: {modelInfo.FilePath}");
                        return false;
                    }
                }

                if (!string.IsNullOrWhiteSpace(modelInfo.ConfigurationName))
                {
                    if (!SolidWorksApiWrapper.SetActiveConfiguration(swModel, modelInfo.ConfigurationName))
                    {
                        modelInfo.ProblemDescription = $"Failed to set configuration: {modelInfo.ConfigurationName}";
                        ErrorHandler.HandleError(proc, modelInfo.ProblemDescription, null, "Warning");
                        return false;
                    }
                }

                modelInfo.InitialMass = SolidWorksApiWrapper.GetModelMass(swModel);
                modelInfo.InitialVolume = SolidWorksApiWrapper.GetModelVolume(swModel);
                if (!NM.Core.Configuration.Logging.IsProductionMode)
                {
                    ErrorHandler.DebugLog($"{proc}: Initial Mass={modelInfo.InitialMass} kg, Volume={modelInfo.InitialVolume} m^3");
                }
                LogMassVolume(swModel, "start");

                // Pre-classification diagnostics (fast, read-only)
                try { LogPreClassificationMetrics(swModel); } catch { }

                if (CheckExistingSheetMetal(swModel))
                {
                    if (!FlattenSheetMetalPart(swModel))
                    {
                        modelInfo.ProblemDescription = "Failed to flatten existing sheet metal";
                        ErrorHandler.HandleError(proc, modelInfo.ProblemDescription);
                        return false;
                    }
                    modelInfo.IsSheetMetal = true;
                    modelInfo.IsFlattened = true;
                    modelInfo.InsertSuccessful = true;
                    return true;
                }

                if (await TryConversionPipeline(swModel, modelInfo, cancellationToken))
                {
                    return true;
                }

                modelInfo.ProblemDescription = "All conversion methods failed";
                _problemTracker?.AddProblemPart(modelInfo, modelInfo.ProblemDescription);
                return false;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, $"Exception: {ex.Message}", ex, "Error", context: modelInfo?.FilePath);
                modelInfo.ProblemDescription = ex.Message;
                return false;
            }
            finally
            {
                try
                {
                    if (model != null)
                    {
                        SolidWorksApiWrapper.ClearSelection(model);
                    }
                }
                catch { }

                PerformanceTracker.Instance.StopTimer(proc);
                ErrorHandler.PopCallStack();
                await Task.CompletedTask; // keep async signature without moving work off-thread
            }
        }

        private void LogPreClassificationMetrics(IModelDoc2 model)
        {
            const string proc = nameof(LogPreClassificationMetrics);
            ErrorHandler.PushCallStack(proc);
            try
            {
                var body = SolidWorksApiWrapper.GetMainBody(model);
                if (body == null) { ErrorHandler.DebugLog($"{proc}: No main body"); return; }
                double totalArea = 0.0, devArea = 0.0, cylArea = 0.0, planeArea = 0.0;
                var faces = ((Body2)body).GetFaces() as object[];
                if (faces != null)
                {
                    foreach (var fo in faces)
                    {
                        var f = fo as IFace2; if (f == null) continue;
                        double a = 0.0; try { a = f.GetArea(); } catch { }
                        totalArea += a;
                        var s = f.IGetSurface() as ISurface; if (s == null) continue;
                        bool isPlane = false, isCyl = false, isCone = false;
                        try { isPlane = s.IsPlane(); } catch { }
                        try { isCyl = s.IsCylinder(); } catch { }
                        try { isCone = s.IsCone(); } catch { }
                        if (isPlane) planeArea += a;
                        if (isCyl) cylArea += a;
                        if (isPlane || isCyl || isCone) devArea += a;
                    }
                }
                double devShare = totalArea > 0 ? devArea / totalArea : 0.0;
                double cylShare = totalArea > 0 ? cylArea / totalArea : 0.0;
                double radius = TryComputeCylinderRadius((Body2)body);
                ErrorHandler.DebugLog($"PreScan: ATotal={totalArea:F6} m^2, ADevShare={devShare:F3}, CylShare={cylShare:F3}, CylRadius(mm)={radius*1000:F2}");
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"{proc}: EX {ex.Message}");
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        public bool CheckExistingSheetMetal(IModelDoc2 swModel)
        {
            const string proc = nameof(CheckExistingSheetMetal);
            ErrorHandler.PushCallStack(proc);
            PerformanceTracker.Instance.StartTimer(proc);
            try
            {
                if (swModel == null) return false;
                var isSm = SolidWorksApiWrapper.HasSheetMetalFeature(swModel);
                if (!NM.Core.Configuration.Logging.IsProductionMode)
                {
                    ErrorHandler.DebugLog($"{proc}: {swModel.GetTitle()} SheetMetal={isSm}");
                }
                return isSm;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Exception during check", ex);
                return false;
            }
            finally
            {
                PerformanceTracker.Instance.StopTimer(proc);
                ErrorHandler.PopCallStack();
            }
        }

        private async Task<bool> TryConversionPipeline(IModelDoc2 swModel, ModelInfo info, CancellationToken token)
        {
            if (token.IsCancellationRequested) return false;

            // Decide strategy order from geometry
            double cylShare = 0.0, wall = 0.0; double[] axis;
            double cylRadius = 0.0; double cylDiamIn = 0.0;
            try
            {
                var body = SolidWorksApiWrapper.GetMainBody(swModel) as Body2;
                if (body != null) ComputeCylinderShareAndWallThickness(body, out cylShare, out wall, out axis);
                // quick radius estimate for large-diameter routing
                if (body != null)
                {
                    cylRadius = TryComputeCylinderRadius(body);
                    if (cylRadius > 0) cylDiamIn = cylRadius * 2.0 * 39.3701;
                }
            }
            catch { }

            // Decision logic for conversion strategy
            bool preferBendsFirst = cylShare <= 0.10; // mostly planar  // ?? BREAKPOINT 4
            bool cylinderDominant = cylShare >= 0.90; // mostly cylindrical
            bool largeCylinder = cylinderDominant && cylDiamIn >= 24.0;

            DLog($"PreScreenDecision: cylShare={cylShare:F3}, wall(mm)={wall*1000:F2}, cylDiam(in)={cylDiamIn:F1}, preferBendsFirst={preferBendsFirst}, cylinderDominant={cylinderDominant}, largeCylinder={largeCylinder}");

            if (cylinderDominant && largeCylinder)
            {
                // Prefer ConvertToSheetMetal first for large cylinders (rolled plate)
                ErrorHandler.DebugLog("Strategy1: ConvertToSheetMetal2 (large cylinder)");
                var bk1 = GetLastFeature(swModel);
                if (await TryInsertConvertToSheetMetalAsync(swModel, info, token))
                {
                    var ok1 = await ValidateAndSaveResultAsync(swModel, info, token);
                    if (ok1) return true;
                    ErrorHandler.DebugLog("Strategy1 failed; reverting");
                    try { RevertFeaturesAfter(swModel, bk1); } catch { }
                }

                if (token.IsCancellationRequested) return false;
                ErrorHandler.DebugLog("Strategy2: InsertBends on edge (fallback after CTS2)");
                var bk2 = GetLastFeature(swModel);
                if (await TryInsertBendsOnEdgeAsync(swModel, info, token))
                {
                    var ok2 = await ValidateAndSaveResultAsync(swModel, info, token);
                    if (ok2) return true;
                    ErrorHandler.DebugLog("Strategy2 failed; reverting");
                    try { RevertFeaturesAfter(swModel, bk2); } catch { }
                }
            }
            else if (cylinderDominant)
            {
                ErrorHandler.DebugLog("Strategy2: InsertBends on edge (cylinder-dominant)");
                var bk2 = GetLastFeature(swModel);
                if (await TryInsertBendsOnEdgeAsync(swModel, info, token))
                {
                    var ok2 = await ValidateAndSaveResultAsync(swModel, info, token);
                    if (ok2) return true;
                    ErrorHandler.DebugLog("Strategy2 failed; reverting");
                    try { RevertFeaturesAfter(swModel, bk2); } catch { }
                }

                if (token.IsCancellationRequested) return false;
                ErrorHandler.DebugLog("Strategy1: ConvertToSheetMetal2 (fallback after bends)");
                var bk1 = GetLastFeature(swModel);
                if (await TryInsertConvertToSheetMetalAsync(swModel, info, token))
                {
                    var ok1 = await ValidateAndSaveResultAsync(swModel, info, token);
                    if (ok1) return true;
                    ErrorHandler.DebugLog("Strategy1 failed; reverting");
                    try { RevertFeaturesAfter(swModel, bk1); } catch { }
                }
            }
            else if (preferBendsFirst)
            {
                ErrorHandler.DebugLog("Strategy2: InsertBends on edge (planar-dominant)");
                var bk2 = GetLastFeature(swModel);
                if (await TryInsertBendsOnEdgeAsync(swModel, info, token))
                {
                    var ok2 = await ValidateAndSaveResultAsync(swModel, info, token);
                    if (ok2) return true;
                    ErrorHandler.DebugLog("Strategy2 failed; reverting");
                    try { RevertFeaturesAfter(swModel, bk2); } catch { }
                }

                if (token.IsCancellationRequested) return false;
                ErrorHandler.DebugLog("Strategy1: ConvertToSheetMetal2");
                var bk1 = GetLastFeature(swModel);
                if (await TryInsertConvertToSheetMetalAsync(swModel, info, token))
                {
                    var ok1 = await ValidateAndSaveResultAsync(swModel, info, token);
                    if (ok1) return true;
                    ErrorHandler.DebugLog("Strategy1 failed; reverting");
                    try { RevertFeaturesAfter(swModel, bk1); } catch { }
                }
            }
            else
            {
                // Mixed: keep original order
                ErrorHandler.DebugLog("Strategy1: ConvertToSheetMetal2");
                var bk1 = GetLastFeature(swModel);
                if (await TryInsertConvertToSheetMetalAsync(swModel, info, token))
                {
                    var ok = await ValidateAndSaveResultAsync(swModel, info, token);
                    if (ok) return true;
                    ErrorHandler.DebugLog("Strategy1 failed; reverting");
                    try { RevertFeaturesAfter(swModel, bk1); } catch { }
                }
                if (token.IsCancellationRequested) return false;

                ErrorHandler.DebugLog("Strategy2: InsertBends on edge");
                var bk2 = GetLastFeature(swModel);
                if (await TryInsertBendsOnEdgeAsync(swModel, info, token))
                {
                    var ok = await ValidateAndSaveResultAsync(swModel, info, token);
                    if (ok) return true;
                    ErrorHandler.DebugLog("Strategy2 failed; reverting");
                    try { RevertFeaturesAfter(swModel, bk2); } catch { }
                }
            }

            if (token.IsCancellationRequested) return false;
            ErrorHandler.DebugLog("Strategy3: Face-based fallback");
            var bk3 = GetLastFeature(swModel);
            if (await InsertBendsFromFaceAsync(swModel, info, token))
            {
                var ok3 = await ValidateAndSaveResultAsync(swModel, info, token);
                if (ok3) return true;
                ErrorHandler.DebugLog("Strategy3 failed; reverting");
                try { RevertFeaturesAfter(swModel, bk3); } catch { }
            }

            return false;
        }

        // Strategy: InsertConvertToSheetMetal2 with two attempts: face-only then face+edges
        private async Task<bool> TryInsertConvertToSheetMetalAsync(IModelDoc2 swModel, ModelInfo info, CancellationToken token)
        {
            const string proc = nameof(TryInsertConvertToSheetMetalAsync);
            ErrorHandler.PushCallStack(proc);
            PerformanceTracker.Instance.StartTimer(proc);
            IFeature bookmark = null;
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, proc)) return false;

                var body = SolidWorksApiWrapper.GetMainBody(swModel);
                if (body == null)
                {
                    info.ProblemDescription = "Main body not found";
                    return false;
                }

                IFace2 baseFace = null;
                double derivedThickness = 0.0; double thkCoverage = 0.0;
                TryFindThicknessFromParallelFacesEx((Body2)body, out derivedThickness, out baseFace, out thkCoverage);

                double diameter; bool seamCylinder = IsRolledCylinder((Body2)body, out diameter);
                double thickness = info.ThicknessInMeters > 0 ? info.ThicknessInMeters : (derivedThickness > 0 ? derivedThickness : 0.001);
                // Clamp unrealistic thickness
                thickness = Math.Min(thickness, MaxSheetThicknessMeters);
                double radius = info.BendRadius > 0 ? info.BendRadius : (seamCylinder ? diameter / 2.0 : Math.Max(0.0005, thickness * 1.5));
                double kFactor = info.KFactor > 0 ? info.KFactor : (seamCylinder ? 0.5 : 0.4);
                ErrorHandler.DebugLog($"CTS2: derivedThk(mm)={(derivedThickness*1000):F2}, cov={thkCoverage:P0}, chosenThk(mm)={(thickness*1000):F2}, seamCyl={seamCylinder}, diam(mm)={(diameter*1000):F1}, R(mm)={(radius*1000):F2}, K={kFactor:F2}");

                SolidWorksApiWrapper.ClearSelection(swModel);
                if (baseFace == null)
                    baseFace = SolidWorksApiWrapper.GetLargestPlanarFace(body);

                var featsBefore = SnapshotFeatures(swModel);
                LogSelectionState(swModel, "pre-CTS2");

                // Attempt A: face only
                if (baseFace != null)
                {
                    SelectEntityWithMark(swModel, baseFace as IEntity, 0, append: false);
                    bookmark = GetLastFeature(swModel);
                    bool okA = swModel.FeatureManager.InsertConvertToSheetMetal2(
                        thickness, false, false, radius, 0.001, 0, 0.5, 1, 0.5, false);
                    var featsAfterA = SnapshotFeatures(swModel);
                    LogFeatureDiff(featsBefore, featsAfterA, "CTS2-A");
                    if (okA && SolidWorksApiWrapper.HasSheetMetalFeature(swModel))
                    {
                        info.ThicknessInMeters = thickness; info.BendRadius = radius; info.KFactor = kFactor;
                        DLog("SUCCESS: CTS2-A");
                        LogMassVolume(swModel, "post-CTS2-A");
                        return true;
                    }
                    // Revert and retry with edges
                    try { RevertFeaturesAfter(swModel, bookmark); } catch { }
                    SolidWorksApiWrapper.ClearSelection(swModel);
                }

                // Attempt B: face + up to 3 linear edges on that face
                if (baseFace != null)
                {
                    SelectEntityWithMark(swModel, baseFace as IEntity, 0, append: false);
                    int edgesSel = 0;
                    foreach (var e in GetCandidateRipEdges(baseFace, 3))
                    { if (IsLinearEdge(e)) { SelectEntityWithMark(swModel, e as IEntity, 1, append: true); edgesSel++; } }
                    ErrorHandler.DebugLog($"CTS2: baseFace selected, ripEdges={edgesSel}");
                }

                LogSelectionState(swModel, "pre-CTS2-B");
                bookmark = GetLastFeature(swModel);
                bool ok = swModel.FeatureManager.InsertConvertToSheetMetal2(
                    thickness, false, false, radius, 0.001, 0, 0.5, 1, 0.5, false);
                var featsAfterB = SnapshotFeatures(swModel);
                LogFeatureDiff(featsBefore, featsAfterB, "CTS2-B");
                if (!ok)
                {
                    info.ProblemDescription = "InsertConvertToSheetMetal2 returned false";
                    ErrorHandler.DebugLog("CTS2: API returned false");
                    RevertFeaturesAfter(swModel, bookmark);
                    return false;
                }

                if (!SolidWorksApiWrapper.HasSheetMetalFeature(swModel))
                {
                    info.ProblemDescription = "ConvertToSheetMetal2 created no feature";
                    RevertFeaturesAfter(swModel, bookmark);
                    return false;
                }

                info.ThicknessInMeters = thickness; info.BendRadius = radius; info.KFactor = kFactor;
                DLog("SUCCESS: CTS2-B");
                LogMassVolume(swModel, "post-CTS2-B");
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Exception in InsertConvertToSheetMetal2", ex);
                try { RevertFeaturesAfter(swModel, bookmark); } catch { }
                return false;
            }
            finally
            {
                PerformanceTracker.Instance.StopTimer(proc);
                ErrorHandler.PopCallStack();
                await Task.CompletedTask;
            }
        }

        private async Task<bool> TryInsertBendsOnEdgeAsync(IModelDoc2 swModel, ModelInfo info, CancellationToken token)
        {
            const string proc = nameof(TryInsertBendsOnEdgeAsync);
            ErrorHandler.PushCallStack(proc);
            PerformanceTracker.Instance.StartTimer(proc);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, proc)) return false;

                var body = SolidWorksApiWrapper.GetMainBody(swModel) as Body2;
                if (body == null) { info.ProblemDescription = "No main body"; return false; }

                double cylShare; double wallThk; double[] cylAxis;
                ComputeCylinderShareAndWallThickness(body, out cylShare, out wallThk, out cylAxis);
                bool forceCylFlow = cylShare >= 0.90;

                var part = swModel as IPartDoc; if (part == null) return false;

                if (forceCylFlow)
                {
                    // This logic remains unchanged as it handles a specific cylindrical case
                    // and does not use the two-phase approach.
                    ErrorHandler.DebugLog("InsertBends: cylinder flow");
                    IEdge fixedEdge = null;
                    var faces = body.GetFaces() as object[];
                    if (faces != null)
                    {
                        double bestLen = 0.0;
                        foreach (var fo in faces)
                        {
                            var f = fo as IFace2; if (f == null) continue;
                            var s = f.IGetSurface() as ISurface; if (s == null) continue;
                            bool isCyl = false; try { isCyl = s.IsCylinder(); } catch { }
                            if (!isCyl) continue;
                            var edges = f.GetEdges() as object[]; if (edges == null) continue;
                            foreach (var eo in edges)
                            {
                                var e = eo as IEdge; if (e == null) continue;
                                if (!IsLinearEdge(e)) continue;
                                double len = GetEdgeChordLength(e);
                                if (len > bestLen) { bestLen = len; fixedEdge = e; }
                            }
                        }
                    }
                    if (wallThk > 0) info.ThicknessInMeters = Math.Min(wallThk, MaxSheetThicknessMeters);

                    SolidWorksApiWrapper.ClearSelection(swModel);
                    if (fixedEdge != null) SelectEntityWithMark(swModel, fixedEdge as IEntity, 0, append: false);

                    double radius = info.BendRadius > 0 ? info.BendRadius : Math.Max(0.0005, (info.ThicknessInMeters > 0 ? info.ThicknessInMeters * 1.5 : 0.005));
                    double kFactorCyl = info.KFactor > 0 ? info.KFactor : NM.Core.Configuration.Defaults.DefaultKFactor;

                    bool okCyl = part.InsertBends2(radius, string.Empty, kFactorCyl, -1, true, 0.5, true);
                    if (!okCyl || !SolidWorksApiWrapper.HasSheetMetalFeature(swModel))
                    {
                        info.ProblemDescription = "Cylinder InsertBends2 failed";
                        var bookmark = GetLastFeature(swModel);
                        RevertFeaturesAfter(swModel, bookmark);
                        return false;
                    }

                    TryApplySheetMetalParameters(swModel, new ModelInfo
                    {
                        ThicknessInMeters = info.ThicknessInMeters,
                        BendRadius = radius,
                        KFactor = kFactorCyl
                    });
                    try { swModel.EditRebuild3(); } catch { }
                    DLog("SUCCESS: IB-cyl");
                    LogMassVolume(swModel, "post-IB-cyl");
                    return true;
                }
                else
                {
                    // --- Start of new two-phase implementation for planar/mixed geometry ---
                    DLog("=== START TWO-PHASE CONVERSION ===");
                    double initialVolume = SolidWorksApiWrapper.GetModelVolume(swModel);
                    DLog($"Initial state: Vol={initialVolume:E6} m?");

                    // Determine selection strategy (face or edge)
                    IFace2 largestFace = SolidWorksApiWrapper.GetLargestPlanarFace(body);
                    bool useFaceSelection = (largestFace != null);
                    IEdge longestEdge = null;

                    // Phase 1: Select geometry
                    SolidWorksApiWrapper.ClearSelection(swModel);
                    if (useFaceSelection)
                    {
                        DLog($"Phase 1: Selecting largest planar face");
                        var faceArea = 0.0;
                        try { faceArea = largestFace.GetArea(); } catch { }
                        DLog($"  Face area: {faceArea:E6} m?");
                        SelectEntityWithMark(swModel, largestFace as IEntity, 0, append: false);
                    }
                    else
                    {
                        DLog($"Phase 1: Selecting longest edge");
                        longestEdge = FindLongestLinearEdge(body);
                        if (longestEdge == null) { info.ProblemDescription = "No suitable face or edge for InsertBends"; return false; }
                        var edgeLen = GetEdgeChordLength(longestEdge);
                        DLog($"  Edge length: {edgeLen:F3} m");
                        SelectEntityWithMark(swModel, longestEdge as IEntity, 0, append: false);
                    }
                    LogSelectionState(swModel, "pre-Phase1");

                    // Phase 1: Test conversion with minimal radius
                    const double testRadius = 0.001;
                    DLog($"Phase 1: Calling InsertBends2(radius={testRadius}, K=0.5)");
                    bool phase1Success = part.InsertBends2(testRadius, string.Empty, 0.5, -1, true, 1.0, true);
                    DLog($"  InsertBends2 returned: {phase1Success}");

                    bool hasFeature = SolidWorksApiWrapper.HasSheetMetalFeature(swModel);
                    DLog($"  Has SheetMetal feature: {hasFeature}");

                    if (!phase1Success || !hasFeature)
                    {
                        var featureCount = 0;
                        try
                        {
                            var f = swModel.FirstFeature() as IFeature;
                            while (f != null) { featureCount++; f = f.GetNextFeature() as IFeature; }
                        }
                        catch { }
                        DLog($"  Total features in model: {featureCount}");
                        DLog($"  FAILURE: Phase 1 did not create sheet metal");
                        info.ProblemDescription = "Phase 1 InsertBends2 failed to create a sheet metal feature.";
                        return false;
                    }

                    // Validate volume change after Phase 1
                    double phase1Volume = SolidWorksApiWrapper.GetModelVolume(swModel);
                    double volumeRatio = initialVolume > 0 ? phase1Volume / initialVolume : 1.0;
                    double percentChange = (volumeRatio - 1.0) * 100.0;
                    DLog($"  Volume check: Initial={initialVolume:E6}, Phase1={phase1Volume:E6}");
                    DLog($"  Ratio={volumeRatio:F6}, Change={percentChange:F3}%");

                    if (volumeRatio < 0.995 || volumeRatio > 1.005)
                    {
                        DLog($"  FAILURE: Volume change exceeds ?0.5% tolerance");
                        swModel.EditUndo2(2); // Undo the failed test
                        info.ProblemDescription = $"Volume validation failed after Phase 1 (change: {percentChange:F2}%).";
                        return false;
                    }

                    // Extract thickness from the feature SolidWorks just created
                    DLog($"Phase 1: Extracting thickness from feature");
                    double derivedThickness = GetSheetMetalThicknessFromFeature(swModel);
                    DLog($"  Raw extracted: {derivedThickness} m");

                    if (derivedThickness <= 0)
                    {
                        DLog($"  Direct extraction failed, scanning features...");
                        var feat = swModel.FirstFeature() as IFeature;
                        while (feat != null)
                        {
                            string typeName = "";
                            try { typeName = feat.GetTypeName2(); } catch { }
                            if (typeName.Contains("SheetMetal"))
                            {
                                DLog($"    Found: {feat.Name} ({typeName})");
                            }
                            feat = feat.GetNextFeature() as IFeature;
                        }
                        derivedThickness = 0.001; // Fallback
                    }

                    // Determine final parameters
                    double finalThickness = info.ThicknessInMeters > 0 ? info.ThicknessInMeters : derivedThickness;
                    finalThickness = Math.Min(finalThickness, MaxSheetThicknessMeters);
                    double finalKFactor = info.KFactor > 0 ? info.KFactor : 0.5;
                    DLog($"  Final params: Thickness={finalThickness * 1000:F2}mm, K={finalKFactor:F2}");

                    // CRITICAL: Undo Phase 1 before starting Phase 2
                    DLog($"=== UNDOING PHASE 1 ===");
                    LogSelectionState(swModel, "pre-undo");
                    swModel.EditUndo2(2);
                    swModel.EditRebuild3();
                    DLog($"  Undo complete, model rebuilt");

                    bool stillHasFeature = SolidWorksApiWrapper.HasSheetMetalFeature(swModel);
                    DLog($"  Has SheetMetal after undo: {stillHasFeature} (should be false)");

                    // Phase 2: Re-query and re-select geometry (references are now stale)
                    DLog($"=== PHASE 2: RE-SELECTION ===");
                    body = SolidWorksApiWrapper.GetMainBody(swModel) as Body2;
                    if (body == null)
                    {
                        DLog($"  FAILURE: Lost main body after undo");
                        info.ProblemDescription = "Failed to get main body after undo.";
                        return false;
                    }

                    SolidWorksApiWrapper.ClearSelection(swModel);
                    if (useFaceSelection)
                    {
                        var faceForPhase2 = SolidWorksApiWrapper.GetLargestPlanarFace(body);
                        if (faceForPhase2 == null)
                        {
                            DLog($"  FAILURE: Cannot re-find planar face");
                            info.ProblemDescription = "Failed to re-find face for Phase 2.";
                            return false;
                        }
                        SelectEntityWithMark(swModel, faceForPhase2 as IEntity, 0, append: false);
                        DLog($"  Re-selected planar face for Phase 2");
                    }
                    else
                    {
                        var edgeForPhase2 = FindLongestLinearEdge(body);
                        if (edgeForPhase2 == null)
                        {
                            DLog($"  FAILURE: Cannot re-find edge");
                            info.ProblemDescription = "Failed to re-find edge for Phase 2.";
                            return false;
                        }
                        SelectEntityWithMark(swModel, edgeForPhase2 as IEntity, 0, append: false);
                        DLog($"  Re-selected edge for Phase 2");
                    }
                    LogSelectionState(swModel, "pre-Phase2");

                    // Phase 2: Apply with final parameters (VBA convention: radius = thickness)
                    DLog($"=== PHASE 2: FINAL CONVERSION ===");
                    DLog($"  Calling InsertBends2(radius={finalThickness}, K={finalKFactor})");
                    bool phase2Success = part.InsertBends2(finalThickness, string.Empty, finalKFactor, -1, true, 1.0, true);
                    DLog($"  InsertBends2 returned: {phase2Success}");

                    bool hasFinalFeature = SolidWorksApiWrapper.HasSheetMetalFeature(swModel);
                    DLog($"  Has SheetMetal feature: {hasFinalFeature}");

                    if (!phase2Success || !hasFinalFeature)
                    {
                        DLog($"  FAILURE: Phase 2 did not create sheet metal");
                        info.ProblemDescription = "Phase 2 InsertBends2 failed.";
                        return false;
                    }

                    swModel.EditRebuild3();
                    double finalVolume = SolidWorksApiWrapper.GetModelVolume(swModel);
                    DLog($"  Final volume: {finalVolume:E6} m?");
                    DLog($"  Final/Initial ratio: {finalVolume / initialVolume:F6}");

                    info.ThicknessInMeters = finalThickness;
                    info.BendRadius = finalThickness; // Per VBA convention
                    info.KFactor = finalKFactor;
                    info.IsSheetMetal = true;

                    DLog($"=== SUCCESS: TWO-PHASE CONVERSION COMPLETE ===");
                    LogMassVolume(swModel, "final");
                    return true;
                }
            }
            catch (Exception ex)
            {
                DLog($"EXCEPTION: {ex.Message}");
                DLog($"Stack: {ex.StackTrace}");
                ErrorHandler.HandleError(proc, "Exception in InsertBendsOnEdge", ex);
                try { RevertFeaturesAfter(swModel, GetLastFeature(swModel)); } catch { }
                return false;
            }
            finally
            {
                PerformanceTracker.Instance.StopTimer(proc);
                ErrorHandler.PopCallStack();
                await Task.CompletedTask;
            }
        }

        private async Task<bool> InsertBendsFromFaceAsync(IModelDoc2 swModel, ModelInfo info, CancellationToken token)
        {
            const string proc = nameof(InsertBendsFromFaceAsync);
            ErrorHandler.PushCallStack(proc);
            PerformanceTracker.Instance.StartTimer(proc);
            IFeature bookmark = null;
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, proc))
                {
                    info.ProblemDescription = "Invalid model";
                    return false;
                }

                bookmark = GetLastFeature(swModel);

                var body = SolidWorksApiWrapper.GetMainBody(swModel);
                if (body == null)
                {
                    info.ProblemDescription = "Failed to get main body";
                    ErrorHandler.HandleError(proc, info.ProblemDescription);
                    return false;
                }

                var face = SolidWorksApiWrapper.GetLargestPlanarFace(body);
                if (face == null)
                {
                    info.ProblemDescription = "Failed to get valid planar face";
                    ErrorHandler.HandleError(proc, info.ProblemDescription);
                    return false;
                }

                SolidWorksApiWrapper.ClearSelection(swModel);
                var ent = face as IEntity;
                bool selOk = ent != null && ent.Select4(false, null);
                if (!selOk)
                {
                    info.ProblemDescription = "Failed to select planar face";
                    ErrorHandler.HandleError(proc, info.ProblemDescription);
                    return false;
                }

                double thk = info.ThicknessInMeters > 0 ? info.ThicknessInMeters : 0.001;
                thk = Math.Min(thk, MaxSheetThicknessMeters);
                double kFac = info.KFactor > 0 ? info.KFactor : 0.5;

                // VBA behavior: seed with small radius first
                const double seedRadius = 0.001; // meters (1 mm)

                var featsBefore = SnapshotFeatures(swModel);
                LogSelectionState(swModel, "pre-IB-face-fallback");
                LogMassVolume(swModel, "pre-IB-face-fallback");

                bool inserted = false;
                try
                {
                    var part = swModel as IPartDoc;
                    if (part != null)
                    {
                        inserted = part.InsertBends2(seedRadius, string.Empty, kFac, -1, true, 1.0, true);
                    }
                }
                catch { inserted = false; }

                var featsAfter = SnapshotFeatures(swModel);
                LogFeatureDiff(featsBefore, featsAfter, "IB-face-fallback");

                if (!inserted || !SolidWorksApiWrapper.HasSheetMetalFeature(swModel))
                {
                    info.ProblemDescription = "InsertBends2(face) failed";
                    try { RevertFeaturesAfter(swModel, bookmark); } catch { }
                    ErrorHandler.HandleError(proc, info.ProblemDescription);
                    return false;
                }

                if (!await ValidateSheetMetalConversionAsync(swModel, info, token))
                {
                    try { RevertFeaturesAfter(swModel, bookmark); } catch { }
                    return false;
                }

                // Finalize parameters: radius = thickness (like VBA), keep derived k-factor
                info.BendRadius = thk;
                if (info.KFactor <= 0) info.KFactor = kFac;
                if (info.BendRadius > 0 && thk > 0)
                {
                    info.AddBendParameter(info.BendRadius, thk, info.KFactor);
                }

                TryApplySheetMetalParameters(swModel, new ModelInfo
                {
                    ThicknessInMeters = thk,
                    BendRadius = thk,
                    KFactor = info.KFactor > 0 ? info.KFactor : kFac
                });
                try { swModel.EditRebuild3(); } catch { }

                info.IsSheetMetal = true;

                DLog("SUCCESS: IB-face-fallback");
                LogMassVolume(swModel, "post-IB-face-fallback");
                return true;
            }
            catch (Exception ex)
            {
                info.InsertSuccessful = false;
                ErrorHandler.HandleError(proc, "Exception during face-based conversion", ex);
                return false;
            }
            finally
            {
                try { SolidWorksApiWrapper.ClearSelection(swModel); } catch { }
                PerformanceTracker.Instance.StopTimer(proc);
                ErrorHandler.PopCallStack();
                await Task.CompletedTask;
            }
        }

        // Geometry analysis helpers
        private static void TryFindThicknessFromParallelFaces(Body2 body, out double thickness, out IFace2 baseFace)
        {
            thickness = 0.0; baseFace = null;
            try
            {
                var faces = body.GetFaces() as object[];
                if (faces == null || faces.Length == 0) return;

                var planars = new List<(IFace2 face, double area, double[] normal, double[] point)>();
                foreach (var fo in faces)
                {
                    var f = fo as IFace2; if (f == null) continue;
                    var surf = f.IGetSurface() as ISurface;
                    if (surf == null) continue;
                    bool isPlane = false; try { isPlane = surf.IsPlane(); } catch { }
                    if (!isPlane) continue;
                    double area = 0.0; try { area = f.GetArea(); } catch { }
                    double[] pp = null; double[] nn = null;
                    try
                    {
                        var p = surf.PlaneParams as object;
                        var arr = p as double[];
                        if (arr != null && arr.Length >= 6)
                        {
                            nn = new[] { arr[0], arr[1], arr[2] };
                            pp = new[] { arr[3], arr[4], arr[5] };
                        }
                    }
                    catch { }
                    if (pp == null || nn == null) continue;
                    planars.Add((f, area, nn, pp));
                }
                if (planars.Count < 2) return;

                double bestScore = double.NegativeInfinity; double bestThk = 0.0; IFace2 bestBase = null;
                for (int i = 0; i < planars.Count; i++)
                {
                    for (int j = i + 1; j < planars.Count; j++)
                    {
                        var a = planars[i]; var b = planars[j];
                        double[] an = Normalize(a.normal); double[] bn = Normalize(b.normal);
                        double dot = Math.Abs(Dot(an, bn));
                        if (Math.Abs(dot - 1.0) > 0.05) continue;
                        double[] d = { b.point[0] - a.point[0], b.point[1] - a.point[1], b.point[2] - a.point[2] };
                        double dist = Math.Abs(Dot(d, an));
                        if (dist <= 0) continue;
                        double score = Math.Min(a.area, b.area) - Math.Abs(1 - dot) * 1000.0;
                        if (score > bestScore)
                        {
                            bestScore = score; bestThk = dist; bestBase = a.area >= b.area ? a.face : b.face;
                        }
                    }
                }
                if (bestThk > 0) { thickness = bestThk; baseFace = bestBase; }
            }
            catch { }
        }

        // Extended: find thickness using the best pair of parallel planar faces and compute coverage
        private static void TryFindThicknessFromParallelFacesEx(Body2 body, out double thickness, out IFace2 baseFace, out double coverage)
        {
            thickness = 0.0; baseFace = null; coverage = 0.0;
            try
            {
                var faces = body?.GetFaces() as object[]; if (faces == null || faces.Length < 2) return;
                var planars = new List<(IFace2 face, double area, double[] normal, double[] point)>();
                double totalArea = 0.0;
                foreach (var fo in faces)
                {
                    var f = fo as IFace2; if (f == null) continue;
                    double a = 0.0; try { a = f.GetArea(); } catch { }
                    totalArea += a;
                    var surf = f.IGetSurface() as ISurface; if (surf == null) continue;
                    bool isPlane = false; try { isPlane = surf.IsPlane(); } catch { }
                    if (!isPlane) continue;
                    double[] pp = null; double[] nn = null;
                    try
                    {
                        var p = surf.PlaneParams as object; var arr = p as double[];
                        if (arr != null && arr.Length >= 6)
                        { nn = new[] { arr[0], arr[1], arr[2] }; pp = new[] { arr[3], arr[4], arr[5] }; }
                    }
                    catch { }
                    if (pp == null || nn == null) continue;
                    planars.Add((f, a, nn, pp));
                }
                if (planars.Count < 2 || totalArea <= 0) return;

                double bestScore = double.NegativeInfinity; double bestThk = 0.0; IFace2 bestBase = null; double bestPairArea = 0.0;
                for (int i = 0; i < planars.Count; i++)
                {
                    for (int j = i + 1; j < planars.Count; j++)
                    {
                        var a = planars[i]; var b = planars[j];
                        double[] an = Normalize(a.normal); double[] bn = Normalize(b.normal);
                        double dot = Math.Abs(Dot(an, bn));
                        if (Math.Abs(dot - 1.0) > 0.05) continue;
                        double[] d = { b.point[0] - a.point[0], b.point[1] - a.point[1], b.point[2] - a.point[2] };
                        double dist = Math.Abs(Dot(d, an));
                        if (dist <= 0) continue;
                        double pairArea = Math.Min(a.area, b.area);
                        double score = pairArea - Math.Abs(1 - dot) * 1000.0;
                        if (score > bestScore)
                        {
                            bestScore = score; bestThk = dist; bestBase = a.area >= b.area ? a.face : b.face; bestPairArea = pairArea;
                        }
                    }
                }
                if (bestThk > 0)
                {
                    thickness = bestThk; baseFace = bestBase; coverage = Math.Max(0.0, Math.Min(1.0, bestPairArea / totalArea));
                }
            }
            catch { thickness = 0.0; baseFace = null; coverage = 0.0; }
        }

        // Selection validation
        private static bool ValidateSelection(IModelDoc2 model, int mark, int expectedType, int minCount)
        {
            try
            {
                var selMgr = model?.ISelectionManager;
                if (selMgr == null) return false;
                int cnt = selMgr.GetSelectedObjectCount2(mark);
                if (cnt < minCount) return false;
                int typeMatches = 0;
                for (int i = 1; i <= cnt; i++)
                {
                    int t = selMgr.GetSelectedObjectType3(i, mark);
                    if (t == expectedType) typeMatches++;
                }
                return typeMatches >= minCount;
            }
            catch { return false; }
        }

        // Longest linear edge on a face
        private static IEdge PickLongestLinearEdge(IFace2 face)
        {
            try
            {
                IEdge best = null;
                double bestLen = 0.0;
                var edges = face?.GetEdges() as object[];
                if (edges == null) return null;
                foreach (var o in edges)
                {
                    var e = o as IEdge; if (e == null) continue;
                    if (!IsLinearEdge(e)) continue;
                    double len = GetEdgeChordLength(e);
                    if (len > bestLen)
                    {
                        bestLen = len;
                        best = e;
                    }
                }
                return best;
            }
            catch { return null; }
        }

        private static bool IsLinearEdge(IEdge edge)
        {
            try
            {
                var c = edge?.GetCurve() as ICurve;
                if (c != null)
                {
                    var mi = c.GetType().GetMethod("IsLine");
                    if (mi != null)
                    {
                        var r = mi.Invoke(c, null);
                        if (r is bool b) return b;
                    }
                }
            }
            catch { }
            return true; // assume ok if unknown
        }

        private static double GetEdgeChordLength(IEdge edge)
        {
            try
            {
                var sv = edge.GetStartVertex() as IVertex;
                var ev = edge.GetEndVertex() as IVertex;
                var sp = sv?.GetPoint() as object;
                var ep = ev?.GetPoint() as object;
                var sa = sp as double[];
                var ea = ep as double[];
                if (sa != null && ea != null && sa.Length >= 3 && ea.Length >= 3)
                {
                    double dx = ea[0] - sa[0];
                    double dy = ea[1] - sa[1];
                    double dz = ea[2] - sa[2];
                    return Math.Sqrt(dx * dx + dy * dy + dz * dz);
                }
            }
            catch { }
            return 0.0;
        }

        // Helper method to find longest linear edge on any face of a body
        private static IEdge FindLongestLinearEdge(Body2 body)
        {
            IEdge longestEdge = null;
            double maxLength = 0.0;
            if (body == null) return null;

            var faces = body.GetFaces() as object[];
            if (faces == null) return null;

            foreach (var faceObj in faces)
            {
                var face = faceObj as IFace2;
                if (face == null) continue;

                var edges = face.GetEdges() as object[];
                if (edges == null) continue;

                foreach (var edgeObj in edges)
                {
                    var edge = edgeObj as IEdge;
                    if (edge == null || !IsLinearEdge(edge)) continue;

                    double length = GetEdgeChordLength(edge);
                    if (length > maxLength)
                    {
                        maxLength = length;
                        longestEdge = edge;
                    }
                }
            }
            return longestEdge;
        }

        // Helper to get thickness directly from the Sheet-Metal feature definition
        private static double GetSheetMetalThicknessFromFeature(IModelDoc2 model)
        {
            var feat = model?.FirstFeature() as IFeature;
            while (feat != null)
            {
                string typeName = string.Empty;
                try { typeName = feat.GetTypeName2(); } catch { }

                if (!string.IsNullOrEmpty(typeName) && typeName.IndexOf("SheetMetal", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    try
                    {
                        var def = feat.GetDefinition();
                        var prop = def?.GetType().GetProperty("Thickness");
                        if (prop != null)
                        {
                            var val = prop.GetValue(def, null);
                            double thickness = Convert.ToDouble(val);
                            if (thickness > 0) return thickness;
                        }
                    }
                    catch { }
                }
                feat = feat.GetNextFeature() as IFeature;
            }
            return 0.0; // Not found
        }

        private async Task<bool> ValidateAndSaveResultAsync(IModelDoc2 swModel, ModelInfo info, CancellationToken token)
        {
            const string proc = nameof(ValidateAndSaveResultAsync);
            ErrorHandler.PushCallStack(proc);
            try
            {
                try { SolidWorksApiWrapper.ForceRebuildDoc(swModel, SwRebuildOptions.SwForceRebuildAll); } catch { }

                if (!await ValidateSheetMetalConversionAsync(swModel, info, token))
                {
                    ErrorHandler.DebugLog($"{proc}: conversion validation failed");
                    return false;
                }

                TryApplySheetMetalParameters(swModel, info);
                try { SolidWorksApiWrapper.SetCustomProperty(swModel, "IsSheetMetal", "Yes", ""); } catch { }

                bool flattened = FlattenSheetMetalPart(swModel);
                ErrorHandler.DebugLog($"{proc}: Flatten result={flattened}");
                if (!flattened)
                {
                    info.ProblemDescription = "Flatten verification failed";
                    ErrorHandler.HandleError(proc, info.ProblemDescription, null, "Error");
                    return false;
                }

                _fileOps.SaveSWDocument(swModel, swSaveAsOptions_e.swSaveAsOptions_Silent);
                info.IsFlattened = true;
                info.InsertSuccessful = true;
                info.IsSheetMetal = true;
                ErrorHandler.DebugLog($"{proc}: SUCCESS (converted + flattened + saved)");

                return true;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        private void TryApplySheetMetalParameters(IModelDoc2 swModel, ModelInfo info)
        {
            const string proc = nameof(TryApplySheetMetalParameters);
            ErrorHandler.PushCallStack(proc);
            try
            {
                IFeature smFeat = null;
                var feat = swModel.FirstFeature() as IFeature;
                while (feat != null)
                {
                    string t = string.Empty; try { t = feat.GetTypeName2(); } catch { }
                    if (!string.IsNullOrEmpty(t) && t.IndexOf("SheetMetal", StringComparison.OrdinalIgnoreCase) >= 0)
                    { smFeat = feat; break; }
                    feat = feat.GetNextFeature() as IFeature;
                }
                if (smFeat == null) return;

                var def = smFeat.GetDefinition();
                if (def == null) return;

                if (info.ThicknessInMeters > 0)
                {
                    var pTh1 = def.GetType().GetProperty("Thickness");
                    var pTh2 = def.GetType().GetProperty("SheetThickness");
                    try { pTh1?.SetValue(def, info.ThicknessInMeters, null); } catch { }
                    try { pTh2?.SetValue(def, info.ThicknessInMeters, null); } catch { }
                }
                if (info.BendRadius > 0)
                { var pBr = def.GetType().GetProperty("BendRadius"); try { pBr?.SetValue(def, info.BendRadius, null); } catch { } }
                if (info.KFactor > 0 && info.KFactor <= 1)
                {
                    try
                    {
                        var mGetCba = def.GetType().GetMethod("GetCustomBendAllowance");
                        var cba = mGetCba != null ? mGetCba.Invoke(def, null) : null;
                        if (cba != null)
                        {
                            var pType = cba.GetType().GetProperty("Type");
                            var pK = cba.GetType().GetProperty("KFactor");
                            try { pType?.SetValue(cba, (int)swBendAllowanceTypes_e.swBendAllowanceKFactor, null); } catch { }
                            try { pK?.SetValue(cba, info.KFactor, null); } catch { }
                            var mSetCba = def.GetType().GetMethod("SetCustomBendAllowance");
                            try { mSetCba?.Invoke(def, new[] { cba }); } catch { }
                        }
                    }
                    catch { }
                }

                try
                {
                    bool modOk = smFeat.ModifyDefinition(def, swModel, null);
                    if (!modOk) ErrorHandler.DebugLog($"{proc}: ModifyDefinition returned false");
                    else swModel.EditRebuild3();
                }
                catch { }
            }
            catch { }
            finally { ErrorHandler.PopCallStack(); }
        }

        private async Task<bool> ValidateSheetMetalConversionAsync(IModelDoc2 swModel, ModelInfo info, CancellationToken token)
        {
            const string proc = nameof(ValidateSheetMetalConversionAsync);
            ErrorHandler.PushCallStack(proc);
            PerformanceTracker.Instance.StartTimer(proc);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, proc))
                { info.ProblemDescription = "Invalid model"; return false; }

                if (!SolidWorksApiWrapper.HasSheetMetalFeature(swModel))
                {
                    info.ProblemDescription = "No SheetMetal feature found after conversion";
                    ErrorHandler.HandleError(proc, info.ProblemDescription);
                    return false;
                }

                // Update thickness from the final feature if not already set
                if (info.ThicknessInMeters <= 0)
                {
                    info.ThicknessInMeters = GetSheetMetalThicknessFromFeature(swModel);
                }
                double thicknessMeters = info.ThicknessInMeters;

                // Clamp to 1 inch max for safety
                if (thicknessMeters > 0) thicknessMeters = Math.Min(thicknessMeters, MaxSheetThicknessMeters);
                if (thicknessMeters > 0) info.ThicknessInMeters = thicknessMeters;

                info.FinalVolume = SolidWorksApiWrapper.GetModelVolume(swModel);
                if (!ValidateVolumeChange(info))
                {
                    info.ProblemDescription = $"Volume validation failed: Initial={info.InitialVolume:0.000000}, Final={info.FinalVolume:0.000000}";
                    ErrorHandler.HandleError(proc, info.ProblemDescription);
                    return false;
                }

                if (info.BendRadius <= 0 && info.ThicknessInMeters > 0) info.BendRadius = info.ThicknessInMeters * 1.5;
                if (info.KFactor <= 0) info.KFactor = 0.5;
                info.IsSheetMetal = true;

                if (!NM.Core.Configuration.Logging.IsProductionMode)
                {
                    ErrorHandler.DebugLog($"{proc}: Thickness={info.ThicknessInMeters} m ({info.ThicknessInInches} in)");
                }

                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Exception during validation", ex);
                return false;
            }
            finally
            {
                PerformanceTracker.Instance.StopTimer(proc);
                ErrorHandler.PopCallStack();
                await Task.CompletedTask;
            }
        }

        private static bool ValidateVolumeChange(ModelInfo info)
        {
            const double relTol = 0.005; // 0.5%
            double a = Math.Max(0.0, info.InitialVolume);
            double b = Math.Max(0.0, info.FinalVolume);
            if (a == 0 && b == 0) return true;
            if (a == 0 || b == 0) return false; // catch zeroed results
            double relErr = Math.Abs(a - b) / Math.Max(a, b);
            return relErr <= relTol;
        }

        private static IFeature GetLastFeature(IModelDoc2 model)
        {
            IFeature f = model.FirstFeature() as IFeature; IFeature last = null;
            while (f != null) { last = f; f = f.GetNextFeature() as IFeature; }
            return last;
        }

        private static void RevertFeaturesAfter(IModelDoc2 model, IFeature bookmark)
        {
            if (model == null) return;
            var toDelete = new List<IFeature>(); bool after = bookmark == null; IFeature f = model.FirstFeature() as IFeature;
            while (f != null)
            {
                if (after) toDelete.Add(f); else if (object.ReferenceEquals(f, bookmark)) after = true;
                f = f.GetNextFeature() as IFeature;
            }
            var ext = model.Extension;
            for (int i = toDelete.Count - 1; i >= 0; i--)
            {
                try { var feat = toDelete[i]; if (feat == null) continue; feat.Select2(false, -1); ext.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed); }
                catch { }
            }
            model.EditRebuild3();
        }

        private static void SelectEntityWithMark(IModelDoc2 model, IEntity entity, int mark, bool append)
        {
            if (model == null || entity == null) return;
            var selMgr = model.ISelectionManager; SelectData sd = selMgr?.CreateSelectData(); if (sd != null) sd.Mark = mark; entity.Select4(append, sd);
        }

        private static IEnumerable<IEdge> GetCandidateRipEdges(IFace2 face, int max)
        {
            var list = new List<IEdge>(); var obj = face?.GetEdges() as object[]; if (obj == null) return list;
            foreach (var o in obj)
            { if (o is IEdge e) { list.Add(e); if (list.Count >= max) break; } }
            return list;
        }

        private double TryComputeCylinderRadius(Body2 body)
        {
            try
            {
                var faces = body.GetFaces() as object[]; if (faces == null) return 0.0;
                double cylArea = 0.0, totalArea = 0.0; double radiusSample = 0.0;
                foreach (var fo in faces)
                {
                    var f = fo as IFace2; if (f == null) continue; double a = 0.0; try { a = f.GetArea(); } catch { } totalArea += a;
                    var surf = f.IGetSurface() as ISurface; if (surf == null) continue; bool isCyl = false; try { isCyl = surf.IsCylinder(); } catch { }
                    if (!isCyl) continue; cylArea += a; try { var cpObj = surf.CylinderParams as object; var cp = cpObj as double[]; if (cp != null && cp.Length >= 7) { double r = Math.Abs(cp[6]); if (r > 0) radiusSample = Math.Max(radiusSample, r); } } catch { }
                }
                if (totalArea <= 0) return 0.0; double ratio = cylArea / totalArea; if (ratio >= 0.6 && radiusSample > 0) return radiusSample; return 0.0;
            }
            catch { return 0.0; }
        }

        private bool FlattenSheetMetalPart(IModelDoc2 swModel)
        {
            const string proc = nameof(FlattenSheetMetalPart);
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, proc)) return false;
                DLog("--- Starting FlattenSheetMetalPart ---");

                // Helper local functions for bend-state/flat-pattern checks
                bool TryGetBendState(IModelDoc2 m, out int bendState)
                {
                    bendState = -1;
                    try
                    {
                        var t = ((object)m).GetType();
                        var getBs = t.GetMethod("GetBendState");
                        if (getBs != null)
                        {
                            var curObj = getBs.Invoke(m, null);
                            bendState = Convert.ToInt32(curObj);
                            DLog($"TryGetBendState: Found bend state API. Current state: {bendState} (1=Formed, 2=Flattened)");
                            return true;
                        }
                        DLog("TryGetBendState: Bend state API not available on this model type.");
                    }
                    catch (Exception ex) { DLog($"TryGetBendState: EXCEPTION - {ex.Message}"); }
                    return false;
                }

                bool TrySetBendState(IModelDoc2 m, int target)
                {
                    try
                    {
                        var t = ((object)m).GetType();
                        var setBs = t.GetMethod("SetBendState");
                        if (setBs != null)
                        {
                            DLog($"TrySetBendState: Attempting to set bend state to {target}.");
                            var okObj = setBs.Invoke(m, new object[] { target });
                            bool ok = okObj is bool b ? b : true;
                            DLog($"TrySetBendState: API returned {ok}.");
                            return ok;
                        }
                    }
                    catch (Exception ex) { DLog($"TrySetBendState: EXCEPTION - {ex.Message}"); }
                    return false;
                }

                IFeature GetFlatPatternFeature(IModelDoc2 m)
                {
                    try
                    {
                        var feat = m.FirstFeature() as IFeature;
                        while (feat != null)
                        {
                            string tn = string.Empty; try { tn = feat.GetTypeName2(); } catch { }
                            if (!string.IsNullOrEmpty(tn) && tn.IndexOf("FlatPattern", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                DLog($"GetFlatPatternFeature: Found 'Flat-Pattern' feature: {feat.Name}");
                                return feat;
                            }
                            feat = feat.GetNextFeature() as IFeature;
                        }
                    }
                    catch (Exception ex) { DLog($"GetFlatPatternFeature: EXCEPTION - {ex.Message}"); }
                    DLog("GetFlatPatternFeature: 'Flat-Pattern' feature not found.");
                    return null;
                }

                bool IsFlatPatternUnsuppressed(IFeature fp)
                {
                    if (fp == null) return false;
                    try
                    {
                        var ret = fp.IsSuppressed2((int)swInConfigurationOpts_e.swThisConfiguration, null);
                        bool isSuppressed = Convert.ToBoolean(ret);
                        DLog($"IsFlatPatternUnsuppressed: Feature '{fp.Name}' IsSuppressed = {isSuppressed}.");
                        return !isSuppressed;
                    }
                    catch (Exception ex) { DLog($"IsFlatPatternUnsuppressed: EXCEPTION - {ex.Message}"); return false; }
                }

                bool UnsuppressFlatPattern(IFeature fp)
                {
                    if (fp == null) return false;
                    try
                    {
                        DLog($"UnsuppressFlatPattern: Attempting to unsuppress '{fp.Name}'.");
                        // Unsuppress the specific feature in all configurations for reliability
                        var ok = SolidWorksApiWrapper.UnsuppressFeature(fp, true);
                        DLog($"UnsuppressFlatPattern: result={ok}");
                        return ok;
                    }
                    catch (Exception ex) { DLog($"UnsuppressFlatPattern: EXCEPTION - {ex.Message}"); }
                    return false;
                }

                // Main logic:
                int initialBendState;
                TryGetBendState(swModel, out initialBendState);

                var flatPatternFeat = GetFlatPatternFeature(swModel);
                bool flatPatternExists = (flatPatternFeat != null);

                if (flatPatternExists && IsFlatPatternUnsuppressed(flatPatternFeat))
                {
                    DLog("--- Flattened state confirmed ---");
                    return true; // Already in desired state
                }

                // Check current state and perform necessary actions
                if (initialBendState == 1)
                {
                    DLog($"Current bend state: {initialBendState} (Formed)");
                    if (flatPatternExists)
                    {
                        DLog($"Unsuppressing 'Flat-Pattern' feature to switch to flattened state.");
                        UnsuppressFlatPattern(flatPatternFeat);
                    }
                    else
                    {
                        DLog($"No 'Flat-Pattern' feature found, forcing rebuild.");
                        swModel.EditRebuild3();
                    }
                }

                if (initialBendState != 2 && !flatPatternExists)
                {
                    DLog($"Unexpected state: {initialBendState}. Attempting to set bend state to Flattened.");
                    TrySetBendState(swModel, 2);
                }

                // Final checks
                flatPatternFeat = GetFlatPatternFeature(swModel);
                if (flatPatternFeat != null && IsFlatPatternUnsuppressed(flatPatternFeat))
                {
                    DLog("--- Successfully transitioned to flattened state ---");
                    return true;
                }

                DLog("Failed to achieve flattened state: " + (flatPatternFeat != null ? "Feature found but in wrong state." : "Feature not found."));
                return false;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Exception in FlattenSheetMetalPart", ex);
                return false;
            }
            finally
            {
                PerformanceTracker.Instance.StopTimer(proc);
                ErrorHandler.PopCallStack();
            }
        }
    }
}
