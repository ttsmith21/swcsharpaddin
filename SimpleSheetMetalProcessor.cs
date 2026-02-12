using System;
using System.IO;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin
{
    /// <summary>
    /// Minimal, step-by-step sheet metal converter for easy debugging and validation.
    /// Honors ProcessingOptions for Bend Table vs K-Factor.
    /// </summary>
    public sealed class SimpleSheetMetalProcessor
    {
        private readonly ISldWorks _swApp;

        public SimpleSheetMetalProcessor(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Converts a solid body to sheet metal using the VBA two-phase approach:
        /// PHASE 1: Probe insert → extract thickness → volume check → flatten → mass check
        /// PHASE 2: Undo probe → final insert (bend table first, K-factor fallback) → flatten
        /// </summary>
        public bool ConvertToSheetMetalAndOptionallyFlatten(ModelInfo info, IModelDoc2 model, bool flatten = true, ProcessingOptions options = null)
        {
            const string proc = nameof(ConvertToSheetMetalAndOptionallyFlatten);
            ErrorHandler.PushCallStack(proc);
            ErrorHandler.DebugLog("[SMDBG] ====== ConvertToSheetMetalAndOptionallyFlatten() ENTER (VBA Two-Phase) ======");
            ErrorHandler.DebugLog($"[SMDBG] Parameters: model={(model != null ? "valid" : "NULL")}, info={(info != null ? "valid" : "NULL")}, flatten={flatten}");
            try
            {
                options = options ?? new ProcessingOptions();
                ErrorHandler.DebugLog($"[SMDBG] ProcessingOptions: KFactor={options.KFactor}, BendTable='{options.BendTable ?? "null"}'");

                // ========================================
                // PRE-CHECKS
                // ========================================
                if (model == null)
                {
                    ErrorHandler.DebugLog("[SMDBG] FAIL: model is NULL");
                    Fail(info, "No active model");
                    return false;
                }

                int modelType = model.GetType();
                if (modelType != (int)swDocumentTypes_e.swDocPART)
                {
                    ErrorHandler.DebugLog("[SMDBG] FAIL: Not a part document");
                    Fail(info, "Active document is not a part");
                    return false;
                }

                var part = model as IPartDoc;
                if (part == null)
                {
                    ErrorHandler.DebugLog("[SMDBG] FAIL: Could not cast model to IPartDoc");
                    Fail(info, "Not a part document");
                    return false;
                }

                ErrorHandler.DebugLog("[SMDBG] VBA approach: Try sheet metal first, let geometry decide");
                var body = SwGeometryHelper.GetMainBody(model);
                if (body == null)
                {
                    ErrorHandler.DebugLog("[SMDBG] FAIL: No solid body detected");
                    Fail(info, "No solid body detected");
                    return false;
                }
                ErrorHandler.DebugLog("[SMDBG] [1] Solid body OK");

                // Check if already sheet metal
                if (SwGeometryHelper.HasSheetMetalFeature(model))
                {
                    ErrorHandler.DebugLog("[SMDBG] [2] Already sheet metal");

                    // Extract thickness from existing SM features (critical for re-runs)
                    double existingThickness = GetSheetMetalThicknessViaLinkedProperty(model);
                    if (existingThickness <= 0)
                        existingThickness = GetSheetMetalThicknessFromFeatureDefinition(model);
                    ErrorHandler.DebugLog($"[SMDBG] Existing SM thickness: {existingThickness:E6} m ({existingThickness * 39.37:F4} inches)");

                    if (flatten && !TryFlatten(model, info))
                    {
                        ErrorHandler.DebugLog("[SMDBG] FAIL: TryFlatten returned false");
                        return false;
                    }
                    info.IsSheetMetal = true;
                    info.IsFlattened = flatten;
                    info.InsertSuccessful = true;
                    if (existingThickness > 0)
                        info.ThicknessInMeters = existingThickness;
                    ErrorHandler.DebugLog("[SMDBG] SUCCESS: Part already had sheet metal features");
                    return true;
                }

                // ========================================
                // SELECT REFERENCE (planar face or linear edge)
                // ========================================
                ErrorHandler.DebugLog("[SMDBG] Part needs conversion - selecting reference...");
                SolidWorksApiWrapper.ClearSelection(model);
                var largestFace = GetLargestFace(body);
                if (largestFace == null)
                {
                    ErrorHandler.DebugLog("[SMDBG] FAIL: No faces found on body");
                    Fail(info, "No faces found on body");
                    return false;
                }

                bool isPlane = false, isCyl = false;
                try
                {
                    var surf = largestFace.IGetSurface() as ISurface;
                    if (surf != null)
                    {
                        try { isPlane = surf.IsPlane(); } catch { }
                        try { isCyl = surf.IsCylinder(); } catch { }
                    }
                }
                catch { }

                if (!SelectReference(model, body, largestFace, isPlane, isCyl, info))
                    return false;

                double volBefore = SafeGetVolume(model);
                ErrorHandler.DebugLog($"[SMDBG] Volume before: {volBefore:E6} m³");

                // ========================================
                // PHASE 1: PROBE INSERT (extract thickness, validate)
                // ========================================
                ErrorHandler.DebugLog("[SMDBG] === PHASE 1: PROBE INSERT ===");
                const double seedRadius = 0.001; // 1mm - VBA uses 0.001
                const double seedK = 0.5;

                ErrorHandler.DebugLog($"[SMDBG] Calling InsertBends2 (PROBE): radius={seedRadius}m, K={seedK}");
                PerformanceTracker.Instance.StartTimer("InsertBends2_Probe");
                bool okProbe = part.InsertBends2(seedRadius, string.Empty, seedK, -1, true, 1.0, true);
                PerformanceTracker.Instance.StopTimer("InsertBends2_Probe");

                if (!okProbe || !SwGeometryHelper.HasSheetMetalFeature(model))
                {
                    ErrorHandler.DebugLog("[SMDBG] FAIL: Probe InsertBends2 failed - part cannot be converted");
                    Fail(info, "InsertBends2 probe failed - part may not be convertible to sheet metal");
                    return false;
                }
                ErrorHandler.DebugLog("[SMDBG] Probe insert successful");

                // Extract thickness from probe (MUST succeed)
                double thickness = GetSheetMetalThicknessViaLinkedProperty(model);
                if (thickness <= 0)
                {
                    thickness = GetSheetMetalThicknessFromFeatureDefinition(model);
                }
                ErrorHandler.DebugLog($"[SMDBG] Probe thickness: {thickness:E6} m ({thickness * 39.37:F4} inches)");

                if (thickness <= 0)
                {
                    ErrorHandler.DebugLog("[SMDBG] FAIL: Could not extract thickness from probe");
                    try { model.EditUndo2(2); } catch { }
                    CleanupFailedConversion(model);
                    Fail(info, "Could not extract sheet metal thickness from probe feature");
                    return false;
                }

                // Volume check on probe (±3% - VBA tolerance)
                double volAfterProbe = SafeGetVolume(model);
                ErrorHandler.DebugLog($"[SMDBG] Volume after probe: {volAfterProbe:E6} m³");
                if (volBefore > 0)
                {
                    double up = volBefore * 1.03, dn = volBefore * 0.97;
                    bool withinVol = (volAfterProbe <= up) && (volAfterProbe >= dn);
                    ErrorHandler.DebugLog($"[SMDBG] Probe volume check (±3%): {withinVol}");
                    if (!withinVol)
                    {
                        ErrorHandler.DebugLog("[SMDBG] FAIL: Probe volume check failed (±3%)");
                        try { model.EditUndo2(2); } catch { }
                        CleanupFailedConversion(model);
                        Fail(info, "Probe volume check failed (±3%)");
                        return false;
                    }
                }

                // Flatten probe and do mass comparison (±3%)
                ErrorHandler.DebugLog("[SMDBG] Flattening probe for mass comparison...");
                PerformanceTracker.Instance.StartTimer("TryFlatten_Probe");
                bool probeFlattened = TryFlatten(model, info);
                PerformanceTracker.Instance.StopTimer("TryFlatten_Probe");
                if (!probeFlattened)
                {
                    ErrorHandler.DebugLog("[SMDBG] FAIL: Could not flatten probe");
                    try { model.EditUndo2(2); } catch { }
                    CleanupFailedConversion(model);
                    Fail(info, "Could not flatten probe for validation");
                    return false;
                }

                // Mass comparison: blankArea × thickness ≈ actualVolume (±3%)
                body = SwGeometryHelper.GetMainBody(model);
                double biggestAreaM2;
                GetLargestFace(body, out biggestAreaM2);
                if (biggestAreaM2 > 0 && thickness > 0)
                {
                    double calcVolumeM3 = biggestAreaM2 * thickness;
                    double actualVolumeM3 = SafeGetVolume(model);

                    // VBA: swVolumeUP = swVolume * 1.03, swVolumeDN = swVolume * 0.97
                    double swVolumeUP = actualVolumeM3 * 1.03;
                    double swVolumeDN = actualVolumeM3 * 0.97;
                    bool bVolumeUP = swVolumeUP > calcVolumeM3;
                    bool bVolumeDN = swVolumeDN < calcVolumeM3;
                    bool withinMass = bVolumeUP && bVolumeDN;

                    double percentDiff = actualVolumeM3 > 0 ? Math.Abs(calcVolumeM3 - actualVolumeM3) / actualVolumeM3 * 100 : 100;
                    ErrorHandler.DebugLog($"[SMDBG] Mass comparison: calcVol={calcVolumeM3:E6}, actualVol={actualVolumeM3:E6}, diff={percentDiff:F1}%, pass={withinMass}");

                    if (!withinMass)
                    {
                        ErrorHandler.DebugLog("[SMDBG] FAIL: Mass comparison failed (±3%)");
                        try { model.EditUndo2(2); } catch { }
                        CleanupFailedConversion(model);
                        Fail(info, $"Mass validation failed (±3%): {percentDiff:F1}% difference");
                        return false;
                    }
                }

                // ========================================
                // PHASE 2: UNDO PROBE, DO FINAL INSERT
                // ========================================
                ErrorHandler.DebugLog("[SMDBG] === PHASE 2: FINAL INSERT ===");
                ErrorHandler.DebugLog("[SMDBG] Undoing probe...");
                try { model.EditUndo2(2); } catch { }
                try { model.EditRebuild3(); } catch { }

                // Reselect reference after undo
                body = SwGeometryHelper.GetMainBody(model);
                if (body == null)
                {
                    ErrorHandler.DebugLog("[SMDBG] FAIL: Body lost after undo");
                    Fail(info, "Body lost after undo");
                    return false;
                }
                SolidWorksApiWrapper.ClearSelection(model);
                if (!SelectReference(model, body, null, isPlane, isCyl, info))
                    return false;

                // Calculate final parameters
                double defaultK = isPlane ? 0.4 : 0.5;
                double finalK = (options.KFactor > 0) ? options.KFactor : defaultK;
                double finalR = isPlane ? (thickness * 1.5) : thickness;
                ErrorHandler.DebugLog($"[SMDBG] Final params: K={finalK}, R={finalR * 1000:F3}mm");

                // Resolve bend table
                string bendPath = NM.Core.BendTableResolver.Resolve(options);
                bool hasBendTable = !string.IsNullOrEmpty(bendPath)
                    && !string.Equals(bendPath, NM.Core.Configuration.FilePaths.BendTableNone, StringComparison.OrdinalIgnoreCase)
                    && File.Exists(bendPath);
                ErrorHandler.DebugLog($"[SMDBG] Bend table: '{bendPath}', exists={hasBendTable}");

                // FINAL INSERT: Try bend table FIRST (VBA order), K-factor as fallback
                bool finalOk = false;
                bool usedBendTable = false;

                if (hasBendTable)
                {
                    // With bend table: K=-1, BA=-1 (table provides values)
                    ErrorHandler.DebugLog($"[SMDBG] Trying final insert with bend table...");
                    PerformanceTracker.Instance.StartTimer("InsertBends2_Final_BendTable");
                    finalOk = part.InsertBends2(finalR, bendPath, -1.0, -1, true, 1.0, true);
                    PerformanceTracker.Instance.StopTimer("InsertBends2_Final_BendTable");
                    if (finalOk && SwGeometryHelper.HasSheetMetalFeature(model))
                    {
                        usedBendTable = true;
                        ErrorHandler.DebugLog($"[SMDBG] Final insert with bend table: SUCCESS");
                        TryLogSheetMetalBendSettings(model);
                    }
                    else
                    {
                        ErrorHandler.DebugLog($"[SMDBG] Bend table insert failed, will try K-factor");
                        finalOk = false;
                        // Need to reselect for K-factor attempt
                        body = SwGeometryHelper.GetMainBody(model);
                        if (body != null)
                        {
                            SolidWorksApiWrapper.ClearSelection(model);
                            SelectReference(model, body, null, isPlane, isCyl, null); // Don't fail on reselect
                        }
                    }
                }

                if (!finalOk)
                {
                    // K-factor insert (primary if no bend table, fallback if table failed)
                    ErrorHandler.DebugLog($"[SMDBG] Final insert with K-factor: K={finalK}");
                    PerformanceTracker.Instance.StartTimer("InsertBends2_Final_KFactor");
                    finalOk = part.InsertBends2(finalR, string.Empty, finalK, -1, true, 1.0, true);
                    PerformanceTracker.Instance.StopTimer("InsertBends2_Final_KFactor");
                    if (finalOk)
                    {
                        ErrorHandler.DebugLog($"[SMDBG] Final insert with K-factor: SUCCESS");
                    }
                }

                if (!finalOk || !SwGeometryHelper.HasSheetMetalFeature(model))
                {
                    ErrorHandler.DebugLog("[SMDBG] FAIL: Final InsertBends2 failed - undoing broken features");
                    try { model.EditUndo2(2); } catch { }
                    try { model.EditRebuild3(); } catch { }
                    CleanupFailedConversion(model);
                    Fail(info, "Final InsertBends2 failed");
                    return false;
                }

                // ========================================
                // PHASE 3: FINAL FLATTEN
                // ========================================
                if (flatten)
                {
                    ErrorHandler.DebugLog("[SMDBG] Final flatten...");
                    PerformanceTracker.Instance.StartTimer("TryFlatten_Final");
                    bool finalFlattened = TryFlatten(model, info);
                    PerformanceTracker.Instance.StopTimer("TryFlatten_Final");
                    if (!finalFlattened)
                    {
                        ErrorHandler.DebugLog("[SMDBG] FAIL: Final TryFlatten failed - undoing InsertBends2");
                        try { model.EditUndo2(2); } catch { }
                        try { model.EditRebuild3(); } catch { }
                        CleanupFailedConversion(model);
                        return false;
                    }
                    ErrorHandler.DebugLog("[SMDBG] Final flatten: SUCCESS");
                }

                // Store results
                info.IsSheetMetal = true;
                info.IsFlattened = flatten;
                info.InsertSuccessful = true;
                info.ThicknessInMeters = thickness;
                info.BendRadius = finalR;
                info.KFactor = usedBendTable ? -1 : finalK; // -1 indicates bend table was used

                ErrorHandler.DebugLog("[SMDBG] ====== ConvertToSheetMetalAndOptionallyFlatten() SUCCESS ======");
                ErrorHandler.DebugLog($"[SMDBG] Final: thickness={thickness * 39.37:F4}in, R={finalR * 1000:F2}mm, usedBendTable={usedBendTable}");
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[SMDBG] EXCEPTION: {ex.GetType().Name}: {ex.Message}");
                Fail(info, ex.Message, ex);
                return false;
            }
            finally
            {
                ErrorHandler.DebugLog("[SMDBG] <<< ConvertToSheetMetalAndOptionallyFlatten() EXIT");
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Selects the reference entity (planar face or linear edge) for InsertBends.
        /// VBA: Set swSelMgr = swPart.SelectionManager
        ///      Set swSelData = swSelMgr.CreateSelectData
        ///      bRet = swEnt.Select4(True, swSelData)
        /// </summary>
        private bool SelectReference(IModelDoc2 model, IBody2 body, IFace2 largestFace, bool isPlane, bool isCyl, ModelInfo info)
        {
            // VBA creates SelectionData object for proper selection context
            // Note: Select4 expects SelectData (concrete), not ISelectData (interface)
            SelectData selData = null;
            try
            {
                var selMgr = model.SelectionManager as ISelectionMgr;
                if (selMgr != null)
                {
                    selData = selMgr.CreateSelectData() as SelectData;
                }
            }
            catch { }

            if (isPlane)
            {
                var face = largestFace ?? GetLargestFace(body);
                if (face == null)
                {
                    if (info != null) Fail(info, "Cannot find planar face");
                    return false;
                }
                // VBA: bRet = swEnt.Select4(True, swSelData) - Append=TRUE, use SelectionData
                bool ok = (face as IEntity)?.Select4(true, selData) ?? false;
                if (!ok)
                {
                    if (info != null) Fail(info, "Failed to select planar face");
                    return false;
                }
                ErrorHandler.DebugLog("[SMDBG] Selected planar face (with SelectionData)");
                return true;
            }
            else if (isCyl)
            {
                var edge = FindLongestLinearEdge(body);
                if (edge == null)
                {
                    if (info != null) Fail(info, "Cannot find linear edge for cylindrical face");
                    return false;
                }
                // VBA: bRet = swEnt.Select4(True, swSelData) - Append=TRUE, use SelectionData
                bool ok = (edge as IEntity)?.Select4(true, selData) ?? false;
                if (!ok)
                {
                    if (info != null) Fail(info, "Failed to select linear edge");
                    return false;
                }
                ErrorHandler.DebugLog("[SMDBG] Selected linear edge (with SelectionData)");
                return true;
            }
            else
            {
                if (info != null) Fail(info, "Largest face not planar/cylindrical");
                return false;
            }
        }

        // GetLargestFaceArea removed — use GetLargestFace(body, out area) instead to avoid redundant face scan

        // NOTE: IsTubeByModelProperties and LooksLikeTube methods were REMOVED
        // VBA approach: Try InsertBends first, let geometry decide - don't pre-classify from properties
        // Pre-classification from stale/contaminated properties caused incorrect results

        private static void TryLogSheetMetalBendSettings(IModelDoc2 model)
        {
            try
            {
                var feat = model?.FirstFeature() as IFeature;
                IFeature lastSm = null;
                while (feat != null)
                {
                    string t = string.Empty;
                    try { t = feat.GetTypeName2(); } catch { }
                    if (!string.IsNullOrEmpty(t) && t.IndexOf("SheetMetal", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        lastSm = feat; // keep last
                    }
                    feat = feat.GetNextFeature() as IFeature;
                }
                if (lastSm == null) { ErrorHandler.DebugLog("[SMDBG] No SheetMetal feature found for inspection"); return; }
                var def = lastSm.GetDefinition();
                if (def == null) { ErrorHandler.DebugLog("[SMDBG] SheetMetal.GetDefinition returned null"); return; }

                var defType = def.GetType();
                int bendType = -1;
                string tablePath = null;
                try
                {
                    var propType = defType.GetProperty("BendAllowanceType");
                    if (propType != null)
                    {
                        var val = propType.GetValue(def, null);
                        if (val is int i) bendType = i; else bendType = Convert.ToInt32(val);
                    }
                }
                catch { }
                try
                {
                    var propTable = defType.GetProperty("BendTable");
                    if (propTable != null)
                    {
                        var val = propTable.GetValue(def, null);
                        tablePath = val as string;
                    }
                }
                catch { }

                ErrorHandler.DebugLog($"[SMDBG] BendAllowanceType={bendType}, BendTable='{tablePath ?? ""}'");
            }
            catch { }
        }

        private static string SafeGetActiveConfigName(IModelDoc2 model)
        {
            try { return model?.ConfigurationManager?.ActiveConfiguration?.Name ?? string.Empty; } catch { return string.Empty; }
        }

        private static bool HasFlatPatternFeature(IModelDoc2 model)
        {
            try
            {
                var feat = model.FirstFeature() as IFeature;
                while (feat != null)
                {
                    string t = string.Empty; try { t = feat.GetTypeName2(); } catch { }
                    if (!string.IsNullOrEmpty(t) && t.IndexOf("FlatPattern", StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                    feat = feat.GetNextFeature() as IFeature;
                }
            }
            catch { }
            return false;
        }

        private static IFace2 GetLargestFace(IBody2 body)
        {
            return GetLargestFace(body, out _);
        }

        private static IFace2 GetLargestFace(IBody2 body, out double largestArea)
        {
            PerformanceTracker.Instance.StartTimer("GetLargestFace");
            IFace2 best = null; double bestArea = 0.0;
            try
            {
                var faces = body?.GetFaces() as object[]; if (faces == null) { largestArea = 0; return null; }
                foreach (var fo in faces)
                {
                    var f = fo as IFace2; if (f == null) continue;
                    double a = 0.0; try { a = f.GetArea(); } catch { }
                    if (a > bestArea) { bestArea = a; best = f; }
                }
            }
            catch { }
            finally
            {
                PerformanceTracker.Instance.StopTimer("GetLargestFace");
            }
            largestArea = bestArea;
            return best;
        }

        /// <summary>
        /// Find the longest linear edge on the body.
        /// VBA equivalent: GetLinearEdge() in SP.bas
        /// - Gets edges directly from body (not from faces)
        /// - Checks swCurve.Identity = 3001 (line)
        /// - Uses GetEndParams + GetLength3 for accurate length
        /// Optimized: non-linear edges exit after 2 COM calls (GetCurve + Identity).
        /// </summary>
        private static IEdge FindLongestLinearEdge(IBody2 body)
        {
            PerformanceTracker.Instance.StartTimer("FindLongestLinearEdge");
            IEdge best = null;
            double bestLen = 0.0;

            try
            {
                var edgesRaw = body?.GetEdges() as object[];
                if (edgesRaw == null || edgesRaw.Length == 0)
                {
                    ErrorHandler.DebugLog("[SMDBG] FindLongestLinearEdge: No edges on body");
                    return null;
                }

                ErrorHandler.DebugLog($"[SMDBG] FindLongestLinearEdge: Scanning {edgesRaw.Length} edges");
                int linearCount = 0;

                foreach (var eo in edgesRaw)
                {
                    var edge = eo as IEdge;
                    if (edge == null) continue;

                    var curve = edge.GetCurve() as ICurve;
                    if (curve == null) continue;

                    // Early exit: skip non-linear edges (2 COM calls only)
                    // NOTE: ICurve.IsLine() only exists in SW 2024+, so use Identity() for 2022 compatibility
                    try
                    {
                        if (curve.Identity() != 3001) continue; // 3001 = LINE_TYPE
                    }
                    catch { continue; }

                    linearCount++;

                    // Only linear edges reach here: get params + length
                    // Uses GetEndParams (returns primitives via out params) instead of
                    // GetCurveParams3 (allocates COM object per call) — matches FaceWrapper.cs pattern
                    double edgeLength = 0;
                    try
                    {
                        double start = 0, end = 0;
                        bool isClosed = false, isPeriodic = false;
                        if (curve.GetEndParams(out start, out end, out isClosed, out isPeriodic))
                        {
                            edgeLength = curve.GetLength3(start, end);
                        }
                    }
                    catch { }

                    if (edgeLength <= 0)
                        edgeLength = GetEdgeChordLength(edge);

                    if (edgeLength > bestLen)
                    {
                        bestLen = edgeLength;
                        best = edge;
                    }
                }

                ErrorHandler.DebugLog($"[SMDBG] FindLongestLinearEdge: {linearCount} linear of {edgesRaw.Length} total. Best = {bestLen:F6} m");
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[SMDBG] FindLongestLinearEdge exception: {ex.Message}");
            }
            finally
            {
                PerformanceTracker.Instance.StopTimer("FindLongestLinearEdge");
            }

            return best;
        }

        private static double GetEdgeChordLength(IEdge edge)
        {
            try
            {
                var sv = edge.GetStartVertex() as IVertex;
                var ev = edge.GetEndVertex() as IVertex;
                var sa = (sv?.GetPoint() as object) as double[];
                var ea = (ev?.GetPoint() as object) as double[];
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

        /// <summary>
        /// Gets sheet metal thickness using VBA-style linked property approach.
        /// Creates a temporary custom property with equation "Thickness@$PRP:SW-File Name.SLDPRT"
        /// which SolidWorks evaluates to the actual thickness dimension.
        /// </summary>
        private static double GetSheetMetalThicknessViaLinkedProperty(IModelDoc2 model)
        {
            const string tempPropName = "SMThick_Temp";
            const string proc = "GetSheetMetalThicknessViaLinkedProperty";
            double thickness = 0.0;

            try
            {
                // Get configuration name (VBA uses gstrConfigName)
                string cfg = string.Empty;
                try { cfg = model?.ConfigurationManager?.ActiveConfiguration?.Name ?? string.Empty; }
                catch { cfg = string.Empty; }

                // Delete any existing temp property first
                try { model.DeleteCustomInfo2(cfg, tempPropName); } catch { }
                try { model.DeleteCustomInfo(tempPropName); } catch { }

                // VBA formula: """Thickness@$PRP:""SW-File Name"".SLDPRT"""
                // This creates a linked value that evaluates to the sheet metal thickness
                string linkedFormula = "\"Thickness@$PRP:\"\"SW-File Name\"\".SLDPRT\"";

                // Add linked property (type 30 = swCustomInfoText in VBA)
                // Try config-specific first, then file-level
                bool added = false;
                try
                {
                    if (!string.IsNullOrEmpty(cfg))
                    {
                        var ext = model?.Extension;
                        var cpm = ext?.CustomPropertyManager[cfg];
                        if (cpm != null)
                        {
                            int result = cpm.Add3(tempPropName, (int)swCustomInfoType_e.swCustomInfoText, linkedFormula, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                            added = (result == 0 || result == 1);
                            ErrorHandler.DebugLog($"[{proc}] AddCustomInfo3 (config='{cfg}') result={result}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandler.DebugLog($"[{proc}] Config property exception: {ex.Message}");
                }

                if (!added)
                {
                    // Fallback to file-level property
                    try
                    {
                        var ext = model?.Extension;
                        var cpm = ext?.CustomPropertyManager[""];
                        if (cpm != null)
                        {
                            int result = cpm.Add3(tempPropName, (int)swCustomInfoType_e.swCustomInfoText, linkedFormula, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                            added = (result == 0 || result == 1);
                            ErrorHandler.DebugLog($"[{proc}] AddCustomInfo3 (file-level) result={result}");
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorHandler.DebugLog($"[{proc}] File-level property exception: {ex.Message}");
                    }
                }

                // Read the resolved value
                string thicknessStr = null;
                try
                {
                    if (!string.IsNullOrEmpty(cfg))
                    {
                        var ext = model?.Extension;
                        var cpm = ext?.CustomPropertyManager[cfg];
                        if (cpm != null)
                        {
                            string valOut = null, resolvedOut = null;
                            bool wasResolved = false;
                            cpm.Get5(tempPropName, true, out valOut, out resolvedOut, out wasResolved);
                            thicknessStr = resolvedOut ?? valOut;
                            ErrorHandler.DebugLog($"[{proc}] Get5 (config): val='{valOut}', resolved='{resolvedOut}', wasResolved={wasResolved}");
                        }
                    }
                }
                catch { }

                if (string.IsNullOrEmpty(thicknessStr))
                {
                    try
                    {
                        var ext = model?.Extension;
                        var cpm = ext?.CustomPropertyManager[""];
                        if (cpm != null)
                        {
                            string valOut = null, resolvedOut = null;
                            bool wasResolved = false;
                            cpm.Get5(tempPropName, true, out valOut, out resolvedOut, out wasResolved);
                            thicknessStr = resolvedOut ?? valOut;
                            ErrorHandler.DebugLog($"[{proc}] Get5 (file-level): val='{valOut}', resolved='{resolvedOut}', wasResolved={wasResolved}");
                        }
                    }
                    catch { }
                }

                // VBA check: strThicknessCheck = UCase(Right(strThickness, 7)) ... If strThicknessCheck <> badThickness
                // The "bad" value contains "SLDPRT"" at the end, meaning the linked value wasn't resolved
                if (!string.IsNullOrEmpty(thicknessStr))
                {
                    string check = thicknessStr.Length >= 7 ? thicknessStr.Substring(thicknessStr.Length - 7).ToUpperInvariant() : "";
                    if (check.Contains("SLDPRT"))
                    {
                        ErrorHandler.DebugLog($"[{proc}] Linked property not resolved (contains SLDPRT): '{thicknessStr}'");
                    }
                    else
                    {
                        // VBA: Thickness = strThickness * 0.0254  (inches to meters)
                        if (double.TryParse(thicknessStr, out double thickInches))
                        {
                            thickness = thickInches * 0.0254; // Convert inches to meters
                            ErrorHandler.DebugLog($"[{proc}] Parsed thickness: {thickInches} in = {thickness} m");
                        }
                        else
                        {
                            ErrorHandler.DebugLog($"[{proc}] Could not parse thickness value: '{thicknessStr}'");
                        }
                    }
                }

                // Cleanup: delete temp property
                try { model.DeleteCustomInfo2(cfg, tempPropName); } catch { }
                try { model.DeleteCustomInfo(tempPropName); } catch { }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[{proc}] Exception: {ex.Message}");
            }

            // If linked property approach failed, fall back to feature definition approach
            if (thickness <= 0)
            {
                ErrorHandler.DebugLog($"[{proc}] Linked property failed, trying feature definition...");
                thickness = GetSheetMetalThicknessFromFeatureDefinition(model);
            }

            return thickness;
        }

        /// <summary>
        /// Gets thickness directly from sheet metal feature definition.
        /// VBA approach: objSheetMetal.Thickness (direct property access, no reflection)
        /// </summary>
        private static double GetSheetMetalThicknessFromFeatureDefinition(IModelDoc2 model)
        {
            const string proc = "GetSheetMetalThicknessFromFeatureDefinition";
            try
            {
                var feat = model?.FirstFeature() as IFeature;
                while (feat != null)
                {
                    string typeName = string.Empty;
                    try { typeName = feat.GetTypeName2(); } catch { }

                    // VBA checks: If objFeature.GetTypeName2 = "SheetMetal" Then (exact match)
                    if (string.Equals(typeName, "SheetMetal", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            var def = feat.GetDefinition();
                            if (def != null)
                            {
                                // VBA approach: direct cast and property access
                                // Set objSheetMetal = objFeature.GetDefinition
                                // dblThickness = objSheetMetal.Thickness
                                var smDef = def as ISheetMetalFeatureData;
                                if (smDef != null)
                                {
                                    double t = smDef.Thickness;
                                    ErrorHandler.DebugLog($"[{proc}] ISheetMetalFeatureData.Thickness = {t} m ({t * 1000:F3} mm)");
                                    if (t > 0) return t;

                                    // VBA returns -1 on failure, we return 0 (caller will fail)
                                    ErrorHandler.DebugLog($"[{proc}] WARNING: Thickness property returned {t} (invalid)");
                                }
                                else
                                {
                                    // Cast failed - this is unexpected, log it
                                    ErrorHandler.DebugLog($"[{proc}] WARNING: Could not cast to ISheetMetalFeatureData (def type: {def.GetType().Name})");
                                }
                                // NOTE: Reflection fallback REMOVED - if direct cast fails, we should fail
                                // Using reflection was masking problems and could return wrong values
                            }
                        }
                        catch (Exception ex)
                        {
                            ErrorHandler.DebugLog($"[{proc}] Feature definition exception: {ex.Message}");
                        }
                    }
                    feat = feat.GetNextFeature() as IFeature;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[{proc}] Exception: {ex.Message}");
            }
            return 0.0;
        }

        private static double SafeGetVolume(IModelDoc2 model)
        {
            try { return SwMassPropertiesHelper.GetModelVolume(model); } catch { return 0.0; }
        }

        /// <summary>
        /// Cleans up residual sheet metal features after a failed conversion.
        /// EditUndo2(2) doesn't always fully undo InsertBends on STEP-imported parts,
        /// leaving FlatPattern features unsuppressed and SheetMetal warnings.
        /// </summary>
        private static void CleanupFailedConversion(IModelDoc2 model)
        {
            if (model == null) return;

            try
            {
                // 1. If FlatPattern is unsuppressed (part shows flat), suppress it
                var bendMgr = new NM.SwAddin.SheetMetal.BendStateManager();
                if (bendMgr.IsFlattened(model))
                {
                    bendMgr.UnFlattenPart(model); // suppress flat pattern
                    ErrorHandler.DebugLog("[SMDBG] CleanupFailedConversion: Suppressed FlatPattern");
                }

                // 2. If SheetMetal features still remain after undo, try additional undo
                if (SwGeometryHelper.HasSheetMetalFeature(model))
                {
                    try { model.EditUndo2(1); } catch { }
                    ErrorHandler.DebugLog("[SMDBG] CleanupFailedConversion: Extra undo for residual SM features");
                }

                // 3. Force rebuild to clean state
                try { model.ForceRebuild3(false); } catch { }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[SMDBG] CleanupFailedConversion exception: {ex.Message}");
            }
        }

        private static void Fail(ModelInfo info, string reason, Exception ex = null)
        {
            if (info != null) info.ProblemDescription = reason;
            ErrorHandler.HandleError("SimpleSM", reason, ex);
        }

        private static bool TryFlatten(IModelDoc2 model, ModelInfo info)
        {
            try
            {
                string cfg = SafeGetActiveConfigName(model);
                ErrorHandler.DebugLog($"[6] TryFlatten: start, cfg='{cfg}'");
                var feat = model.FirstFeature() as IFeature;
                IFeature flat = null;
                int scanned = 0;
                while (feat != null)
                {
                    scanned++;
                    string t = string.Empty; string n = string.Empty;
                    try { t = feat.GetTypeName2(); } catch { }
                    try { n = feat.Name; } catch { }
                    if (!string.IsNullOrEmpty(t) && t.IndexOf("FlatPattern", StringComparison.OrdinalIgnoreCase) >= 0)
                    { flat = feat; break; }
                    feat = feat.GetNextFeature() as IFeature;
                }
                ErrorHandler.DebugLog($"[6] TryFlatten: featuresScanned={scanned}, flatFound={(flat!=null)}");

                if (flat != null)
                {
                    try
                    {
                        flat.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature,
                                             (int)swInConfigurationOpts_e.swAllConfiguration, null);
                        model.EditRebuild3();
                        ErrorHandler.DebugLog("[6] Flat-Pattern unsuppressed");
                        return true;
                    }
                    catch (Exception ex)
                    {
                        // If unsuppress fails, try SetBendState fallback before failing
                        ErrorHandler.HandleError("SimpleSM.Flatten", "Failed to unsuppress Flat-Pattern; attempting SetBendState fallback", ex, ErrorHandler.LogLevel.Warning);
                        if (TrySetBendStateFlattened(model)) return true;
                        Fail(info, "Failed to unsuppress Flat-Pattern", ex);
                        return false;
                    }
                }
                else
                {
                    // API fallback when no Flat-Pattern exists
                    TrySetBendStateFlattened(model);

                    ErrorHandler.DebugLog("[6] No Flat-Pattern; skipping flatten");
                    return true;
                }
            }
            catch (Exception ex)
            {
                Fail(info, "Exception while flattening", ex);
                return false;
            }
        }

        /// <summary>
        /// Attempts to flatten the model using reflection-based SetBendState(2) call.
        /// This is a fallback when FlatPattern feature manipulation fails or doesn't exist.
        /// </summary>
        private static bool TrySetBendStateFlattened(IModelDoc2 model)
        {
            try
            {
                var t = ((object)model).GetType();
                var setBs = t.GetMethod("SetBendState");
                if (setBs != null)
                {
                    var okObj = setBs.Invoke(model, new object[] { 2 }); // 2 = flattened
                    bool ok = okObj is bool b ? b : true;
                    if (ok) { ErrorHandler.DebugLog("[SheetMetal] SetBendState(Flattened) OK"); return true; }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("SimpleSM.Flatten", "SetBendState fallback failed", ex, ErrorHandler.LogLevel.Warning);
            }
            return false;
        }
    }
}
