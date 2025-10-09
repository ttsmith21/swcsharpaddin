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

        public bool ConvertToSheetMetalAndOptionallyFlatten(ModelInfo info, IModelDoc2 model, bool flatten = true, ProcessingOptions options = null)
        {
            const string proc = nameof(ConvertToSheetMetalAndOptionallyFlatten);
            ErrorHandler.PushCallStack(proc);
            try
            {
                options = options ?? new ProcessingOptions();

                if (model == null) { Fail(info, "No active model"); return false; }
                if (model.GetType() != (int)swDocumentTypes_e.swDocPART) { Fail(info, "Active document is not a part"); return false; }

                // If ExternalStart identified this as tube stock, or model properties indicate a tube, skip sheet-metal ops
                try
                {
                    if (info != null && info.CustomProperties != null && info.CustomProperties.IsTube)
                    {
                        ErrorHandler.DebugLog("[0] Tube detected via ExternalStart; skipping InsertBends/flatten.");
                        info.IsSheetMetal = false;
                        info.IsFlattened = false;
                        info.InsertSuccessful = true;
                        return true;
                    }
                    // Fallback: read common property names written by external extractor (Shape/IsTube/Profile/Section)
                    var cfg = SafeGetActiveConfigName(model);
                    if (IsTubeByModelProperties(model, cfg))
                    {
                        if (info?.CustomProperties != null) info.CustomProperties.IsTube = true;
                        ErrorHandler.DebugLog("[0] Tube detected via custom properties; skipping InsertBends/flatten.");
                        info.IsSheetMetal = false;
                        info.IsFlattened = false;
                        info.InsertSuccessful = true;
                        return true;
                    }
                }
                catch { /* best-effort guard; continue if properties unavailable */ }

                var body = SolidWorksApiWrapper.GetMainBody(model);
                if (body == null) { Fail(info, "No solid body detected"); return false; }
                ErrorHandler.DebugLog("[1] Solid body OK");

                if (SolidWorksApiWrapper.HasSheetMetalFeature(model))
                {
                    string cfg = SafeGetActiveConfigName(model);
                    bool hasFlat = HasFlatPatternFeature(model);
                    ErrorHandler.DebugLog($"[2] Already sheet metal | flatten={flatten}, cfg='{cfg}', hasFlatPattern={hasFlat}");
                    if (flatten && !TryFlatten(model, info)) return false;
                    info.IsSheetMetal = true;
                    info.IsFlattened = flatten;
                    info.InsertSuccessful = true;
                    return true;
                }

                // Select reference (planar face preferred; else longest linear edge)
                SolidWorksApiWrapper.ClearSelection(model);
                var largestFace = GetLargestFace(body);
                if (largestFace == null) { Fail(info, "No faces found on body"); return false; }
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

                if (isPlane)
                {
                    if (!((largestFace as IEntity)?.Select4(false, null) ?? false)) { Fail(info, "Failed to select planar face"); return false; }
                    ErrorHandler.DebugLog("[3] Selected planar face");
                }
                else if (isCyl)
                {
                    var edge = FindLongestLinearEdge(body);
                    if (edge == null) { Fail(info, "Largest face is cylindrical but no linear edge found"); return false; }
                    if (!((edge as IEntity)?.Select4(false, null) ?? false)) { Fail(info, "Failed to select linear edge"); return false; }
                    ErrorHandler.DebugLog("[3] Selected longest linear edge");
                }
                else
                {
                    Fail(info, "Largest face not planar/cylindrical; skipping InsertBends");
                    return false;
                }

                double volBefore = SafeGetVolume(model);
                var part = model as IPartDoc; if (part == null) { Fail(info, "Not a part document"); return false; }

                // Probe with seed values to infer thickness (temporary feature)
                const double seedRadius = 0.001; // m
                const double seedK = 0.5;
                bool okSeed = part.InsertBends2(seedRadius, string.Empty, seedK, -1, true, 1.0, true);
                if (!okSeed || !SolidWorksApiWrapper.HasSheetMetalFeature(model))
                { Fail(info, "InsertBends2 (seed) failed"); return false; }
                ErrorHandler.DebugLog("[4] Seed bends created");

                double thickness = GetSheetMetalThicknessFromFeature(model);
                if (thickness <= 0) thickness = 0.001; // 1mm fallback

                // Undo probe
                try { model.EditUndo2(2); } catch { }
                try { model.EditRebuild3(); } catch { }

                // Reselect after undo
                body = SolidWorksApiWrapper.GetMainBody(model);
                if (body == null) { Fail(info, "Body lost after undo"); return false; }
                SolidWorksApiWrapper.ClearSelection(model);
                if (isPlane)
                {
                    var f2 = GetLargestFace(body); if (f2 == null) { Fail(info, "Cannot reselect planar face"); return false; }
                    if (!((f2 as IEntity)?.Select4(false, null) ?? false)) { Fail(info, "Failed to reselect planar face"); return false; }
                }
                else
                {
                    var e2 = FindLongestLinearEdge(body); if (e2 == null) { Fail(info, "Cannot reselect linear edge"); return false; }
                    if (!((e2 as IEntity)?.Select4(false, null) ?? false)) { Fail(info, "Failed to reselect linear edge"); return false; }
                }

                // Final parameters honoring options
                double defaultK = isPlane ? 0.4 : 0.5;
                double finalK = (options.KFactor > 0) ? options.KFactor : defaultK;
                double finalR = isPlane ? (thickness * 1.5) : thickness;

                // Resolve bend table with fallback and log
                string bendPath = NM.Core.BendTableResolver.Resolve(options);
                ErrorHandler.DebugLog($"[3.5] Bend table resolved: '{bendPath}'");
                if (!string.Equals(bendPath, NM.Core.Configuration.FilePaths.BendTableNone, StringComparison.OrdinalIgnoreCase))
                {
                    if (!File.Exists(bendPath))
                    {
                        ErrorHandler.HandleError("SimpleSM", $"Bend table not found: '{bendPath}'. Falling back to K-factor {finalK}.", null, "Warning");
                        bendPath = NM.Core.Configuration.FilePaths.BendTableNone;
                    }
                }

                // STEP 1: Always try K-Factor final insert first (probe/validate real params)
                bool okK = part.InsertBends2(finalR, string.Empty, finalK, -1, true, 1.0, true);
                if (!okK || !SolidWorksApiWrapper.HasSheetMetalFeature(model))
                { Fail(info, "InsertBends2 (K-Factor final) failed"); return false; }
                ErrorHandler.DebugLog($"[4] K-Factor insert OK: K={finalK}, R={finalR*1000:F2} mm");

                bool usedBendTable = false;

                // STEP 2: If a bend table is available/selected, undo K and re-apply with table
                if (!string.Equals(bendPath, NM.Core.Configuration.FilePaths.BendTableNone, StringComparison.OrdinalIgnoreCase))
                {
                    // Undo the K-factor insert
                    try { model.EditUndo2(1); } catch { }
                    try { model.EditRebuild3(); } catch { }

                    // Reselect reference
                    body = SolidWorksApiWrapper.GetMainBody(model);
                    if (body == null) { Fail(info, "Body lost after undo (bend table phase)"); return false; }
                    SolidWorksApiWrapper.ClearSelection(model);
                    if (isPlane)
                    {
                        var f3 = GetLargestFace(body); if (f3 == null) { Fail(info, "Cannot reselect planar face (bend table phase)"); return false; }
                        if (!((f3 as IEntity)?.Select4(false, null) ?? false)) { Fail(info, "Failed to reselect planar face (bend table phase)"); return false; }
                    }
                    else
                    {
                        var e3 = FindLongestLinearEdge(body); if (e3 == null) { Fail(info, "Cannot reselect linear edge (bend table phase)"); return false; }
                        if (!((e3 as IEntity)?.Select4(false, null) ?? false)) { Fail(info, "Failed to reselect linear edge (bend table phase)"); return false; }
                    }

                    // IMPORTANT: when using a bend table, pass KFactor= -1 and BendAllowance=-1
                    bool okTable = part.InsertBends2(finalR, bendPath, -1.0, -1, true, 1.0, true);
                    if (okTable && SolidWorksApiWrapper.HasSheetMetalFeature(model))
                    {
                        usedBendTable = true;
                        ErrorHandler.DebugLog($"[4] InsertBends2 via bend table: {bendPath}");
                        // Inspect and log final sheet-metal settings
                        TryLogSheetMetalBendSettings(model);
                    }
                    else
                    {
                        ErrorHandler.DebugLog("[4] Bend table insert failed; falling back to K-factor (reapply)");
                        // Re-apply K-Factor so the model remains converted
                        // Reselect again for safety
                        body = SolidWorksApiWrapper.GetMainBody(model);
                        if (body == null) { Fail(info, "Body lost during fallback to K-Factor"); return false; }
                        SolidWorksApiWrapper.ClearSelection(model);
                        if (isPlane)
                        {
                            var f4 = GetLargestFace(body); if (f4 == null) { Fail(info, "Cannot reselect planar face (fallback)"); return false; }
                            if (!((f4 as IEntity)?.Select4(false, null) ?? false)) { Fail(info, "Failed to reselect planar face (fallback)"); return false; }
                        }
                        else
                        {
                            var e4 = FindLongestLinearEdge(body); if (e4 == null) { Fail(info, "Cannot reselect linear edge (fallback)"); return false; }
                            if (!((e4 as IEntity)?.Select4(false, null) ?? false)) { Fail(info, "Failed to reselect linear edge (fallback)"); return false; }
                        }

                        bool okK2 = part.InsertBends2(finalR, string.Empty, finalK, -1, true, 1.0, true);
                        if (!okK2 || !SolidWorksApiWrapper.HasSheetMetalFeature(model))
                        { Fail(info, "InsertBends2 (K-Factor fallback) failed"); return false; }
                    }
                }

                // Volume consistency check (on final state)
                double volAfter = SafeGetVolume(model);
                if (volBefore > 0)
                {
                    double up = volBefore * 1.005, dn = volBefore * 0.995;
                    bool within = (volAfter <= up) && (volAfter >= dn);
                    ErrorHandler.DebugLog($"[5] Volume check: before={volBefore:E6}, after={volAfter:E6}, within±0.5%={within}");
                    if (!within) { try { model.EditUndo2(2); } catch { } Fail(info, "Volume validation failed (±0.5%)"); return false; }
                }

                if (flatten)
                {
                    if (!TryFlatten(model, info)) return false;
                    ErrorHandler.DebugLog("[6] Flattened OK");
                }

                info.IsSheetMetal = true;
                info.IsFlattened = flatten;
                info.InsertSuccessful = true;
                info.ThicknessInMeters = thickness;
                info.BendRadius = finalR;
                info.KFactor = finalK;
                return true;
            }
            catch (Exception ex)
            {
                Fail(info, ex.Message, ex);
                return false;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        private static bool IsTubeByModelProperties(IModelDoc2 model, string cfg)
        {
            try
            {
                // Common names our legacy extractor uses in properties
                string[] names = { "IsTube", "ISTUBE", "Is Tube", "Shape", "SHAPE", "shape", "Profile", "PROFILE", "Section", "SECTION", "Type", "TYPE", "CrossSection", "CROSSSECTION", "Cross Section", "CROSS SECTION" };
                foreach (var n in names)
                {
                    var v = SolidWorksApiWrapper.GetCustomPropertyValue(model, n, cfg);
                    if (string.IsNullOrEmpty(v)) continue;
                    if (LooksLikeTube(v)) return true;
                    // Shape present (non-empty) is enough to treat as tube
                    if (n.Equals("Shape", System.StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(v)) return true;
                }
                // Fallback to global only
                foreach (var n in names)
                {
                    var v = SolidWorksApiWrapper.GetCustomPropertyValue(model, n, "");
                    if (string.IsNullOrEmpty(v)) continue;
                    if (LooksLikeTube(v)) return true;
                    if (n.Equals("Shape", System.StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(v)) return true;
                }
            }
            catch { }
            return false;
        }

        private static bool LooksLikeTube(string value)
        {
            try
            {
                var s = (value ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(s)) return false;
                s = s.ToUpperInvariant();
                if (s == "YES" || s == "TRUE" || s == "1") return true; // for IsTube
                return s.Contains("TUBE") || s.Contains("TUBING") || s.Contains("PIPE");
            }
            catch { return false; }
        }

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
            IFace2 best = null; double bestArea = 0.0;
            try
            {
                var faces = body?.GetFaces() as object[]; if (faces == null) return null;
                foreach (var fo in faces)
                {
                    var f = fo as IFace2; if (f == null) continue;
                    double a = 0.0; try { a = f.GetArea(); } catch { }
                    if (a > bestArea) { bestArea = a; best = f; }
                }
            }
            catch { }
            return best;
        }

        private static IEdge FindLongestLinearEdge(IBody2 body)
        {
            IEdge best = null; double bestLen = 0.0;
            try
            {
                var faces = body?.GetFaces() as object[]; if (faces == null) return null;
                foreach (var fo in faces)
                {
                    var f = fo as IFace2; if (f == null) continue;
                    var edges = f.GetEdges() as object[]; if (edges == null) continue;
                    foreach (var eo in edges)
                    {
                        var e = eo as IEdge; if (e == null) continue;
                        if (!IsLinearEdge(e)) continue;
                        double len = GetEdgeChordLength(e);
                        if (len > bestLen) { bestLen = len; best = e; }
                    }
                }
            }
            catch { }
            return best;
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
            return true; // assume linear if unknown
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
                            double t = Convert.ToDouble(val);
                            if (t > 0) return t;
                        }
                    }
                    catch { }
                }
                feat = feat.GetNextFeature() as IFeature;
            }
            return 0.0;
        }

        private static double SafeGetVolume(IModelDoc2 model)
        {
            try { return SolidWorksApiWrapper.GetModelVolume(model); } catch { return 0.0; }
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
                        ErrorHandler.HandleError("SimpleSM.Flatten", "Failed to unsuppress Flat-Pattern; attempting SetBendState fallback", ex, "Warning");
                        try
                        {
                            var t = ((object)model).GetType();
                            var setBs = t.GetMethod("SetBendState");
                            ErrorHandler.DebugLog($"[6] TryFlatten: SetBendState method present={(setBs!=null)}");
                            if (setBs != null)
                            {
                                var okObj = setBs.Invoke(model, new object[] { 2 }); // 2 = flattened
                                bool ok = okObj is bool b ? b : true;
                                if (ok) { ErrorHandler.DebugLog("[6] SetBendState(Flattened) OK"); return true; }
                            }
                        }
                        catch (Exception ex2)
                        {
                            ErrorHandler.HandleError("SimpleSM.Flatten", "SetBendState fallback failed", ex2, "Warning");
                        }
                        Fail(info, "Failed to unsuppress Flat-Pattern", ex);
                        return false;
                    }
                }
                else
                {
                    // API fallback when no Flat-Pattern exists
                    try
                    {
                        var t = ((object)model).GetType();
                        var setBs = t.GetMethod("SetBendState");
                        ErrorHandler.DebugLog($"[6] TryFlatten: no Flat-Pattern; SetBendState present={(setBs!=null)}");
                        if (setBs != null)
                        {
                            var okObj = setBs.Invoke(model, new object[] { 2 }); // 2 = flattened
                            bool ok = okObj is bool b ? b : true;
                            if (ok) { ErrorHandler.DebugLog("[6] SetBendState(Flattened) OK"); return true; }
                        }
                    }
                    catch { }

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
    }
}
