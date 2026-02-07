using System;
using System.Linq;
using NM.Core;
using NM.Core.Tubes;
using NM.SwAddin.Geometry;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using static NM.Core.Constants.UnitConversions;

namespace NM.Core.Processing
{
    /// <summary>
    /// Tube processing implementation - initial geometry extraction implemented.
    /// All supporting services exist; this extracts basic tube geometry (OD, wall, length, axis).
    /// </summary>
    public class SimpleTubeProcessor
    {
        private const string LogPrefix = "[TUBE]";
        private readonly ISldWorks _swApp;

        public SimpleTubeProcessor() { }
        public SimpleTubeProcessor(ISldWorks swApp) { _swApp = swApp; }

        public string Name => "Tube";

        /// <summary>
        /// Returns true if the model appears to be a simple tube (hollow cylinder with planar end rings).
        /// </summary>
        public bool CanProcess(IModelDoc2 model)
        {
            try
            {
                var g = ExtractTubeGeometry(model);
                bool ok = g != null && g.Shape != TubeShape.None && g.WallThickness > 0 && g.Length > 0;
                ErrorHandler.DebugLog($"{LogPrefix} CanProcess ? {(ok ? "YES" : "NO")} (Shape={g?.ShapeName ?? "None"}, CrossSection={g?.CrossSection ?? "N/A"}, Wall={g?.WallThickness:F3}in, L={g?.Length:F3}in)");
                return ok;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("SimpleTubeProcessor.CanProcess", "Detection failed", ex, ErrorHandler.LogLevel.Warning);
                return false;
            }
        }

        /// <summary>
        /// Public helper for callers that need geometry without running full processing.
        /// Returns null if geometry is not detected.
        /// </summary>
        public TubeGeometry TryGetGeometry(IModelDoc2 model)
        {
            return ExtractTubeGeometry(model);
        }

        /// <summary>
        /// Uses geometry extraction; downstream services (schedule/material/cutting) to be integrated next.
        /// </summary>
        public bool Process(IModelDoc2 model, ProcessingOptions options)
        {
            return Process(null, model, options);
        }

        /// <summary>
        /// Process with ModelInfo context for custom property updates.
        /// </summary>
        public bool Process(ModelInfo info, IModelDoc2 model, ProcessingOptions options)
        {
            ErrorHandler.PushCallStack("SimpleTubeProcessor.Process");
            try
            {
                var geom = ExtractTubeGeometry(model);
                if (geom == null)
                {
                    ErrorHandler.HandleError("SimpleTubeProcessor.Process", "Not a tube (geometry not detected)", null, ErrorHandler.LogLevel.Warning);
                    return false;
                }

                ErrorHandler.DebugLog($"{LogPrefix} Shape={geom.ShapeName}, CrossSection={geom.CrossSection}, Wall={geom.WallThickness:F3}in, L={geom.Length:F3}in");

                // Write properties to ModelInfo if provided
                if (info?.CustomProperties != null)
                {
                    var inv = System.Globalization.CultureInfo.InvariantCulture;

                    // Shape and cross-section
                    info.CustomProperties.SetPropertyValue("TubeShape", geom.ShapeName, CustomPropertyType.Text);
                    info.CustomProperties.SetPropertyValue("TubeCrossSection", geom.CrossSection, CustomPropertyType.Text);

                    // Basic geometry properties
                    info.CustomProperties.SetPropertyValue("TubeOD", geom.OuterDiameter.ToString("0.###", inv), CustomPropertyType.Number);
                    info.CustomProperties.SetPropertyValue("TubeWall", geom.WallThickness.ToString("0.###", inv), CustomPropertyType.Number);
                    info.CustomProperties.SetPropertyValue("TubeLength", geom.Length.ToString("0.###", inv), CustomPropertyType.Number);
                    info.CustomProperties.SetPropertyValue("TubeCutLength", geom.CutLength.ToString("0.###", inv), CustomPropertyType.Number);
                    info.CustomProperties.SetPropertyValue("TubeHoles", geom.NumberOfHoles.ToString(), CustomPropertyType.Number);
                    info.CustomProperties.IsTube = true;

                    // Resolve pipe schedule (NPS and schedule code)
                    var pipeService = new PipeScheduleService();
                    string materialCategory = info.CustomProperties.MaterialCategory;
                    if (pipeService.TryResolveByOdAndWall(geom.OuterDiameter, geom.WallThickness, materialCategory, out string npsText, out string scheduleCode))
                    {
                        info.CustomProperties.SetPropertyValue("TubeNPS", npsText, CustomPropertyType.Text);
                        info.CustomProperties.SetPropertyValue("TubeSchedule", scheduleCode, CustomPropertyType.Text);
                        ErrorHandler.DebugLog($"{LogPrefix} Schedule resolved: NPS={npsText}, Schedule={scheduleCode}");
                    }
                    else
                    {
                        ErrorHandler.DebugLog($"{LogPrefix} Schedule not found for OD={geom.OuterDiameter:F3}, Wall={geom.WallThickness:F3}");
                    }

                    // Generate OptiMaterial code for tubes
                    string rbMaterial = info.CustomProperties.rbMaterialType;
                    if (!string.IsNullOrWhiteSpace(rbMaterial))
                    {
                        string optiMaterial = TubeMaterialCodeGenerator.Generate(rbMaterial, geom.OuterDiameter, geom.WallThickness);
                        info.CustomProperties.OptiMaterial = optiMaterial;
                        ErrorHandler.DebugLog($"{LogPrefix} OptiMaterial={optiMaterial}");
                    }

                    // Write F300_Length for material handling
                    info.CustomProperties.SetPropertyValue("F300_Length", geom.Length.ToString("0.###", inv), CustomPropertyType.Number);

                    // OP20 Routing: Work center assignment based on OD thresholds
                    // VBA logic: Route tubes to appropriate laser based on size capability
                    var (op20WorkCenter, op20SetupHours) = GetTubeWorkCenterRouting(geom.OuterDiameter, geom.Shape);
                    info.CustomProperties.SetPropertyValue("OP20", op20WorkCenter, CustomPropertyType.Text);
                    info.CustomProperties.SetPropertyValue("OP20_S", op20SetupHours.ToString("0.##", inv), CustomPropertyType.Number);
                    ErrorHandler.DebugLog($"{LogPrefix} OP20={op20WorkCenter}, OP20_S={op20SetupHours}");

                    // Set tube material prefix based on shape
                    // P. = pipe (round), T. = tube (square/rect), A. = angle/channel/beam
                    string prefix = GetTubeMaterialPrefix(geom.Shape, geom.OuterDiameter, geom.WallThickness);
                    info.CustomProperties.SetPropertyValue("TubePrefix", prefix, CustomPropertyType.Text);

                    // Generate description: "{Material} {Shape}"
                    string materialAbbrev = GetMaterialAbbreviation(info.CustomProperties.MaterialCategory);
                    string shapeDesc = GetShapeDescription(geom.Shape);
                    string description = $"{materialAbbrev} {shapeDesc}".Trim();
                    if (!string.IsNullOrEmpty(description))
                    {
                        info.CustomProperties.SetPropertyValue("TubeDescription", description, CustomPropertyType.Text);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("SimpleTubeProcessor.Process", "Tube processing failed", ex, ErrorHandler.LogLevel.Error);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        #region Geometry Extraction
        private TubeGeometry ExtractTubeGeometry(IModelDoc2 model)
        {
            const string proc = "SimpleTubeProcessor.ExtractTubeGeometry";
            ErrorHandler.PushCallStack(proc);
            PerformanceTracker.Instance.StartTimer("TubeGeometryExtraction");
            try
            {
                if (model == null) return null;

                // Use TubeGeometryExtractor for full shape detection (Round, Square, Rectangle, Angle, Channel)
                if (_swApp != null)
                {
                    var extractor = new TubeGeometryExtractor(_swApp);
                    var profile = extractor.Extract(model);

                    if (profile.Success)
                    {
                        ErrorHandler.DebugLog($"{LogPrefix} TubeGeometryExtractor: Shape={profile.ShapeName}, CrossSection={profile.CrossSection}");

                        // Map TubeProfile to TubeGeometry
                        var geom = new TubeGeometry
                        {
                            Shape = profile.Shape,
                            CrossSection = profile.CrossSection,
                            WallThickness = profile.WallThicknessInches,
                            Length = profile.MaterialLengthInches,
                            CutLength = profile.CutLengthInches,
                            NumberOfHoles = profile.NumberOfHoles,
                            OuterDiameter = profile.OuterDiameterInches,
                            InnerDiameter = profile.InnerDiameterInches,
                            Axis = ComputeAxisFromPoints(profile.StartPoint, profile.EndPoint)
                        };

                        return geom;
                    }
                    else
                    {
                        ErrorHandler.DebugLog($"{LogPrefix} TubeGeometryExtractor failed: {profile.Message}");

                        // If TubeGeometryExtractor explicitly rejected the part (face geometry validation,
                        // wall ratio check, etc.), do NOT fall through to legacy - the part is NOT a tube.
                        // Only fall back to legacy if the extractor was unavailable.
                        return null;
                    }
                }

                // Legacy fallback: concentric cylinder detection (round tubes only)
                // This only runs if _swApp == null (TubeGeometryExtractor wasn't available)
                return ExtractRoundTubeGeometryLegacy(model);
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Geometry extraction failed", ex, ErrorHandler.LogLevel.Warning);
                return null;
            }
            finally
            {
                PerformanceTracker.Instance.StopTimer("TubeGeometryExtraction");
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Computes normalized axis direction from start and end points.
        /// </summary>
        private static double[] ComputeAxisFromPoints(double[] startPt, double[] endPt)
        {
            if (startPt == null || endPt == null || startPt.Length < 3 || endPt.Length < 3)
                return new[] { 0.0, 0.0, 1.0 }; // Default Z-axis

            var dx = endPt[0] - startPt[0];
            var dy = endPt[1] - startPt[1];
            var dz = endPt[2] - startPt[2];
            var len = Math.Sqrt(dx * dx + dy * dy + dz * dz);

            if (len < 1e-9)
                return new[] { 0.0, 0.0, 1.0 }; // Default Z-axis

            return new[] { dx / len, dy / len, dz / len };
        }

        /// <summary>
        /// Legacy round tube extraction using concentric cylinder detection.
        /// Used as fallback when TubeGeometryExtractor is unavailable or fails.
        /// </summary>
        private TubeGeometry ExtractRoundTubeGeometryLegacy(IModelDoc2 model)
        {
            var part = model as IPartDoc;
            if (part == null) return null;

            var bodies = part.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
            if (bodies == null || bodies.Length == 0) return null;
            var body = bodies[0] as IBody2;
            if (body == null) return null;

            var faces = body.GetFaces() as object[];
            if (faces == null || faces.Length == 0) return null;

            // Collect cylinder data with axis origin for concentricity check
            double maxR_m = 0.0, minR_m = double.MaxValue;
            double[] axis = null; // normalized
            double[] axisOrigin = null;
            int cylCount = 0;

            // Tolerance for concentricity check: 1mm (0.001m) - cylinders must share same axis line
            const double CONCENTRICITY_TOLERANCE_M = 0.001;

            foreach (var f in faces)
            {
                var face = f as IFace2; if (face == null) continue;
                var surf = face.IGetSurface(); if (surf == null) continue;
                if (surf.IsCylinder())
                {
                    var p = surf.CylinderParams as double[];
                    if (p == null || p.Length < 7) continue;
                    var a = new[] { p[3], p[4], p[5] };
                    var len = Math.Sqrt(a[0] * a[0] + a[1] * a[1] + a[2] * a[2]);
                    if (len < 1e-9) continue;
                    var n = new[] { a[0] / len, a[1] / len, a[2] / len };
                    var r = Math.Abs(p[6]);
                    var origin = new[] { p[0], p[1], p[2] };

                    // Initialize axis with first cylinder
                    if (axis == null)
                    {
                        axis = n;
                        axisOrigin = origin;
                    }
                    else
                    {
                        // Check direction alignment (must be parallel)
                        var dp = Math.Abs(axis[0] * n[0] + axis[1] * n[1] + axis[2] * n[2]);
                        if (dp < 0.999) continue; // Not parallel, skip this cylinder

                        // CONCENTRICITY CHECK: Verify this cylinder's origin lies on the reference axis line
                        var toOrigin = Sub(origin, axisOrigin);
                        var projection = Dot(toOrigin, axis);
                        var projectedPoint = new[] { axisOrigin[0] + projection * axis[0], axisOrigin[1] + projection * axis[1], axisOrigin[2] + projection * axis[2] };
                        var distanceVec = Sub(origin, projectedPoint);
                        var distanceToAxis = Math.Sqrt(Dot(distanceVec, distanceVec));

                        if (distanceToAxis > CONCENTRICITY_TOLERANCE_M)
                        {
                            ErrorHandler.DebugLog($"{LogPrefix} Legacy: Skipping non-concentric cylinder: distance={distanceToAxis * 1000:F2}mm from axis");
                            continue;
                        }
                    }

                    if (r > maxR_m) maxR_m = r;
                    if (r < minR_m) minR_m = r;
                    cylCount++;
                }
            }

            if (cylCount == 0 || axis == null) return null;

            // Estimate length: project all vertices onto axis and compute span
            double minT = double.MaxValue, maxT = double.MinValue;
            var vertsObj = body.GetVertices() as object[];
            if (vertsObj == null || vertsObj.Length == 0)
            {
                // Fallback: use face vertices
                foreach (var f in faces)
                {
                    var face = f as IFace2; if (face == null) continue;
                    var loops = face.GetLoops() as object[]; if (loops == null) continue;
                    foreach (var l in loops)
                    {
                        var loop = l as ILoop2; if (loop == null) continue;
                        var edges = loop.GetEdges() as object[]; if (edges == null) continue;
                        foreach (var e in edges)
                        {
                            var edge = e as IEdge; if (edge == null) continue;
                            var sv = edge.IGetStartVertex();
                            var ev = edge.IGetEndVertex();
                            foreach (var vv in new[] { sv, ev })
                            {
                                if (vv == null) continue;
                                var pt = vv.GetPoint() as double[]; if (pt == null || pt.Length < 3) continue;
                                var t = Dot(Sub(pt, axisOrigin), axis);
                                if (t < minT) minT = t;
                                if (t > maxT) maxT = t;
                            }
                        }
                    }
                }
            }
            else
            {
                foreach (var v in vertsObj)
                {
                    var vv = v as IVertex; if (vv == null) continue;
                    var pt = vv.GetPoint() as double[]; if (pt == null || pt.Length < 3) continue;
                    var t = Dot(Sub(pt, axisOrigin), axis);
                    if (t < minT) minT = t;
                    if (t > maxT) maxT = t;
                }
            }

            double length_m = (maxT > minT && maxT < double.MaxValue && minT > double.MinValue) ? (maxT - minT) : 0.0;

            // Require two cylinders for hollow tube; otherwise it's a solid rod
            if (cylCount < 2) return null;

            var od_m = 2 * maxR_m;
            var wall_m = (maxR_m - minR_m);

            // Minimum wall thickness: 0.015" (0.381mm = 0.000381m) - thin-wall tubing minimum
            const double MIN_WALL_THICKNESS_M = 0.000381;
            if (od_m <= 0 || wall_m <= MIN_WALL_THICKNESS_M || length_m <= 0) return null;

            // Convert to inches for TubeGeometry DTO
            var geom = new TubeGeometry
            {
                Shape = TubeShape.Round,
                OuterDiameter = od_m * MetersToInches,
                InnerDiameter = (2 * minR_m) * MetersToInches,
                WallThickness = wall_m * MetersToInches,
                Length = length_m * MetersToInches,
                Axis = axis,
                CrossSection = $"{od_m * MetersToInches:G6}"
            };
            return geom;
        }

        private static double[] Sub(double[] a, double[] b) => new[] { a[0] - b[0], a[1] - b[1], a[2] - b[2] };
        private static double Dot(double[] a, double[] b) => a[0] * b[0] + a[1] * b[1] + a[2] * b[2];
        #endregion

        #region Work Center Routing
        /// <summary>
        /// Determines the OP20 work center and setup hours based on tube OD.
        /// VBA logic from SP.bas:2520-2560:
        /// - OD ≤ 6" → F110 TUBE LASER (setup 0.15)
        /// - OD 6-10" → F110 (setup 0.5)
        /// - OD 10-10.75" → F110 (setup 1.0)
        /// - OD > 10.75" → N145 5-AXIS LASER (setup 0.25)
        /// </summary>
        private static (string WorkCenter, double SetupHours) GetTubeWorkCenterRouting(double outerDiameterIn, TubeShape shape)
        {
            // For non-round shapes, route to F110 with shape-specific setup
            if (shape != TubeShape.Round)
            {
                // Structural shapes (angle, channel) need more setup than tube shapes (rect, square)
                double setupHrs = (shape == TubeShape.Angle || shape == TubeShape.Channel) ? 0.25 : 0.15;
                return ("F110 - TUBE LASER", setupHrs);
            }

            // Round pipe/tube routing based on OD thresholds
            // Small epsilon (0.05") handles floating point from metric-to-imperial conversion
            if (outerDiameterIn <= 6.05)
            {
                return ("F110 - TUBE LASER", 0.15);
            }
            else if (outerDiameterIn <= 10.05)
            {
                return ("F110 - TUBE LASER", 0.5);
            }
            else if (outerDiameterIn <= 10.80)
            {
                return ("F110 - TUBE LASER", 1.0);
            }
            else
            {
                // Large pipes go to 5-axis laser
                return ("N145 - 5-AXIS LASER", 0.25);
            }
        }

        /// <summary>
        /// Determines tube material prefix based on shape.
        /// VBA logic:
        /// - P. = pipe (round, hollow, wall &lt; 30% of OD)
        /// - T. = tube (square/rectangular)
        /// - A. = angle/channel/beam
        /// </summary>
        private static string GetTubeMaterialPrefix(TubeShape shape, double outerDiameterIn, double wallThicknessIn)
        {
            switch (shape)
            {
                case TubeShape.Round:
                    // Distinguish pipe from thick-wall tube
                    // Pipe: wall < 30% of OD (typical for pipe schedules)
                    double wallRatio = (outerDiameterIn > 0) ? wallThicknessIn / outerDiameterIn : 0;
                    return (wallRatio < 0.30) ? "P." : "T.";

                case TubeShape.Square:
                case TubeShape.Rectangle:
                    return "T.";

                case TubeShape.Angle:
                case TubeShape.Channel:
                    return "A.";

                default:
                    return "";
            }
        }

        /// <summary>
        /// Gets material abbreviation for description.
        /// </summary>
        private static string GetMaterialAbbreviation(string materialCategory)
        {
            if (string.IsNullOrWhiteSpace(materialCategory))
                return "";

            var cat = materialCategory.ToUpperInvariant();
            if (cat.Contains("STAINLESS") || cat.Contains("304") || cat.Contains("316"))
                return "SS";
            if (cat.Contains("ALUMINUM") || cat.Contains("AL"))
                return "AL";
            if (cat.Contains("CARBON") || cat.Contains("STEEL") || cat.Contains("1018") || cat.Contains("A36"))
                return "CS";
            if (cat.Contains("GALV"))
                return "GALV";

            return "CS"; // Default to carbon steel
        }

        /// <summary>
        /// Gets shape description for tube type.
        /// </summary>
        private static string GetShapeDescription(TubeShape shape)
        {
            switch (shape)
            {
                case TubeShape.Round: return "PIPE";
                case TubeShape.Square: return "SQ TUBE";
                case TubeShape.Rectangle: return "RECT TUBE";
                case TubeShape.Angle: return "ANGLE";
                case TubeShape.Channel: return "CHANNEL";
                default: return "TUBE";
            }
        }
        #endregion
    }

    /// <summary>
    /// Data structure for tube geometry - ready for use once extraction works.
    /// Units: inches for OD/Wall/Length.
    /// </summary>
    public class TubeGeometry
    {
        public double OuterDiameter { get; set; }   // inches (for round tubes)
        public double WallThickness { get; set; }   // inches
        public double Length { get; set; }          // inches (material length)
        public double[] Axis { get; set; }          // normalized direction vector [x, y, z]

        /// <summary>
        /// Detected shape (Round, Square, Rectangle, Angle, Channel).
        /// </summary>
        public TubeShape Shape { get; set; } = TubeShape.None;

        /// <summary>
        /// Cross-section description (e.g., "2.5" for round OD, "2 x 3" for rectangular).
        /// </summary>
        public string CrossSection { get; set; } = "";

        /// <summary>
        /// Total cut perimeter length in inches.
        /// </summary>
        public double CutLength { get; set; }

        /// <summary>
        /// Number of holes detected.
        /// </summary>
        public int NumberOfHoles { get; set; }

        /// <summary>
        /// Inner diameter for round tubes, in inches.
        /// </summary>
        public double InnerDiameter { get; set; }

        /// <summary>
        /// Shape name as string.
        /// </summary>
        public string ShapeName => Shape.ToString();
    }
}
