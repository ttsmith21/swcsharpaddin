using System;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using static NM.Core.Constants.UnitConversions;

namespace NM.SwAddin.Validation
{
    /// <summary>
    /// Checks whether a top-level part or assembly exceeds the shop's
    /// handling capacity (weight or physical size).
    /// Triggers when:
    ///   - Total mass > 10,000 lbs, OR
    ///   - 2 of 3 bounding-box dimensions exceed 12 feet.
    /// </summary>
    public static class OversizedPartCheck
    {
        public const double MAX_WEIGHT_LBS = 10_000.0;
        public const double MAX_DIMENSION_FEET = 12.0;

        private const double MAX_DIMENSION_M = MAX_DIMENSION_FEET / MetersToFeet; // ~3.6576 m

        public sealed class OversizedResult
        {
            public bool IsOversized { get; set; }
            public string Reason { get; set; }
            public double MassLbs { get; set; }
            public double DimXFeet { get; set; }
            public double DimYFeet { get; set; }
            public double DimZFeet { get; set; }
        }

        /// <summary>
        /// Runs the oversized check on the top-level document (part or assembly).
        /// </summary>
        public static OversizedResult Check(ISldWorks swApp, IModelDoc2 model)
        {
            const string proc = nameof(OversizedPartCheck) + ".Check";
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (model == null)
                    return new OversizedResult { IsOversized = false };

                int docType = model.GetType();
                if (docType != (int)swDocumentTypes_e.swDocPART &&
                    docType != (int)swDocumentTypes_e.swDocASSEMBLY)
                    return new OversizedResult { IsOversized = false };

                // --- Mass ---
                double massKg = GetMassKg(model);
                double massLbs = massKg * KgToLbs;

                // --- Bounding box (meters) ---
                double dimX = 0, dimY = 0, dimZ = 0;
                if (docType == (int)swDocumentTypes_e.swDocPART)
                    GetPartBoundingBox(model, out dimX, out dimY, out dimZ);
                else
                    GetAssemblyBoundingBox(swApp, model, out dimX, out dimY, out dimZ);

                double dimXFt = dimX * MetersToFeet;
                double dimYFt = dimY * MetersToFeet;
                double dimZFt = dimZ * MetersToFeet;

                // --- Evaluate criteria ---
                bool overWeight = massLbs > MAX_WEIGHT_LBS;

                int oversizeDimCount = 0;
                if (dimXFt > MAX_DIMENSION_FEET) oversizeDimCount++;
                if (dimYFt > MAX_DIMENSION_FEET) oversizeDimCount++;
                if (dimZFt > MAX_DIMENSION_FEET) oversizeDimCount++;
                bool overSize = oversizeDimCount >= 2;

                var result = new OversizedResult
                {
                    MassLbs = massLbs,
                    DimXFeet = dimXFt,
                    DimYFeet = dimYFt,
                    DimZFeet = dimZFt
                };

                if (overWeight && overSize)
                {
                    result.IsOversized = true;
                    result.Reason = $"Exceeds handling capacity: {massLbs:N0} lbs (max {MAX_WEIGHT_LBS:N0}) " +
                                    $"AND {oversizeDimCount} dimensions exceed {MAX_DIMENSION_FEET}' " +
                                    $"({dimXFt:F1}' x {dimYFt:F1}' x {dimZFt:F1}') — may require outside crane";
                }
                else if (overWeight)
                {
                    result.IsOversized = true;
                    result.Reason = $"Exceeds weight capacity: {massLbs:N0} lbs (max {MAX_WEIGHT_LBS:N0}) " +
                                    $"— may require outside crane";
                }
                else if (overSize)
                {
                    result.IsOversized = true;
                    result.Reason = $"Exceeds size capacity: {oversizeDimCount} of 3 dimensions exceed {MAX_DIMENSION_FEET}' " +
                                    $"({dimXFt:F1}' x {dimYFt:F1}' x {dimZFt:F1}') — may require outside crane";
                }

                if (result.IsOversized)
                    ErrorHandler.DebugLog($"[OVERSIZE-HANDLING] {result.Reason}");

                return result;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Oversized check failed", ex, ErrorHandler.LogLevel.Warning);
                return new OversizedResult { IsOversized = false };
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        private static double GetMassKg(IModelDoc2 model)
        {
            try
            {
                var massProp = model.Extension?.CreateMassProperty() as IMassProperty;
                return massProp?.Mass ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Computes the 3D bounding box of a part by unioning all solid body boxes.
        /// Dimensions returned in meters.
        /// </summary>
        private static void GetPartBoundingBox(IModelDoc2 model, out double dx, out double dy, out double dz)
        {
            dx = dy = dz = 0;
            var part = model as IPartDoc;
            if (part == null) return;

            var bodiesRaw = part.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
            if (bodiesRaw == null || bodiesRaw.Length == 0) return;

            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;

            foreach (IBody2 body in bodiesRaw)
            {
                var box = body.GetBodyBox() as double[];
                if (box == null || box.Length < 6) continue;
                UnionBox(box, ref minX, ref minY, ref minZ, ref maxX, ref maxY, ref maxZ);
            }

            if (minX < double.MaxValue)
            {
                dx = maxX - minX;
                dy = maxY - minY;
                dz = maxZ - minZ;
            }
        }

        /// <summary>
        /// Computes the 3D bounding box of an assembly by traversing top-level
        /// components, transforming their body boxes to assembly space, and
        /// computing the union.
        /// Dimensions returned in meters.
        /// </summary>
        private static void GetAssemblyBoundingBox(ISldWorks swApp, IModelDoc2 model,
            out double dx, out double dy, out double dz)
        {
            dx = dy = dz = 0;

            var config = model.ConfigurationManager?.ActiveConfiguration;
            if (config == null) return;
            var rootComp = config.GetRootComponent3(true) as IComponent2;
            if (rootComp == null) return;

            var mathUtil = swApp.GetMathUtility() as IMathUtility;
            if (mathUtil == null) return;

            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;

            TraverseComponentBounds(rootComp, mathUtil, ref minX, ref minY, ref minZ, ref maxX, ref maxY, ref maxZ);

            if (minX < double.MaxValue)
            {
                dx = maxX - minX;
                dy = maxY - minY;
                dz = maxZ - minZ;
            }
        }

        /// <summary>
        /// Recursively traverses components, getting body boxes and
        /// transforming to assembly coordinates.
        /// </summary>
        private static void TraverseComponentBounds(IComponent2 comp, IMathUtility mathUtil,
            ref double minX, ref double minY, ref double minZ,
            ref double maxX, ref double maxY, ref double maxZ)
        {
            var children = comp.GetChildren() as object[];
            if (children == null) return;

            foreach (IComponent2 child in children)
            {
                try
                {
                    if (child.IsSuppressed()) continue;

                    var xform = child.Transform2 as MathTransform;
                    var childDoc = child.GetModelDoc2() as IModelDoc2;

                    // If this child has its own children, recurse (sub-assembly)
                    var grandChildren = child.GetChildren() as object[];
                    if (grandChildren != null && grandChildren.Length > 0)
                    {
                        TraverseComponentBounds(child, mathUtil,
                            ref minX, ref minY, ref minZ, ref maxX, ref maxY, ref maxZ);
                        continue;
                    }

                    // Leaf component — get its body boxes
                    if (childDoc == null) continue;
                    var partDoc = childDoc as IPartDoc;
                    if (partDoc == null) continue;

                    var bodiesRaw = partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
                    if (bodiesRaw == null) continue;

                    foreach (IBody2 body in bodiesRaw)
                    {
                        var box = body.GetBodyBox() as double[];
                        if (box == null || box.Length < 6) continue;

                        if (xform != null)
                        {
                            // Transform all 8 corners of the box to assembly space
                            TransformBoxCorners(mathUtil, xform, box,
                                ref minX, ref minY, ref minZ, ref maxX, ref maxY, ref maxZ);
                        }
                        else
                        {
                            UnionBox(box, ref minX, ref minY, ref minZ, ref maxX, ref maxY, ref maxZ);
                        }
                    }
                }
                catch
                {
                    // Skip components that fail — don't let one bad component block the check
                }
            }
        }

        /// <summary>
        /// Transforms the 8 corners of an axis-aligned box through a component
        /// transform and updates the running min/max.
        /// </summary>
        private static void TransformBoxCorners(IMathUtility mathUtil, MathTransform xform,
            double[] box,
            ref double minX, ref double minY, ref double minZ,
            ref double maxX, ref double maxY, ref double maxZ)
        {
            // 8 corners of the AABB
            double[] xs = { box[0], box[3] };
            double[] ys = { box[1], box[4] };
            double[] zs = { box[2], box[5] };

            foreach (double x in xs)
            foreach (double y in ys)
            foreach (double z in zs)
            {
                var ptArr = new double[] { x, y, z };
                var pt = mathUtil.CreatePoint(ptArr) as IMathPoint;
                if (pt == null) continue;
                var tPt = pt.MultiplyTransform(xform) as IMathPoint;
                if (tPt == null) continue;
                var arr = tPt.ArrayData as double[];
                if (arr == null || arr.Length < 3) continue;

                if (arr[0] < minX) minX = arr[0];
                if (arr[1] < minY) minY = arr[1];
                if (arr[2] < minZ) minZ = arr[2];
                if (arr[0] > maxX) maxX = arr[0];
                if (arr[1] > maxY) maxY = arr[1];
                if (arr[2] > maxZ) maxZ = arr[2];
            }
        }

        private static void UnionBox(double[] box,
            ref double minX, ref double minY, ref double minZ,
            ref double maxX, ref double maxY, ref double maxZ)
        {
            if (box[0] < minX) minX = box[0];
            if (box[1] < minY) minY = box[1];
            if (box[2] < minZ) minZ = box[2];
            if (box[3] > maxX) maxX = box[3];
            if (box[4] > maxY) maxY = box[4];
            if (box[5] > maxZ) maxZ = box[5];
        }
    }
}
