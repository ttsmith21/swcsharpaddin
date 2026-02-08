using System;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin
{
    /// <summary>
    /// Body, face, and edge geometry operations for SolidWorks models.
    /// Extracted from SolidWorksApiWrapper for single-responsibility.
    /// </summary>
    public static class SwGeometryHelper
    {
        /// <summary>
        /// Gets the first solid body from a part document.
        /// </summary>
        public static IBody2 GetMainBody(IModelDoc2 swModel)
        {
            const string procName = "GetMainBody";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return null;

                var swPart = swModel as IPartDoc;
                if (swPart == null)
                {
                    ErrorHandler.HandleError(procName, "Model is not a part document.");
                    return null;
                }

                var vBodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
                if (vBodies == null || vBodies.Length == 0)
                {
                    ErrorHandler.HandleError(procName, "Part contains no solid bodies.");
                    return null;
                }

                return vBodies[0] as IBody2;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Error getting main body.", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Gets the largest planar face from a given body.
        /// </summary>
        public static IFace2 GetLargestPlanarFace(IBody2 swBody)
        {
            const string procName = "GetLargestPlanarFace";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (swBody == null)
                {
                    ErrorHandler.HandleError(procName, "Input body is null.");
                    return null;
                }

                var vFaces = swBody.GetFaces() as object[];
                if (vFaces == null || vFaces.Length == 0)
                {
                    ErrorHandler.HandleError(procName, "Body contains no faces.");
                    return null;
                }

                IFace2 largestFace = null;
                double maxArea = 0;

                foreach (object face in vFaces)
                {
                    IFace2 swFace = face as IFace2;
                    if (swFace == null) continue;

                    ISurface swSurf = swFace.GetSurface() as ISurface;
                    if (swSurf != null && swSurf.IsPlane())
                    {
                        double currentArea = swFace.GetArea();
                        if (currentArea > maxArea)
                        {
                            maxArea = currentArea;
                            largestFace = swFace;
                        }
                    }
                }

                if (largestFace == null)
                {
                    ErrorHandler.HandleError(procName, "No planar faces found in body.", null, ErrorHandler.LogLevel.Warning);
                }

                return largestFace;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Error getting largest planar face.", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Checks if the model contains a SheetMetal or FlatPattern feature.
        /// </summary>
        public static bool HasSheetMetalFeature(IModelDoc2 swModel)
        {
            const string procName = "HasSheetMetalFeature";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return false;

                var body = GetMainBody(swModel);
                if (body != null)
                {
                    var bodyFeatures = body.GetFeatures() as object[];
                    if (bodyFeatures != null)
                    {
                        foreach (var featObj in bodyFeatures)
                        {
                            var feat = featObj as IFeature;
                            if (feat == null) continue;
                            string featType = feat.GetTypeName2();
                            if (string.Equals(featType, "SheetMetal", StringComparison.OrdinalIgnoreCase))
                            {
                                return true;
                            }
                        }
                    }
                }

                // Fallback: also check model feature tree for FlatPattern (may not be on body)
                var swFeat = swModel.FirstFeature() as IFeature;
                while (swFeat != null)
                {
                    string featType = swFeat.GetTypeName2();
                    if (featType.IndexOf("SheetMetal", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        featType.IndexOf("FlatPattern", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }
                    swFeat = swFeat.GetNextFeature() as IFeature;
                }
                return false;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Error checking for sheet metal features.", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Gets the longest linear edge from a given body.
        /// </summary>
        public static IEdge GetLongestLinearEdge(IBody2 swBody)
        {
            const string procName = "GetLongestLinearEdge";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (swBody == null)
                {
                    ErrorHandler.HandleError(procName, "Input body is null.");
                    return null;
                }

                var vEdges = swBody.GetEdges() as object[];
                if (vEdges == null || vEdges.Length == 0)
                {
                    ErrorHandler.HandleError(procName, "Body contains no edges.");
                    return null;
                }

                IEdge longestEdge = null;
                double maxLength = 0;

                foreach (object edge in vEdges)
                {
                    IEdge swEdge = edge as IEdge;
                    if (swEdge == null) continue;

                    ICurve swCurve = swEdge.GetCurve() as ICurve;
                    // NOTE: ICurve.IsLine() only exists in SW 2024+, use Identity() for 2022 compatibility
                    if (swCurve != null && swCurve.Identity() == 3001) // 3001 = LINE_TYPE
                    {
                        var curveParams = swEdge.GetCurveParams2() as double[];
                        double length = swCurve.GetLength2(curveParams[0], curveParams[1]);
                        if (length > maxLength)
                        {
                            maxLength = length;
                            longestEdge = swEdge;
                        }
                    }
                }

                if (longestEdge == null)
                {
                    ErrorHandler.HandleError(procName, "No linear edges found in body.", null, ErrorHandler.LogLevel.Warning);
                }

                return longestEdge;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Error getting longest linear edge.", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Gets the fixed face from a model (for sheet metal processing).
        /// Returns the first planar face found, or null.
        /// </summary>
        public static IFace2 GetFixedFace(IModelDoc2 model)
        {
            if (model == null) return null;
            try
            {
                var part = model as IPartDoc;
                if (part == null) return null;

                var bodies = part.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
                if (bodies == null || bodies.Length == 0) return null;

                foreach (var b in bodies)
                {
                    var body = b as IBody2;
                    if (body == null) continue;
                    var faces = body.GetFaces() as object[];
                    if (faces == null) continue;
                    foreach (var f in faces)
                    {
                        var face = f as IFace2;
                        if (face == null) continue;
                        var surf = face.IGetSurface();
                        if (surf != null && surf.IsPlane())
                        {
                            return face;
                        }
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
    }
}
