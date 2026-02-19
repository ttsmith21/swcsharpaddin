using System;
using System.Collections.Generic;
using System.Linq;
using NM.Core;
using NM.Core.Drawing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Drawing
{
    /// <summary>
    /// Validates drawings for quality issues:
    /// - Dangling dimensions (detected via color comparison)
    /// - Bend lines without associated dimensions
    /// </summary>
    public sealed class DrawingValidator
    {
        private readonly ISldWorks _swApp;

        public DrawingValidator(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Runs all validation checks on a drawing document.
        /// </summary>
        /// <param name="drawDoc">The drawing document to validate.</param>
        /// <param name="flatPatternView">Optional flat pattern view for bend line checking.</param>
        /// <returns>Validation result with all issues found.</returns>
        public DrawingValidationResult Validate(IDrawingDoc drawDoc, IView flatPatternView = null)
        {
            const string proc = nameof(DrawingValidator) + ".Validate";
            ErrorHandler.PushCallStack(proc);

            var result = new DrawingValidationResult();

            try
            {
                // Check for dangling dimensions
                var dangling = FindDanglingDimensions(drawDoc);
                result.DanglingDimensions.AddRange(dangling);

                // Check bend line dimension completeness
                if (flatPatternView != null)
                {
                    var undimensioned = FindUndimensionedBendLines(drawDoc, flatPatternView);
                    result.UndimensionedBendLines.AddRange(undimensioned);
                }

                result.Success = result.IssueCount == 0;

                if (result.DanglingDimensions.Count > 0)
                    ErrorHandler.DebugLog($"{proc}: Found {result.DanglingDimensions.Count} dangling dimension(s)");
                if (result.UndimensionedBendLines.Count > 0)
                    ErrorHandler.DebugLog($"{proc}: Found {result.UndimensionedBendLines.Count} undimensioned bend line(s)");
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                result.Warnings.Add("Validation failed: " + ex.Message);
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }

            return result;
        }

        /// <summary>
        /// Finds dangling dimensions by checking annotation color against the system
        /// dangling color. SolidWorks renders dangling dimensions in a configurable
        /// color (default: magenta/red). Also uses IAnnotation.IsDangling as backup.
        /// </summary>
        /// <param name="drawDoc">The drawing document.</param>
        /// <returns>List of dangling dimension descriptions.</returns>
        public List<string> FindDanglingDimensions(IDrawingDoc drawDoc)
        {
            var result = new List<string>();
            if (drawDoc == null) return result;

            // Get the system dangling dimension color for fallback detection
            int danglingColor = -1;
            try
            {
                danglingColor = _swApp.GetUserPreferenceIntegerValue(
                    (int)swUserPreferenceIntegerValue_e.swSystemColorsDanglingDimension);
            }
            catch
            {
                // If we can't read the preference, rely solely on IsDangling
            }

            // GetViews() returns object[] of object[] — each inner array is one sheet
            // Inner array element 0 = sheet-level view, elements 1+ = child views
            var sheets = drawDoc.GetViews() as object[];
            if (sheets == null) return result;

            foreach (object sheetObj in sheets)
            {
                var sheetViews = sheetObj as object[];
                if (sheetViews == null) continue;

                foreach (object viewObj in sheetViews)
                {
                    var view = viewObj as IView;
                    if (view == null) continue;

                    var annotations = view.GetAnnotations() as object[];
                    if (annotations == null) continue;

                    foreach (object annObj in annotations)
                    {
                        var annotation = annObj as IAnnotation;
                        if (annotation == null) continue;

                        if (annotation.GetType() != (int)swAnnotationType_e.swDisplayDimension)
                            continue;

                        bool isDangling = false;

                        // Primary: try IsDangling (SW 2022+ API)
                        try
                        {
                            isDangling = annotation.IsDangling();
                        }
                        catch
                        {
                            // IsDangling not available — fall through to color check
                        }

                        // Fallback: compare annotation color against system dangling color
                        if (!isDangling && danglingColor >= 0)
                        {
                            try
                            {
                                int annColor = annotation.Color;
                                if (annColor == danglingColor)
                                    isDangling = true;
                            }
                            catch
                            {
                                // Color property not accessible
                            }
                        }

                        if (isDangling)
                        {
                            string dimName = "unknown";
                            try
                            {
                                var dispDim = annotation.GetSpecificAnnotation() as IDisplayDimension;
                                if (dispDim != null)
                                {
                                    var dim = dispDim.GetDimension2(0);
                                    if (dim != null)
                                        dimName = dim.GetNameForSelection() ?? "unknown";
                                }
                            }
                            catch
                            {
                                // Best effort name extraction
                            }

                            result.Add($"Dangling dimension '{dimName}' in view '{view.Name}'");
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Finds bend lines in a flat pattern view that do not have an associated dimension.
        /// This catches cases where bends at higher levels may not get auto-dimensioned.
        /// </summary>
        /// <param name="drawDoc">The drawing document.</param>
        /// <param name="flatPatternView">The flat pattern view containing bend lines.</param>
        /// <returns>List of undimensioned bend line descriptions.</returns>
        public List<string> FindUndimensionedBendLines(IDrawingDoc drawDoc, IView flatPatternView)
        {
            var result = new List<string>();
            if (drawDoc == null || flatPatternView == null) return result;

            // Use ViewGeometryAnalyzer to find all bend lines (same as dimensioning code)
            var analyzer = new ViewGeometryAnalyzer(_swApp);
            List<BendElement> horzBends, vertBends;
            analyzer.FindBendLines(drawDoc, flatPatternView, out horzBends, out vertBends);

            var allBends = horzBends.Concat(vertBends).ToList();
            if (allBends.Count == 0) return result;

            // Build a dictionary of bend line segment names → BendElement
            var bendsBySegName = new Dictionary<string, BendElement>();
            foreach (var bend in allBends)
            {
                var seg = bend.SketchSegment as ISketchSegment;
                if (seg == null) continue;

                string segName = null;
                try { segName = seg.GetName(); } catch { }
                if (string.IsNullOrEmpty(segName)) continue;

                if (!bendsBySegName.ContainsKey(segName))
                    bendsBySegName[segName] = bend;
            }

            if (bendsBySegName.Count == 0) return result;

            // Collect all dimension annotation text/names referencing bend lines in this view
            var dimensionedSegNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var annotations = flatPatternView.GetAnnotations() as object[];
            if (annotations != null)
            {
                foreach (object annObj in annotations)
                {
                    var annotation = annObj as IAnnotation;
                    if (annotation == null) continue;

                    if (annotation.GetType() != (int)swAnnotationType_e.swDisplayDimension)
                        continue;

                    try
                    {
                        var dispDim = annotation.GetSpecificAnnotation() as IDisplayDimension;
                        if (dispDim == null) continue;

                        var dim = dispDim.GetDimension2(0);
                        if (dim == null) continue;

                        string dimFullName = dim.GetNameForSelection() ?? "";

                        // Check if this dimension references any bend line segment name
                        foreach (string segName in bendsBySegName.Keys)
                        {
                            if (dimFullName.Contains(segName))
                            {
                                dimensionedSegNames.Add(segName);
                            }
                        }
                    }
                    catch
                    {
                        // Skip annotation if it can't be queried
                    }
                }
            }

            // Report undimensioned bend lines
            foreach (var kvp in bendsBySegName)
            {
                if (!dimensionedSegNames.Contains(kvp.Key))
                {
                    string direction = kvp.Value.Angle == 0 ? "horizontal" : "vertical";
                    double pos = kvp.Value.Position;
                    result.Add($"Undimensioned {direction} bend line '{kvp.Key}' at position {pos:F3}");
                }
            }

            return result;
        }
    }
}
