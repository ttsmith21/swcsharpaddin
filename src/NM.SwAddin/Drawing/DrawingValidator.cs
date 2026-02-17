using System;
using System.Collections.Generic;
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
            // TODO: Phase 5 implementation
            // 1. Get the system dangling dimension color from user preferences
            // 2. Iterate all sheets → all views → all annotations
            // 3. For each IDisplayDimension:
            //    a. Get IAnnotation
            //    b. Check annotation.Color against dangling color
            //    c. Also check annotation.IsDangling as backup
            //    d. If dangling, add dimension name/value to result list
            // 4. Also check sheet-level annotations (not in any view)
            return new List<string>();
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
            // TODO: Phase 5 implementation
            // 1. Get all bend lines from flatPatternView.GetBendLines()
            // 2. Build a set of bend line sketch segment names
            // 3. Iterate all dimensions in the view
            // 4. For each dimension, check if it references a bend line segment
            //    (match by constructed name: "{segName}@{sketchName}@{componentName}@{viewName}")
            // 5. Return bend lines not referenced by any dimension
            return new List<string>();
        }
    }
}
