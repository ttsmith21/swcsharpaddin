using System;
using System.Collections.Generic;
using NM.Core;
using NM.Core.Drawing;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.Drawing
{
    /// <summary>
    /// Auto-dimensions drawing views: flat pattern bend-to-bend dimensions,
    /// overall dimensions, tube profile dimensions, and formed view dimensions.
    /// Ported from VBA DimensionDrawing.bas: DimensionFlat, DimensionTube,
    /// DimensionOther, RectProfile, and AlignDims.bas: Align.
    /// </summary>
    public sealed class DrawingDimensioner
    {
        private readonly ISldWorks _swApp;
        private readonly ViewGeometryAnalyzer _analyzer;

        /// <summary>
        /// Conversion factor: inches per meter (used in VBA as 39.3700787401575).
        /// </summary>
        private const double InchesPerMeter = 39.3700787401575;

        public DrawingDimensioner(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
            _analyzer = new ViewGeometryAnalyzer(swApp);
        }

        /// <summary>
        /// Result of a dimensioning operation.
        /// </summary>
        public sealed class DimensionResult
        {
            public bool Success { get; set; }
            public int DimensionsAdded { get; set; }
            public int ViewsCreated { get; set; }
            public string Message { get; set; }
        }

        /// <summary>
        /// Dimensions a flat pattern view: overall width/height, bend-to-bend dimensions
        /// with " BL" suffix, and creates projected views for the formed state.
        /// Ported from VBA DimensionFlat().
        /// </summary>
        /// <param name="drawDoc">The drawing document.</param>
        /// <param name="view">The flat pattern view to dimension.</param>
        /// <returns>Result with count of dimensions added and views created.</returns>
        public DimensionResult DimensionFlatPattern(IDrawingDoc drawDoc, IView view)
        {
            // TODO: Phase 3 implementation
            // 1. FindBendLines → horizontal and vertical bends
            // 2. FindBoundaryEdge for all 4 directions
            // 3. If no boundaries found → fall back to DimensionTube
            // 4. Add overall horizontal dimension (left edge + right edge)
            // 5. Add overall vertical dimension (top edge + bottom edge)
            // 6. For vertical bends: add bend-to-bend horizontal dims with " BL" suffix
            // 7. Create top projected view → DimensionFormedView
            // 8. For horizontal bends: add bend-to-bend vertical dims with " BL" suffix
            // 9. Create right projected view → DimensionFormedView
            // 10. Auto-scale views to fit on sheet
            // 11. AlignAllDimensions
            return new DimensionResult { Message = "Not yet implemented" };
        }

        /// <summary>
        /// Dimensions a tube view: detects circular or rectangular profile,
        /// adds diameter or width/height dimensions.
        /// Ported from VBA DimensionTube().
        /// </summary>
        /// <param name="drawDoc">The drawing document.</param>
        /// <param name="view">The tube view to dimension.</param>
        /// <returns>Result with count of dimensions added.</returns>
        public DimensionResult DimensionTube(IDrawingDoc drawDoc, IView view)
        {
            // TODO: Phase 3 implementation
            // 1. Get view outline
            // 2. Calculate edge positions from outline
            // 3. Try SelectByRay on left edge → check if arc/circle
            // 4. If circular: AddDiameterDimension2
            // 5. If rectangular: DimensionRectProfile (select opposing edges)
            // 6. Fallback: try bottom edge for horizontal dim
            return new DimensionResult { Message = "Not yet implemented" };
        }

        /// <summary>
        /// Dimensions a formed (3D) projected view with overall width and height.
        /// Ported from VBA DimensionOther().
        /// </summary>
        /// <param name="drawDoc">The drawing document.</param>
        /// <param name="view">The formed/projected view.</param>
        /// <returns>Result with count of dimensions added.</returns>
        public DimensionResult DimensionFormedView(IDrawingDoc drawDoc, IView view)
        {
            // TODO: Phase 3 implementation
            // 1. FindBoundaryEdge for all 4 directions
            // 2. Select left + right edges → AddHorizontalDimension2
            // 3. Select top + bottom edges → AddVerticalDimension2
            return new DimensionResult { Message = "Not yet implemented" };
        }

        /// <summary>
        /// Attempts to dimension a rectangular tube profile by selecting opposing edges.
        /// Ported from VBA RectProfile().
        /// </summary>
        /// <returns>True if dimensions were successfully added.</returns>
        public bool DimensionRectProfile(IDrawingDoc drawDoc, IView view,
            EdgeElement left, EdgeElement right, EdgeElement top, EdgeElement bottom)
        {
            // TODO: Phase 3 implementation
            // 1. SelectByRay on left edge + right edge → AddHorizontalDimension2
            // 2. SelectByRay on top edge + bottom edge → AddVerticalDimension2
            // 3. Return false if any selection or dimension fails
            return false;
        }

        /// <summary>
        /// Auto-arranges all dimensions in the drawing using SolidWorks' built-in alignment.
        /// Ported from VBA AlignDims.bas::Align().
        /// </summary>
        /// <param name="drawDoc">The drawing document.</param>
        public void AlignAllDimensions(IDrawingDoc drawDoc)
        {
            // TODO: Phase 3 implementation
            // 1. Get current sheet
            // 2. Get all views on the sheet
            // 3. For each view: get annotations, select all
            // 4. Call Extension.AlignDimensions(swAlignDimensionType_AutoArrange, 0.001)
            // 5. Clear selection, redraw
        }
    }
}
