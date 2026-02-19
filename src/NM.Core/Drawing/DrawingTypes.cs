using System.Collections.Generic;

namespace NM.Core.Drawing
{
    /// <summary>
    /// A line/edge processed into view-space coordinates with computed angle.
    /// Ported from VBA DimensionDrawing.bas ProcessedElement type.
    /// </summary>
    public sealed class ProcessedElement
    {
        /// <summary>The original SolidWorks object (IEdge, ISketchSegment, etc.).</summary>
        public object Obj { get; set; }

        public double X1 { get; set; }
        public double X2 { get; set; }
        public double Y1 { get; set; }
        public double Y2 { get; set; }

        /// <summary>Angle in degrees: 0 = horizontal, 90 = vertical.</summary>
        public double Angle { get; set; }
    }

    /// <summary>
    /// A bend line in a flat pattern drawing view, with its sorted position.
    /// Ported from VBA DimensionDrawing.bas BendElement type.
    /// </summary>
    public sealed class BendElement
    {
        /// <summary>The ISketchSegment representing this bend line.</summary>
        public object SketchSegment { get; set; }

        /// <summary>Primary position (X for vertical bends, Y for horizontal bends) in view-space inches.</summary>
        public double Position { get; set; }

        /// <summary>Secondary position used for sorting co-located bends.</summary>
        public double P2 { get; set; }

        /// <summary>Angle in degrees: 0 = horizontal bend, 90 = vertical bend.</summary>
        public double Angle { get; set; }
    }

    /// <summary>
    /// A boundary edge or vertex at the extreme of a drawing view.
    /// Ported from VBA DimensionDrawing.bas EdgeElement type.
    /// </summary>
    public sealed class EdgeElement
    {
        /// <summary>The SolidWorks object (IEdge, IVertex, ISketchSegment, etc.).</summary>
        public object Obj { get; set; }

        /// <summary>X position in view-space inches.</summary>
        public double X { get; set; }

        /// <summary>Y position in view-space inches.</summary>
        public double Y { get; set; }

        /// <summary>Angle of the edge in degrees.</summary>
        public double Angle { get; set; }

        /// <summary>"Line" or "Point" — indicates whether a full edge or vertex was found.</summary>
        public string Type { get; set; }
    }

    /// <summary>
    /// Direction for boundary edge finding.
    /// </summary>
    public enum BoundaryDirection
    {
        Left,
        Right,
        Top,
        Bottom
    }

    /// <summary>
    /// Result of drawing validation checks (dangling dimensions, undimensioned bend lines).
    /// </summary>
    public sealed class DrawingValidationResult
    {
        public bool Success { get; set; }

        /// <summary>Dimensions that have lost their reference geometry (dangling).</summary>
        public List<string> DanglingDimensions { get; set; } = new List<string>();

        /// <summary>Bend lines that do not have an associated dimension.</summary>
        public List<string> UndimensionedBendLines { get; set; } = new List<string>();

        /// <summary>Non-fatal warnings encountered during validation.</summary>
        public List<string> Warnings { get; set; } = new List<string>();

        /// <summary>Total number of issues found.</summary>
        public int IssueCount => DanglingDimensions.Count + UndimensionedBendLines.Count;
    }

    /// <summary>
    /// Result of the drawing generation + dimensioning process.
    /// </summary>
    public sealed class DrawingCreationResult
    {
        public bool Success { get; set; }
        public string DrawingPath { get; set; }
        public string DxfPath { get; set; }
        public string Message { get; set; }
        public bool WasExisting { get; set; }

        /// <summary>Number of dimensions added to the drawing.</summary>
        public int DimensionsAdded { get; set; }

        /// <summary>Number of views created in the drawing.</summary>
        public int ViewsCreated { get; set; }

        /// <summary>Validation result, if validation was run after creation.</summary>
        public DrawingValidationResult Validation { get; set; }

        /// <summary>DimXpert auto-dimensioning results, if DimXpert was run.</summary>
        public DimXpertSummary DimXpert { get; set; }

        /// <summary>Hole table insertion results, if hole table was inserted.</summary>
        public HoleTableSummary HoleTable { get; set; }
    }

    /// <summary>
    /// Summary of DimXpert auto-dimensioning results for a part.
    /// </summary>
    public sealed class DimXpertSummary
    {
        public bool WasRun { get; set; }
        public bool Success { get; set; }
        public int FeaturesRecognized { get; set; }
        public int AnnotationsImported { get; set; }
        public List<string> FeatureTypes { get; set; } = new List<string>();
        public string FailureReason { get; set; }
    }

    /// <summary>
    /// Summary of hole table insertion results.
    /// </summary>
    public sealed class HoleTableSummary
    {
        public bool WasInserted { get; set; }
        public int HolesFound { get; set; }
        public string FailureReason { get; set; }
    }
}
