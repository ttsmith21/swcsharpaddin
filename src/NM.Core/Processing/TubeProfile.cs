using System;

namespace NM.Core.Processing
{
    /// <summary>
    /// Tube/profile shape types.
    /// Ported from VB.NET ExtractData MCommon.EnumShape.
    /// </summary>
    public enum TubeShape
    {
        None = 0,
        Round = 1,
        Square = 2,
        Rectangle = 3,
        Angle = 4,
        Channel = 5
    }

    /// <summary>
    /// Extracted tube/profile geometry data.
    /// Ported from VB.NET ExtractData CStepFile properties.
    /// </summary>
    public sealed class TubeProfile
    {
        /// <summary>
        /// Detected profile shape (round, square, rectangle, angle, channel).
        /// </summary>
        public TubeShape Shape { get; set; } = TubeShape.None;

        /// <summary>
        /// Cross-section description (e.g., "2.5" for round OD, "2 x 3" for rectangular).
        /// </summary>
        public string CrossSection { get; set; } = "";

        /// <summary>
        /// Wall thickness in meters.
        /// </summary>
        public double WallThicknessMeters { get; set; }

        /// <summary>
        /// Material/tube length in meters.
        /// </summary>
        public double MaterialLengthMeters { get; set; }

        /// <summary>
        /// Total cut perimeter length in meters.
        /// </summary>
        public double CutLengthMeters { get; set; }

        /// <summary>
        /// Number of holes detected.
        /// </summary>
        public int NumberOfHoles { get; set; }

        /// <summary>
        /// Outer diameter for round tubes, in meters.
        /// </summary>
        public double OuterDiameterMeters { get; set; }

        /// <summary>
        /// Inner diameter for round tubes, in meters.
        /// </summary>
        public double InnerDiameterMeters { get; set; }

        /// <summary>
        /// Start point of material length measurement.
        /// </summary>
        public double[] StartPoint { get; set; } = new double[3];

        /// <summary>
        /// End point of material length measurement.
        /// </summary>
        public double[] EndPoint { get; set; } = new double[3];

        /// <summary>
        /// Whether extraction was successful.
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// Error or status message.
        /// </summary>
        public string Message { get; set; } = "";

        /// <summary>
        /// Wall thickness in inches.
        /// </summary>
        public double WallThicknessInches => WallThicknessMeters * 39.3701;

        /// <summary>
        /// Material length in inches.
        /// </summary>
        public double MaterialLengthInches => MaterialLengthMeters * 39.3701;

        /// <summary>
        /// Cut length in inches.
        /// </summary>
        public double CutLengthInches => CutLengthMeters * 39.3701;

        /// <summary>
        /// Outer diameter in inches (for round tubes).
        /// </summary>
        public double OuterDiameterInches => OuterDiameterMeters * 39.3701;

        /// <summary>
        /// Inner diameter in inches (for round tubes).
        /// </summary>
        public double InnerDiameterInches => InnerDiameterMeters * 39.3701;

        /// <summary>
        /// Shape name as string.
        /// </summary>
        public string ShapeName
        {
            get
            {
                switch (Shape)
                {
                    case TubeShape.Round: return "Round";
                    case TubeShape.Square: return "Square";
                    case TubeShape.Rectangle: return "Rectangle";
                    case TubeShape.Angle: return "Angle";
                    case TubeShape.Channel: return "Channel";
                    default: return "";
                }
            }
        }

        public override string ToString()
        {
            if (!Success)
                return $"TubeProfile: Failed - {Message}";

            return $"TubeProfile: {ShapeName}, CrossSection={CrossSection}, " +
                   $"Wall={WallThicknessInches:F3}\", Length={MaterialLengthInches:F3}\", " +
                   $"CutLength={CutLengthInches:F3}\", Holes={NumberOfHoles}";
        }
    }
}
