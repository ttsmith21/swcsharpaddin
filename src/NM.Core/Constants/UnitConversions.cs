namespace NM.Core.Constants
{
    /// <summary>
    /// Physical unit conversion constants. These are physics — they never change.
    /// Single source of truth; all code should reference these instead of local copies.
    /// </summary>
    public static class UnitConversions
    {
        public const double MetersToInches = 39.37007874015748;
        public const double InchesToMeters = 1.0 / 39.37007874015748;
        public const double KgToLbs = 2.2046226218487757;
        public const double LbsToKg = 1.0 / 2.2046226218487757;
        public const double MetersToFeet = 3.28084;
        public const double InchesToFeet = 1.0 / 12.0;
        public const double FeetToInches = 12.0;
    }
}
