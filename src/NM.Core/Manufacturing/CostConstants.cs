namespace NM.Core.Manufacturing
{
    // Legacy work center hourly costs and pricing modifiers (exact values from VBA)
    public static class CostConstants
    {
        // Hourly rates ($/hr)
        public const double F115_COST = 120.0;  // Laser cutting
        public const double F300_COST = 44.0;   // Material handling
        public const double F210_COST = 42.0;   // Deburring
        public const double F140_COST = 80.0;   // Press brake
        public const double F145_COST = 175.0;  // CNC bending
        public const double F155_COST = 120.0;  // Waterjet (deprecated)
        public const double F220_COST = 65.0;   // Tapping
        public const double F325_COST = 65.0;   // Roll forming
        public const double F400_COST = 48.0;   // Welding
        public const double F385_COST = 37.0;   // Assembly
        public const double F500_COST = 48.0;   // Finishing
        public const double F525_COST = 47.0;   // Packaging
        public const double ENG_COST = 50.0;    // Engineering

        // Pricing modifiers
        public const double MATERIAL_MARKUP = 1.05; // 5%
        public const double TIGHT_PERCENT = 1.15;   // +15%
        public const double NORMAL_PERCENT = 1.0;   // base
        public const double LOOSE_PERCENT = 0.95;   // -5%

        // Order processing (fixed $)
        public const double ORDER_SETUP = 20.0;
        public const double ORDER_RUN = 3.0;
    }
}
