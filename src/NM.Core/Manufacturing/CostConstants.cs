using NM.Core.Config;

namespace NM.Core.Manufacturing
{
    /// <summary>
    /// Legacy work center hourly costs — now delegates to NmConfigProvider.
    /// Kept for backward compatibility; new code should use NmConfigProvider.Current.WorkCenters directly.
    /// </summary>
    public static class CostConstants
    {
        // Hourly rates ($/hr) — delegated to config
        public static double F115_COST => NmConfigProvider.Current.WorkCenters.F115_LaserCutting;
        public static double F300_COST => NmConfigProvider.Current.WorkCenters.F300_MaterialHandling;
        public static double F210_COST => NmConfigProvider.Current.WorkCenters.F210_Deburring;
        public static double F140_COST => NmConfigProvider.Current.WorkCenters.F140_PressBrake;
        public static double F145_COST => NmConfigProvider.Current.WorkCenters.F145_CncBending;
        public static double F155_COST => NmConfigProvider.Current.WorkCenters.F155_Waterjet;
        public static double F220_COST => NmConfigProvider.Current.WorkCenters.F220_Tapping;
        public static double F325_COST => NmConfigProvider.Current.WorkCenters.F325_RollForming;
        public static double F400_COST => NmConfigProvider.Current.WorkCenters.F400_Welding;
        public static double F385_COST => NmConfigProvider.Current.WorkCenters.F385_Assembly;
        public static double F500_COST => NmConfigProvider.Current.WorkCenters.F500_Finishing;
        public static double F525_COST => NmConfigProvider.Current.WorkCenters.F525_Packaging;
        public static double ENG_COST => NmConfigProvider.Current.WorkCenters.ENG_Engineering;

        // Pricing modifiers — delegated to config
        public static double MATERIAL_MARKUP => NmConfigProvider.Current.Pricing.MaterialMarkup;
        public static double TIGHT_PERCENT => NmConfigProvider.Current.Pricing.TightTolerance;
        public static double NORMAL_PERCENT => NmConfigProvider.Current.Pricing.NormalTolerance;
        public static double LOOSE_PERCENT => NmConfigProvider.Current.Pricing.LooseTolerance;

        // Order processing — delegated to config
        public static double ORDER_SETUP => NmConfigProvider.Current.Pricing.OrderSetup;
        public static double ORDER_RUN => NmConfigProvider.Current.Pricing.OrderRun;
    }
}
