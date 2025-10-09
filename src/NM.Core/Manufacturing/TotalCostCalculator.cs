using System;

namespace NM.Core.Manufacturing
{
    public sealed class TotalCostInputs
    {
        public int Quantity { get; set; }
        public double RawWeightLb { get; set; }
        public double MaterialCostPerLb { get; set; }
        public double F115Price { get; set; }
        public double F140Price { get; set; }
        public double F220Price { get; set; }
        public double F325Price { get; set; }
        public DifficultyLevel Difficulty { get; set; } = DifficultyLevel.Normal;
    }

    public sealed class TotalCostResult
    {
        public int Quantity { get; set; }
        public double MaterialCostPerLB { get; set; }
        public double OrderCost { get; set; }
        public double MaterialCost { get; set; }
        public double TotalCost { get; set; }
    }

    public static class TotalCostCalculator
    {
        public static TotalCostResult Compute(TotalCostInputs i)
        {
            if (i == null) i = new TotalCostInputs();
            int qty = i.Quantity > 0 ? i.Quantity : 1;

            // Order processing costs
            double orderCost = CostConstants.ORDER_SETUP + CostConstants.ORDER_RUN * Math.Max(0, qty - 1);

            // Material cost with markup
            double matCost = (i.RawWeightLb * qty) * Math.Max(0.0, i.MaterialCostPerLb) * CostConstants.MATERIAL_MARKUP;

            // Sum of work centers
            double wc = Safe(i.F115Price) + Safe(i.F140Price) + Safe(i.F220Price) + Safe(i.F325Price);

            double total = wc + orderCost + matCost;

            // Difficulty modifier
            switch (i.Difficulty)
            {
                case DifficultyLevel.Tight: total *= CostConstants.TIGHT_PERCENT; break;
                case DifficultyLevel.Loose: total *= CostConstants.LOOSE_PERCENT; break;
                default: total *= CostConstants.NORMAL_PERCENT; break;
            }

            // Volume discounts
            if (total > 10000.0)
            {
                total -= qty;
            }
            else if (total > 1000.0)
            {
                total -= ((total - 1000.0) / 9000.0) * qty;
            }

            return new TotalCostResult
            {
                Quantity = qty,
                MaterialCostPerLB = i.MaterialCostPerLb,
                OrderCost = orderCost,
                MaterialCost = matCost,
                TotalCost = total
            };
        }

        private static double Safe(double v) => double.IsNaN(v) || double.IsInfinity(v) ? 0.0 : v;
    }
}
