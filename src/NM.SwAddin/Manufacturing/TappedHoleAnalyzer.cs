using System;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.Manufacturing
{
    public static class TappedHoleAnalyzer
    {
        private const double M_TO_IN = 39.37007874015748;

        public sealed class Result
        {
            public int Setups { get; set; }
            public int Holes { get; set; }
            public bool StainlessNote { get; set; }
        }

        public static Result Analyze(IModelDoc2 model, string materialCode)
        {
            var r = new Result();
            if (model == null) return r;

            try
            {
                IFeature feat = model.FirstFeature() as IFeature;
                while (feat != null)
                {
                    string t = feat.GetTypeName2() ?? string.Empty;
                    if (t.Equals("HoleWzd", StringComparison.OrdinalIgnoreCase))
                    {
                        r.Setups++;
                        try
                        {
                            var defObj = feat.GetDefinition();
                            if (defObj is IWizardHoleFeatureData2 data)
                            {
                                double diaIn = data.HoleDiameter * M_TO_IN;
                                if (diaIn < 1.0) r.Holes++;
                            }
                        }
                        catch { }
                    }
                    feat = feat.GetNextFeature() as IFeature;
                }
            }
            catch { }

            try
            {
                if (!string.IsNullOrWhiteSpace(materialCode))
                {
                    var u = materialCode.ToUpperInvariant();
                    r.StainlessNote = u.Contains("304") || u.Contains("316");
                }
            }
            catch { }

            return r;
        }
    }
}
