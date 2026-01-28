using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using NM.Core;
using NM.Core.Manufacturing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Manufacturing
{
    /// <summary>
    /// Analyzes parts for tapped holes and calculates F220 (tapping) costs.
    /// Ported from VBA modMaterialCost.bas TappedHoles() function.
    /// </summary>
    public static class TappedHoleAnalyzer
    {
        private const double M_TO_IN = 39.37007874015748;

        // Regex patterns for detecting tapped holes by feature name
        private static readonly Regex TapPattern = new Regex(
            @"(?:TAP|TAPPED|THREAD|1/4-20|1/4-28|5/16-18|5/16-24|3/8-16|3/8-24|1/2-13|1/2-20|#\d+-\d+|M\d+x[\d.]+)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public sealed class TappedHoleResult
        {
            public int TotalSetups { get; set; }
            public int TotalHoles { get; set; }
            public int TappedHoleCount { get; set; }
            public int ClearanceHoleCount { get; set; }
            public bool RequiresStainlessNote { get; set; }
            public string StainlessNoteText { get; set; }
            public List<TappedHoleInfo> HoleDetails { get; } = new List<TappedHoleInfo>();

            // F220 Calculations
            public double F220_Setup_Hours { get; set; }
            public double F220_Run_Hours { get; set; }
            public double F220_Price { get; set; }
        }

        public sealed class TappedHoleInfo
        {
            public string FeatureName { get; set; }
            public string Size { get; set; }
            public double DiameterIn { get; set; }
            public int HoleCount { get; set; }
            public bool IsTapped { get; set; }
        }

        /// <summary>
        /// Analyzes model for tapped holes and calculates F220 costs.
        /// </summary>
        public static TappedHoleResult Analyze(IModelDoc2 model, string materialCode, int quantity = 1)
        {
            const string proc = nameof(Analyze);
            ErrorHandler.PushCallStack(proc);

            var result = new TappedHoleResult();
            if (model == null)
            {
                ErrorHandler.PopCallStack();
                return result;
            }

            try
            {
                var uniqueSizes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                IFeature feat = model.FirstFeature() as IFeature;
                while (feat != null)
                {
                    string typeName = feat.GetTypeName2() ?? string.Empty;
                    string featName = feat.Name ?? string.Empty;

                    // Check Hole Wizard features
                    if (typeName.Equals("HoleWzd", StringComparison.OrdinalIgnoreCase))
                    {
                        ProcessHoleWizardFeature(feat, featName, result, uniqueSizes);
                    }
                    // Check for simple holes with TAP in name
                    else if (typeName.Equals("HoleSimple", StringComparison.OrdinalIgnoreCase) ||
                             typeName.Equals("Cut", StringComparison.OrdinalIgnoreCase))
                    {
                        if (TapPattern.IsMatch(featName))
                        {
                            result.TappedHoleCount++;
                            result.TotalHoles++;
                            if (!uniqueSizes.Contains(featName))
                            {
                                uniqueSizes.Add(featName);
                                result.TotalSetups++;
                            }
                            result.HoleDetails.Add(new TappedHoleInfo
                            {
                                FeatureName = featName,
                                Size = ExtractSizeFromName(featName),
                                IsTapped = true,
                                HoleCount = 1
                            });
                        }
                    }

                    feat = feat.GetNextFeature() as IFeature;
                }

                // Check for stainless steel note requirement
                CheckStainlessNote(result, materialCode);

                // Calculate F220 costs
                CalculateF220Costs(result, quantity);
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }

            return result;
        }

        private static void ProcessHoleWizardFeature(IFeature feat, string featName, TappedHoleResult result, HashSet<string> uniqueSizes)
        {
            try
            {
                var defObj = feat.GetDefinition();
                if (defObj is IWizardHoleFeatureData2 data)
                {
                    double diaIn = data.HoleDiameter * M_TO_IN;
                    int holeType = data.Type;
                    int holeCount = 1;

                    // Try to get hole count from feature
                    try
                    {
                        // Hole wizard can have multiple instances
                        var selData = data as object;
                        // Note: Getting exact count requires accessing sketch points
                    }
                    catch { }

                    // Determine if this is a tapped hole
                    bool isTapped = IsTappedHoleType(holeType) || TapPattern.IsMatch(featName);
                    string size = ExtractSizeFromData(data) ?? ExtractSizeFromName(featName);

                    var info = new TappedHoleInfo
                    {
                        FeatureName = featName,
                        Size = size,
                        DiameterIn = diaIn,
                        HoleCount = holeCount,
                        IsTapped = isTapped
                    };
                    result.HoleDetails.Add(info);

                    if (isTapped)
                    {
                        result.TappedHoleCount += holeCount;
                        result.TotalHoles += holeCount;

                        // Count unique setups by size
                        if (!string.IsNullOrEmpty(size) && !uniqueSizes.Contains(size))
                        {
                            uniqueSizes.Add(size);
                            result.TotalSetups++;
                        }
                        else if (string.IsNullOrEmpty(size))
                        {
                            result.TotalSetups++;
                        }
                    }
                    else
                    {
                        result.ClearanceHoleCount += holeCount;
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Checks if the hole wizard type indicates a tapped hole.
        /// </summary>
        private static bool IsTappedHoleType(int holeType)
        {
            // swWzdHoleType_e values:
            // 0 = Hole, 1 = Countersink, 2 = Counterbore, 3 = Tapped, 4 = Pipe Tap, 5 = Legacy
            return holeType == 3 || holeType == 4;
        }

        private static string ExtractSizeFromData(IWizardHoleFeatureData2 data)
        {
            try
            {
                // Get hole diameter and try to match to standard sizes
                double diaIn = data.HoleDiameter * M_TO_IN;
                // Return diameter as string for size identification
                return $"{diaIn:0.###}";
            }
            catch { }
            return null;
        }

        private static string ExtractSizeFromName(string featName)
        {
            if (string.IsNullOrEmpty(featName))
                return null;

            // Try to extract thread size from feature name
            var match = Regex.Match(featName, @"(\d+/\d+-\d+|#\d+-\d+|M\d+x[\d.]+)", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;

            return featName;
        }

        private static void CheckStainlessNote(TappedHoleResult result, string materialCode)
        {
            if (result.TappedHoleCount == 0)
                return;

            if (string.IsNullOrWhiteSpace(materialCode))
                return;

            var m = materialCode.ToUpperInvariant();
            if (m.Contains("304") || m.Contains("316") || m.Contains("309") || m.Contains("SS"))
            {
                result.RequiresStainlessNote = true;
                result.StainlessNoteText = "Confirm drill size for SS tapped holes";
            }
        }

        private static void CalculateF220Costs(TappedHoleResult result, int quantity)
        {
            if (result.TappedHoleCount == 0)
                return;

            // Use the F220Calculator for consistent costing
            var input = new F220Input
            {
                Setups = result.TotalSetups,
                Holes = result.TappedHoleCount
            };

            var f220Result = F220Calculator.Compute(input);

            result.F220_Setup_Hours = f220Result.SetupHours;
            result.F220_Run_Hours = f220Result.RunHours * Math.Max(1, quantity);
            result.F220_Price = (result.F220_Setup_Hours + result.F220_Run_Hours) * CostConstants.F220_COST;
        }
    }
}
