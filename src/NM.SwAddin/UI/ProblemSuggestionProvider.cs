using System.Collections.Generic;
using NM.Core.ProblemParts;

namespace NM.SwAddin.UI
{
    /// <summary>
    /// Provides problem-category-specific suggestions for fixing problem parts.
    /// Single source of truth — replaces duplicated logic in ProblemWizardForm and ProblemReviewDialog.
    /// </summary>
    public static class ProblemSuggestionProvider
    {
        /// <summary>
        /// Returns actionable suggestions based on the problem category and optional error text.
        /// </summary>
        public static List<string> GetSuggestions(ProblemPartManager.ProblemCategory category, string errorText = null)
        {
            var list = new List<string>();
            var errLower = (errorText ?? "").ToLowerInvariant();

            switch (category)
            {
                case ProblemPartManager.ProblemCategory.GeometryValidation:
                    if (errLower.Contains("surface") || errLower.Contains("no solid"))
                    {
                        list.Add("Part may be surface-only - needs solid geometry");
                        list.Add("Use Insert > Surface > Thicken to create solid");
                        list.Add("Check if part file is corrupted or empty");
                    }
                    else
                    {
                        list.Add("Review part geometry in SolidWorks");
                        list.Add("Check for invalid features or geometry errors");
                    }
                    break;

                case ProblemPartManager.ProblemCategory.MaterialMissing:
                    list.Add("Right-click part in FeatureManager > Material > Edit Material");
                    list.Add("Or: Click the material icon in FeatureManager tree");
                    list.Add("Assign appropriate material (e.g., AISI 304, Plain Carbon Steel)");
                    break;

                case ProblemPartManager.ProblemCategory.SheetMetalConversion:
                    list.Add("Check that part has uniform wall thickness");
                    list.Add("Verify parallel faces exist for sheet metal conversion");
                    list.Add("Ensure no complex features block conversion");
                    list.Add("Try using Insert > Sheet Metal > Insert Bends manually");
                    break;

                case ProblemPartManager.ProblemCategory.FileAccess:
                    list.Add("Verify file exists at the shown path");
                    list.Add("Check file is not read-only or locked");
                    list.Add("Ensure file is not open in another application");
                    break;

                case ProblemPartManager.ProblemCategory.Suppressed:
                    list.Add("Component is suppressed in assembly");
                    list.Add("Right-click component > Set to Resolved");
                    list.Add("Check if component was intentionally excluded");
                    break;

                case ProblemPartManager.ProblemCategory.Lightweight:
                    list.Add("Component is in Lightweight mode");
                    list.Add("Right-click > Set to Resolved");
                    list.Add("Or: Tools > Options > Assemblies > uncheck Lightweight");
                    break;

                case ProblemPartManager.ProblemCategory.MixedBody:
                    list.Add("Part has both solid and surface bodies");
                    list.Add("Delete surface bodies: FeatureManager > right-click Surface Body > Delete");
                    list.Add("If this is a purchased/catalog part, click PUR button below");
                    list.Add("Surface bodies can interfere with geometry analysis");
                    break;

                case ProblemPartManager.ProblemCategory.MultiBody:
                    list.Add("This part has multiple solid bodies");
                    list.Add("If this is an oversized part you split: click 'Split → Assy' to save each body as a separate part");
                    list.Add("If unintentional: use Insert > Features > Combine to merge bodies, then Retry");
                    list.Add("Or delete unwanted bodies in FeatureManager, then Retry");
                    break;

                case ProblemPartManager.ProblemCategory.ThicknessExtraction:
                    list.Add("Could not determine sheet metal thickness");
                    list.Add("Check that part has uniform wall thickness");
                    list.Add("May need manual thickness specification");
                    break;

                case ProblemPartManager.ProblemCategory.ManufacturingWarning:
                    list.Add("Hole or cutout is too close to a bend line");
                    list.Add("Feature will distort during bending");
                    list.Add("Option 1: Relocate feature further from bend line");
                    list.Add("Option 2: Add relief slot between feature and bend");
                    break;

                case ProblemPartManager.ProblemCategory.NestingEfficiencyOverride:
                    list.Add("Part fills less than 80% of its bounding rectangle (high scrap ratio)");
                    list.Add("Weight calculation switched from 80% nesting efficiency to bounding-box L\u00d7W mode");
                    list.Add("Click 'Revert to 80%' if this part nests well with other parts on the same sheet");
                    list.Add("Otherwise, accept the override for a more accurate material estimate");
                    break;

                default:
                    list.Add("Review part in SolidWorks");
                    list.Add("Check error message for specific guidance");
                    break;
            }

            return list;
        }
    }
}
