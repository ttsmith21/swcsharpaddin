using System;
using System.IO;
using System.Linq;
using NM.Core.DataModel;
using NM.Core.Pdf.Models;
using NM.Core.Reconciliation.Models;

namespace NM.Core.Reconciliation
{
    /// <summary>
    /// Merges 3D model data (PartData) with PDF drawing data (DrawingData).
    /// Detects conflicts, fills gaps, generates routing suggestions, and suggests renames.
    /// </summary>
    public sealed class ReconciliationEngine
    {
        private readonly RoutingNoteInterpreter _routingInterpreter;

        public ReconciliationEngine()
        {
            _routingInterpreter = new RoutingNoteInterpreter();
        }

        /// <summary>
        /// Reconciles 3D model data with 2D drawing data.
        /// Either parameter may be null (produces a partial result).
        /// </summary>
        public ReconciliationResult Reconcile(PartData model, DrawingData drawing)
        {
            var result = new ReconciliationResult();

            if (drawing == null) return result;

            // 1. Cross-validate overlapping fields
            if (model != null)
            {
                CompareMaterials(model, drawing, result);
                CompareThickness(model, drawing, result);
                CompareDescription(model, drawing, result);
                ComparePartNumber(model, drawing, result);
            }

            // 2. Fill gaps — drawing data fills empty model fields
            if (model != null)
            {
                TryFillGap(result, "PartNumber", GetPartNumber(model), drawing.PartNumber, "PDF title block");
                TryFillGap(result, "Description", GetDescription(model), drawing.Description, "PDF title block");
                TryFillGap(result, "Revision", GetRevision(model), drawing.Revision, "PDF title block");
                TryFillGap(result, "Material", model.Material, drawing.Material, "PDF title block");
                TryFillGap(result, "Finish", null, drawing.Finish, "PDF title block");
            }
            else
            {
                // No model — all drawing data becomes gap fills
                if (!string.IsNullOrEmpty(drawing.PartNumber))
                    result.GapFills.Add(new GapFill { Field = "PartNumber", Value = drawing.PartNumber, Source = "PDF title block", Confidence = 0.85 });
                if (!string.IsNullOrEmpty(drawing.Description))
                    result.GapFills.Add(new GapFill { Field = "Description", Value = drawing.Description, Source = "PDF title block", Confidence = 0.80 });
                if (!string.IsNullOrEmpty(drawing.Revision))
                    result.GapFills.Add(new GapFill { Field = "Revision", Value = drawing.Revision, Source = "PDF title block", Confidence = 0.90 });
                if (!string.IsNullOrEmpty(drawing.Material))
                    result.GapFills.Add(new GapFill { Field = "Material", Value = drawing.Material, Source = "PDF title block", Confidence = 0.85 });
            }

            // 3. Routing suggestions from drawing notes
            var routingSuggestions = _routingInterpreter.InterpretRoutingHints(drawing.RoutingHints);
            result.RoutingSuggestions.AddRange(routingSuggestions);

            // 4. File rename suggestion
            if (model != null && !string.IsNullOrEmpty(drawing.PartNumber))
            {
                SuggestRename(model, drawing, result);
            }

            return result;
        }

        private void CompareMaterials(PartData model, DrawingData drawing, ReconciliationResult result)
        {
            if (string.IsNullOrEmpty(model.Material) || string.IsNullOrEmpty(drawing.Material))
                return;

            string modelMat = NormalizeMaterial(model.Material);
            string drawingMat = NormalizeMaterial(drawing.Material);

            if (string.Equals(modelMat, drawingMat, StringComparison.OrdinalIgnoreCase))
            {
                result.Confirmations.Add($"Material matches: {model.Material}");
                return;
            }

            // Check if they're equivalent (e.g., "304" vs "304 STAINLESS STEEL")
            if (MaterialsEquivalent(modelMat, drawingMat))
            {
                result.Confirmations.Add($"Material equivalent: {model.Material} ≈ {drawing.Material}");
                return;
            }

            result.Conflicts.Add(new DataConflict
            {
                Field = "Material",
                ModelValue = model.Material,
                DrawingValue = drawing.Material,
                Severity = ConflictSeverity.High,
                Recommendation = ConflictResolution.HumanRequired,
                Reason = "3D model material does not match drawing material"
            });
        }

        private void CompareThickness(PartData model, DrawingData drawing, ReconciliationResult result)
        {
            if (model.Thickness_m <= 0 || !drawing.Thickness_in.HasValue)
                return;

            double modelThickness_in = model.Thickness_m / 0.0254; // meters → inches
            double drawingThickness_in = drawing.Thickness_in.Value;

            double tolerance = 0.005; // 0.005" tolerance for comparison

            if (Math.Abs(modelThickness_in - drawingThickness_in) <= tolerance)
            {
                result.Confirmations.Add($"Thickness matches: {modelThickness_in:F3}\"");
                return;
            }

            result.Conflicts.Add(new DataConflict
            {
                Field = "Thickness",
                ModelValue = $"{modelThickness_in:F4}\"",
                DrawingValue = $"{drawingThickness_in:F4}\"",
                Severity = ConflictSeverity.High,
                Recommendation = ConflictResolution.UseModel,
                Reason = "3D model geometry is measured; drawing value may be nominal. Recommend using model value."
            });
        }

        private void CompareDescription(PartData model, DrawingData drawing, ReconciliationResult result)
        {
            string modelDesc = GetDescription(model);
            if (string.IsNullOrEmpty(modelDesc) || string.IsNullOrEmpty(drawing.Description))
                return;

            if (string.Equals(modelDesc.Trim(), drawing.Description.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                result.Confirmations.Add($"Description matches: {modelDesc}");
                return;
            }

            // Descriptions differ — low severity since drawings often have more detail
            result.Conflicts.Add(new DataConflict
            {
                Field = "Description",
                ModelValue = modelDesc,
                DrawingValue = drawing.Description,
                Severity = ConflictSeverity.Low,
                Recommendation = ConflictResolution.UseDrawing,
                Reason = "Drawing description is typically more complete than model property"
            });
        }

        private void ComparePartNumber(PartData model, DrawingData drawing, ReconciliationResult result)
        {
            string modelPn = GetPartNumber(model);
            if (string.IsNullOrEmpty(modelPn) || string.IsNullOrEmpty(drawing.PartNumber))
                return;

            if (string.Equals(modelPn.Trim(), drawing.PartNumber.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                result.Confirmations.Add($"Part number matches: {modelPn}");
                return;
            }

            result.Conflicts.Add(new DataConflict
            {
                Field = "PartNumber",
                ModelValue = modelPn,
                DrawingValue = drawing.PartNumber,
                Severity = ConflictSeverity.Medium,
                Recommendation = ConflictResolution.UseDrawing,
                Reason = "Drawing part number is the official reference; model property may be outdated"
            });
        }

        private void TryFillGap(ReconciliationResult result, string field, string modelValue, string drawingValue, string source)
        {
            if (!string.IsNullOrEmpty(modelValue) || string.IsNullOrEmpty(drawingValue))
                return;

            double confidence = 0.85;
            if (field == "Revision") confidence = 0.90;
            if (field == "Description") confidence = 0.80;
            if (field == "Finish") confidence = 0.75;

            result.GapFills.Add(new GapFill
            {
                Field = field,
                Value = drawingValue,
                Source = source,
                Confidence = confidence
            });
        }

        private void SuggestRename(PartData model, DrawingData drawing, ReconciliationResult result)
        {
            if (string.IsNullOrEmpty(model.FilePath) || string.IsNullOrEmpty(drawing.PartNumber))
                return;

            string currentFilename = Path.GetFileNameWithoutExtension(model.FilePath);
            string drawingPN = SanitizeFilename(drawing.PartNumber);

            // If the filename already matches the part number, no rename needed
            if (string.Equals(currentFilename, drawingPN, StringComparison.OrdinalIgnoreCase))
                return;

            // Build new filename
            string ext = Path.GetExtension(model.FilePath);
            string dir = Path.GetDirectoryName(model.FilePath);
            string newName = drawingPN;

            // Add description if available
            if (!string.IsNullOrEmpty(drawing.Description))
            {
                string safeDesc = SanitizeFilename(drawing.Description);
                if (safeDesc.Length <= 30) // Keep filename reasonable length
                    newName = drawingPN + "_" + safeDesc;
            }

            string newPath = Path.Combine(dir ?? "", newName + ext);

            result.Rename = new RenameSuggestion
            {
                OldPath = model.FilePath,
                NewPath = newPath,
                Reason = $"Drawing part number '{drawing.PartNumber}' differs from filename '{currentFilename}'",
                Confidence = 0.85
            };

            // Check for companion drawing file
            string drawingExt = ".slddrw";
            string oldDrawingPath = Path.Combine(dir ?? "", currentFilename + drawingExt);
            if (File.Exists(oldDrawingPath))
            {
                result.Rename.OldDrawingPath = oldDrawingPath;
                result.Rename.NewDrawingPath = Path.Combine(dir ?? "", newName + drawingExt);
            }
        }

        // --- Helper methods ---

        private static string GetPartNumber(PartData model)
        {
            if (model.Extra.TryGetValue("PartNumber", out string pn) && !string.IsNullOrWhiteSpace(pn))
                return pn;
            return null;
        }

        private static string GetDescription(PartData model)
        {
            if (model.Extra.TryGetValue("Description", out string desc) && !string.IsNullOrWhiteSpace(desc))
                return desc;
            return null;
        }

        private static string GetRevision(PartData model)
        {
            if (model.Extra.TryGetValue("Revision", out string rev) && !string.IsNullOrWhiteSpace(rev))
                return rev;
            return null;
        }

        private static string NormalizeMaterial(string material)
        {
            if (string.IsNullOrEmpty(material)) return "";
            return material.Trim()
                .Replace("STAINLESS STEEL", "SS")
                .Replace("CARBON STEEL", "CS")
                .Replace("ALUMINUM", "AL")
                .Replace("  ", " ")
                .ToUpperInvariant();
        }

        private static bool MaterialsEquivalent(string a, string b)
        {
            // Extract the core alloy number (e.g., "304" from "304 SS" or "304 STAINLESS STEEL")
            string coreA = ExtractAlloyCoreNumber(a);
            string coreB = ExtractAlloyCoreNumber(b);

            if (!string.IsNullOrEmpty(coreA) && !string.IsNullOrEmpty(coreB))
                return string.Equals(coreA, coreB, StringComparison.OrdinalIgnoreCase);

            // Check if one contains the other
            return a.Contains(b) || b.Contains(a);
        }

        private static string ExtractAlloyCoreNumber(string material)
        {
            if (string.IsNullOrEmpty(material)) return null;

            // Extract leading digits (e.g., "304", "1018", "6061", "A36")
            var match = System.Text.RegularExpressions.Regex.Match(material, @"\b([A]?\d{3,5}[L]?)\b");
            return match.Success ? match.Groups[1].Value : null;
        }

        private static string SanitizeFilename(string name)
        {
            if (string.IsNullOrEmpty(name)) return name;

            char[] invalid = Path.GetInvalidFileNameChars();
            string result = name;
            foreach (char c in invalid)
                result = result.Replace(c, '_');

            // Replace spaces with underscores for cleaner filenames
            result = result.Replace(' ', '_').Replace("__", "_").Trim('_');
            return result.ToUpperInvariant();
        }
    }
}
