using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NM.Core.DataModel;

namespace NM.Core.Export
{
    /// <summary>
    /// Builds ErpExportData from a collection of PartData DTOs.
    /// Maps PartData fields to the ERP export format expected by ErpExportFormat.
    /// </summary>
    public static class ErpExportDataBuilder
    {
        private const double KG_TO_LB = 2.20462;
        private const double M_TO_IN = 39.3701;

        /// <summary>
        /// Creates ErpExportData from a collection of processed PartData objects.
        /// </summary>
        /// <param name="parts">Collection of processed part data.</param>
        /// <param name="parentPartNumber">Parent assembly/quote part number.</param>
        /// <param name="customer">Customer name for ERP records.</param>
        /// <param name="parentDescription">Description for the parent record.</param>
        /// <returns>Populated ErpExportData ready for export.</returns>
        public static ErpExportData FromPartDataCollection(
            IEnumerable<PartData> parts,
            string parentPartNumber,
            string customer,
            string parentDescription = "")
        {
            if (parts == null)
                throw new ArgumentNullException(nameof(parts));

            var data = new ErpExportData
            {
                ParentPartNumber = parentPartNumber ?? "",
                ParentDescription = parentDescription,
                ParentDrawing = "",
                ParentRevision = ""
            };

            int pieceNumber = 1;
            foreach (var pd in parts.Where(p => p != null && p.Status == ProcessingStatus.Success))
            {
                var erpPart = MapPartData(pd, customer);
                data.Parts.Add(erpPart);

                // Add BOM relationship if there's a parent
                if (!string.IsNullOrEmpty(parentPartNumber))
                {
                    data.BomRelationships.Add(new BomRelationship
                    {
                        ParentPartNumber = parentPartNumber,
                        ChildPartNumber = erpPart.PartNumber,
                        PieceNumber = pieceNumber.ToString(),
                        Quantity = Math.Max(1, pd.QuoteQty)
                    });
                    pieceNumber++;
                }
            }

            return data;
        }

        /// <summary>
        /// Creates ErpExportData for a single part (no parent assembly).
        /// </summary>
        public static ErpExportData FromSinglePart(PartData part, string customer)
        {
            if (part == null)
                throw new ArgumentNullException(nameof(part));

            var data = new ErpExportData();
            var erpPart = MapPartData(part, customer);
            data.Parts.Add(erpPart);
            return data;
        }

        private static ErpPartData MapPartData(PartData pd, string customer)
        {
            var erpPart = new ErpPartData
            {
                PartNumber = ExtractPartNumber(pd),
                Drawing = pd.PartName ?? "",
                Description = pd.Extra.TryGetValue("Description", out var desc) ? desc : "",
                Revision = pd.Extra.TryGetValue("Revision", out var rev) ? rev : "",
                Customer = customer ?? "",
                Quantity = Math.Max(1, pd.QuoteQty),
                IsAssembly = pd.Classification == PartType.Assembly,
                OptiMaterial = pd.OptiMaterial ?? pd.Material ?? "",
                RawWeight = pd.Mass_kg * KG_TO_LB
            };

            // Determine material type from classification
            switch (pd.Classification)
            {
                case PartType.SheetMetal:
                    erpPart.MaterialType = MaterialType.SheetMetal;
                    erpPart.F300Length = pd.BBoxLength_m > 0 ? pd.BBoxLength_m * M_TO_IN : 0;
                    break;
                case PartType.Tube:
                    erpPart.MaterialType = MaterialType.Tube;
                    erpPart.F300Length = pd.Tube.Length_m > 0 ? pd.Tube.Length_m * M_TO_IN : 0;
                    break;
                default:
                    erpPart.MaterialType = MaterialType.Generic;
                    break;
            }

            // Determine part type (Standard, Outsourced, Purchased, etc.)
            if (pd.IsPurchased)
            {
                erpPart.PartType = ErpPartType.Purchased;
                erpPart.PurchasedPartNumber = erpPart.PartNumber;
            }
            else
            {
                erpPart.PartType = ErpPartType.Standard;
            }

            // Map routing operations from cost data
            MapRoutingOperations(erpPart, pd);

            return erpPart;
        }

        private static void MapRoutingOperations(ErpPartData erpPart, PartData pd)
        {
            var cost = pd.Cost;

            // OP20 - Laser/Waterjet cutting
            if (cost.OP20_S_min > 0 || cost.OP20_R_min > 0 || cost.F115_Price > 0)
            {
                erpPart.Operations["OP20"] = new RoutingOperation
                {
                    WorkCenter = !string.IsNullOrEmpty(cost.OP20_WorkCenter) ? cost.OP20_WorkCenter : "F115",
                    OpNumber = 20,
                    Setup = cost.OP20_S_min / 60.0, // Convert minutes to hours for routing
                    Run = cost.OP20_R_min / 60.0,
                    Enabled = true
                };
            }

            // F210 - Deburr
            if (cost.F210_S_min > 0 || cost.F210_R_min > 0)
            {
                erpPart.Operations["F210"] = new RoutingOperation
                {
                    WorkCenter = "F210",
                    OpNumber = 30,
                    Setup = cost.F210_S_min / 60.0,
                    Run = cost.F210_R_min / 60.0,
                    Enabled = true
                };
            }

            // F140 - Press Brake
            if (cost.F140_S_min > 0 || cost.F140_R_min > 0)
            {
                erpPart.Operations["F140"] = new RoutingOperation
                {
                    WorkCenter = "F140",
                    OpNumber = 40,
                    Setup = cost.F140_S_min / 60.0,
                    Run = cost.F140_R_min / 60.0,
                    Enabled = true
                };
            }

            // F220 - Tapping
            if (cost.F220_S_min > 0 || cost.F220_R_min > 0 || cost.F220_RN > 0)
            {
                erpPart.Operations["F220"] = new RoutingOperation
                {
                    WorkCenter = "F220",
                    OpNumber = 50,
                    Setup = cost.F220_S_min / 60.0,
                    Run = cost.F220_R_min / 60.0,
                    Note = cost.F220_Note ?? (cost.F220_RN > 0 ? $"{cost.F220_RN} tapped holes" : ""),
                    Enabled = true
                };
            }

            // F325 - Roll Forming
            if (cost.F325_S_min > 0 || cost.F325_R_min > 0)
            {
                erpPart.Operations["F325"] = new RoutingOperation
                {
                    WorkCenter = "F325",
                    OpNumber = 60,
                    Setup = cost.F325_S_min / 60.0,
                    Run = cost.F325_R_min / 60.0,
                    Enabled = true
                };
            }
        }

        private static string ExtractPartNumber(PartData pd)
        {
            // Try to get part number from file name (without extension)
            if (!string.IsNullOrEmpty(pd.FilePath))
            {
                return Path.GetFileNameWithoutExtension(pd.FilePath);
            }

            // Fall back to part name
            if (!string.IsNullOrEmpty(pd.PartName))
            {
                // Remove extension if present
                var name = pd.PartName;
                if (name.EndsWith(".sldprt", StringComparison.OrdinalIgnoreCase) ||
                    name.EndsWith(".sldasm", StringComparison.OrdinalIgnoreCase))
                {
                    name = Path.GetFileNameWithoutExtension(name);
                }
                return name;
            }

            return "UNKNOWN";
        }
    }
}
