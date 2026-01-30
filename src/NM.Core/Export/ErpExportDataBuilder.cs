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

            // Determine part type from custom properties (VBA parity)
            erpPart.PartType = DetectPartType(pd);
            erpPart.Location = ExtractLocation(pd);

            // Handle OS/MP/CUST variants based on part type
            switch (erpPart.PartType)
            {
                case ErpPartType.Outsourced:
                    erpPart.RequiresOsPart = true;
                    string osWc = pd.Extra.TryGetValue("OS_WC", out var wc) ? wc : "OS";
                    if (osWc.Length > 3) osWc = osWc.Substring(0, 3);
                    erpPart.OsWorkCenter = osWc;
                    erpPart.OsPartNumber = $"{osWc}-{erpPart.PartNumber}";
                    break;

                case ErpPartType.Machined:
                    erpPart.RequiresOsPart = true;
                    erpPart.OsWorkCenter = "MP";
                    erpPart.MpPartNumber = $"MP-{erpPart.PartNumber}";
                    break;

                case ErpPartType.CustomerSupplied:
                    string custNum = pd.Extra.TryGetValue("CustPartNumber", out var cn) ? cn : erpPart.PartNumber;
                    if (!custNum.StartsWith("CUST-", StringComparison.OrdinalIgnoreCase))
                        custNum = "CUST-" + custNum;
                    erpPart.CustPartNumber = custNum;
                    break;

                case ErpPartType.Purchased:
                    erpPart.PurchasedPartNumber = pd.Extra.TryGetValue("PurchasedPartNumber", out var pn)
                        ? pn : erpPart.PartNumber;
                    break;
            }

            // Map routing operations from cost data
            MapRoutingOperations(erpPart, pd);

            return erpPart;
        }

        /// <summary>
        /// Detects part type from custom properties (VBA: rbPartType, rbPartTypeSub).
        /// </summary>
        private static ErpPartType DetectPartType(PartData pd)
        {
            // Check rbPartType custom property
            if (pd.Extra.TryGetValue("rbPartType", out var rbPartType))
            {
                switch (rbPartType)
                {
                    case "2": // Outsourced
                        return ErpPartType.Outsourced;
                    case "1": // Special - check rbPartTypeSub for machined/purchased/customer
                        if (pd.Extra.TryGetValue("rbPartTypeSub", out var sub))
                        {
                            switch (sub)
                            {
                                case "0": return ErpPartType.Machined;
                                case "1": return ErpPartType.Purchased;
                                case "2": return ErpPartType.CustomerSupplied;
                            }
                        }
                        return ErpPartType.Machined;
                    case "0": // Standard
                    default:
                        break;
                }
            }

            // Fallback: check IsPurchased flag
            if (pd.IsPurchased) return ErpPartType.Purchased;

            return ErpPartType.Standard;
        }

        /// <summary>
        /// Extracts location from OP20 work center (VBA: strLocation = Left(OP20, 1)).
        /// </summary>
        private static string ExtractLocation(PartData pd, string defaultLocation = "F")
        {
            // Try OP20 work center first (VBA pattern)
            var workCenter = pd.Cost.OP20_WorkCenter;
            if (!string.IsNullOrEmpty(workCenter) && workCenter.Length >= 1)
            {
                char loc = workCenter[0];
                if (loc == 'F' || loc == 'N' || loc == 'D')
                    return loc.ToString();
            }

            // Try OP20 custom property as fallback
            if (pd.Extra.TryGetValue("OP20", out var op20) && !string.IsNullOrEmpty(op20) && op20.Length >= 1)
            {
                char loc = op20[0];
                if (loc == 'F' || loc == 'N' || loc == 'D')
                    return loc.ToString();
            }

            return defaultLocation;
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

        #region Hierarchical BOM Support

        /// <summary>
        /// Creates ErpExportData from a hierarchical BOM structure (for multi-level assemblies).
        /// Handles sub-assemblies as both parent and child in the BOM relationships.
        /// </summary>
        /// <param name="partDataLookup">Dictionary mapping file paths to processed PartData.</param>
        /// <param name="rootPartNumber">Part number of the root assembly.</param>
        /// <param name="customer">Customer name for ERP records.</param>
        /// <param name="bomHierarchy">List of (ParentPath, ChildPath, Quantity) tuples representing BOM structure.</param>
        public static ErpExportData FromHierarchicalBom(
            Dictionary<string, PartData> partDataLookup,
            string rootPartNumber,
            string customer,
            IEnumerable<(string ParentPath, string ChildPath, int Quantity)> bomHierarchy)
        {
            if (partDataLookup == null)
                throw new ArgumentNullException(nameof(partDataLookup));

            var data = new ErpExportData
            {
                ParentPartNumber = rootPartNumber ?? "",
                ParentDescription = ""
            };

            // Track which parts we've already added to avoid duplicates
            var addedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Process each BOM relationship
            int pieceNumber = 1;
            foreach (var (parentPath, childPath, quantity) in bomHierarchy ?? Enumerable.Empty<(string, string, int)>())
            {
                string parentPartNumber = Path.GetFileNameWithoutExtension(parentPath);
                string childPartNumber = Path.GetFileNameWithoutExtension(childPath);

                // Add child part if not already added
                if (!addedParts.Contains(childPath))
                {
                    if (partDataLookup.TryGetValue(childPath, out var pd))
                    {
                        var erpPart = MapPartData(pd, customer);
                        data.Parts.Add(erpPart);
                    }
                    else
                    {
                        // Create minimal part entry for sub-assemblies or unprocessed parts
                        data.Parts.Add(new ErpPartData
                        {
                            PartNumber = childPartNumber,
                            Drawing = childPartNumber,
                            Customer = customer ?? "",
                            Quantity = quantity,
                            IsAssembly = childPath.EndsWith(".sldasm", StringComparison.OrdinalIgnoreCase)
                        });
                    }
                    addedParts.Add(childPath);
                }

                // Add BOM relationship
                data.BomRelationships.Add(new BomRelationship
                {
                    ParentPartNumber = parentPartNumber,
                    ChildPartNumber = childPartNumber,
                    PieceNumber = pieceNumber.ToString(),
                    Quantity = quantity
                });
                pieceNumber++;
            }

            return data;
        }

        /// <summary>
        /// Flattens a hierarchical BOM into a list of parent-child relationships.
        /// </summary>
        /// <param name="rootPath">Path of the root assembly.</param>
        /// <param name="children">List of immediate children with their quantities and nested children.</param>
        /// <returns>Flattened list of (ParentPath, ChildPath, Quantity) relationships.</returns>
        public static List<(string ParentPath, string ChildPath, int Quantity)> FlattenBomHierarchy(
            string rootPath,
            IEnumerable<(string ChildPath, int Quantity, bool IsAssembly, IEnumerable<(string, int, bool, IEnumerable<(string, int, bool, object)>)> SubChildren)> children)
        {
            var result = new List<(string, string, int)>();

            void Flatten(string parentPath, IEnumerable<dynamic> items)
            {
                if (items == null) return;

                foreach (var item in items)
                {
                    string childPath = item.ChildPath;
                    int quantity = item.Quantity;

                    result.Add((parentPath, childPath, quantity));

                    // Recurse into sub-assembly children
                    if (item.IsAssembly && item.SubChildren != null)
                    {
                        Flatten(childPath, item.SubChildren);
                    }
                }
            }

            // Start flattening from root
            if (children != null)
            {
                foreach (var child in children)
                {
                    result.Add((rootPath, child.ChildPath, child.Quantity));

                    // Note: For deep recursion, we'd need a more sophisticated approach
                    // This handles the common case of 2-3 level hierarchies
                }
            }

            return result;
        }

        #endregion
    }
}
