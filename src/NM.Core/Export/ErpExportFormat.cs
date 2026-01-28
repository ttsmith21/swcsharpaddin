using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using NM.Core.DataModel;
using NM.Core.Utils;

namespace NM.Core.Export
{
    /// <summary>
    /// ERP Export format generator - creates Import.prn file for ERP system.
    /// Ported from VBA modExport.bas PopulateItemMaster, PopulateProductStructure, PopulateRouting.
    /// </summary>
    public sealed class ErpExportFormat
    {
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        // Default configuration
        public string DefaultLocation { get; set; } = "F"; // F=Northern, N=Nu, D=N2
        public int StandardLot { get; set; } = 1;
        public string Customer { get; set; } = "";

        /// <summary>
        /// Generates the full Import.prn file for ERP import.
        /// </summary>
        public void ExportToImportPrn(ErpExportData data, string outputPath)
        {
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");

            using (var sw = new StreamWriter(outputPath, false, Encoding.UTF8))
            {
                // Section 1: Item Master
                WriteItemMaster(sw, data);

                // Section 2: Material Locations
                WriteMaterialLocations(sw, data);

                // Section 3: Product Structure (BOM)
                WriteProductStructure(sw, data);

                // Section 4: Part Material Relationships
                WritePartMaterialRelationships(sw, data);

                // Section 5: Routing
                WriteRouting(sw, data);

                // Section 6: Routing Notes
                WriteRoutingNotes(sw, data);
            }
        }

        #region Item Master (IM)

        private void WriteItemMaster(StreamWriter sw, ErpExportData data)
        {
            sw.WriteLine("DECL(IM) ADD IM-KEY IM-DRAWING IM-DESCR IM-REV IM-TYPE IM-CLASS IM-CATALOG IM-COMMODITY IM-SAVE-DEMAND-SW IM-BUYER IM-STOCK-SW IM-PLAN-SW IM-STD-LOT IM-GL-ACCT IM-STD-MAT IM-ISSUE-SW");
            sw.WriteLine("END");

            // Parent assembly/part
            if (!string.IsNullOrEmpty(data.ParentPartNumber))
            {
                sw.WriteLine(FormatItemMaster(data.ParentPartNumber, data.ParentDrawing, data.ParentDescription,
                    data.ParentRevision, 1, 9, Customer, "F", StandardLot));

                // Parent OS if required
                if (!string.IsNullOrEmpty(data.ParentOsWorkCenter))
                {
                    string osNumber = $"{data.ParentOsWorkCenter}-{data.ParentPartNumber}";
                    sw.WriteLine(FormatItemMaster(osNumber, "", data.ParentDescription, "", 2, 1,
                        Customer, data.ParentOsWorkCenter, 1, "6112.1", 2));
                }
            }

            // Child parts
            foreach (var part in data.Parts)
            {
                int stdLot = StandardLot * Math.Max(1, part.Quantity);
                sw.WriteLine(FormatItemMaster(part.PartNumber, part.Drawing, part.Description,
                    part.Revision, 1, 9, part.Customer ?? Customer, "F", stdLot));

                // OS/MP/CUST parts if required
                if (part.RequiresOsPart && !string.IsNullOrEmpty(part.OsPartNumber))
                {
                    string glAcct = part.PartType == ErpPartType.Outsourced ? "6112.1" : "6110.1";
                    sw.WriteLine(FormatItemMaster(part.OsPartNumber, "", part.Description, "", 2, 1,
                        part.Customer ?? Customer, part.OsWorkCenter, 1, glAcct, 2));
                }
            }
        }

        private string FormatItemMaster(string partNumber, string drawing, string description, string revision,
            int imType, int imClass, string customer, string commodity, int stdLot,
            string glAcct = "", int issueSw = 0)
        {
            return $"{Q(partNumber)}{Q(drawing)}{Q(description)}{Q(revision)}{imType} {imClass} " +
                   $"{Q(customer)}{Q(commodity)}0 {Q("2014")}0 0 {stdLot} {Q(glAcct)}{0} {issueSw}";
        }

        #endregion

        #region Material Locations (ML)

        private void WriteMaterialLocations(StreamWriter sw, ErpExportData data)
        {
            sw.WriteLine();
            sw.WriteLine("DECL(ML) ML-IMKEY ML-LOCATION");
            sw.WriteLine("END");

            // Parent
            if (!string.IsNullOrEmpty(data.ParentPartNumber))
            {
                sw.WriteLine($"{Q(data.ParentPartNumber)}{Q("PRIMARY")}");
            }

            // All parts
            foreach (var part in data.Parts)
            {
                sw.WriteLine($"{Q(part.PartNumber)}{Q("PRIMARY")}");

                if (part.RequiresOsPart && !string.IsNullOrEmpty(part.OsPartNumber))
                {
                    sw.WriteLine($"{Q(part.OsPartNumber)}{Q("PRIMARY")}");
                }
            }
        }

        #endregion

        #region Product Structure (PS) - BOM

        private void WriteProductStructure(StreamWriter sw, ErpExportData data)
        {
            sw.WriteLine();
            sw.WriteLine("DECL(PS) ADD PS-PARENT-KEY PS-SUBORD-KEY PS-REV PS-PIECE-NO PS-QTY-P");
            sw.WriteLine("END");

            foreach (var rel in data.BomRelationships)
            {
                string pieceNo = StringUtils.PadItemNumber(rel.PieceNumber, 2);
                sw.WriteLine($"{Q(rel.ParentPartNumber)}{Q(rel.ChildPartNumber)}{Q("COMMON SET")}{Q(pieceNo)}{rel.Quantity}");
            }
        }

        #endregion

        #region Part Material Relationships

        private void WritePartMaterialRelationships(StreamWriter sw, ErpExportData data)
        {
            sw.WriteLine();
            sw.WriteLine("DECL(PS) ADD PS-PARENT-KEY PS-SUBORD-KEY PS-REV PS-PIECE-NO PS-QTY-P PS-DIM-1 PS-ISSUE-SW PS-BFLOCATION-SW PS-BFQTY-SW PS-BFZEROQTY-SW PS-OP-NUM");
            sw.WriteLine("END");

            foreach (var part in data.Parts)
            {
                if (part.IsAssembly) continue; // Skip assemblies for material relationships

                string optiMaterial = part.OptiMaterial ?? "";
                double rawWeight = part.RawWeight;
                double f300Length = part.F300Length;

                switch (part.PartType)
                {
                    case ErpPartType.Standard:
                    case ErpPartType.Outsourced:
                        if (part.MaterialType == MaterialType.SheetMetal)
                        {
                            // Sheet metal: use 1 as raw weight (sheet count), F300_Length for dimension
                            sw.WriteLine($"{Q(part.PartNumber)}{Q(optiMaterial)}{Q("COMMON SET")}{Q("01")}" +
                                       $"{1:0.####} {f300Length} 2 1 1 1 20");
                        }
                        else if (part.MaterialType == MaterialType.Insulation)
                        {
                            // Insulation: use Matl_SF as quantity
                            sw.WriteLine($"{Q(part.PartNumber)}{Q(optiMaterial)}{Q("COMMON SET")}{Q("01")}" +
                                       $"{part.MatlSf:0.####} 0 2 1 1 1 20");
                        }
                        else
                        {
                            // Generic: use raw weight
                            if (rawWeight <= 0) rawWeight = 0.0001;
                            sw.WriteLine($"{Q(part.PartNumber)}{Q(optiMaterial)}{Q("COMMON SET")}{Q("01")}" +
                                       $"{rawWeight:0.####} 0 2 1 1 1 20");
                        }

                        // Add OS material relationship if outsourced
                        if (part.PartType == ErpPartType.Outsourced && !string.IsNullOrEmpty(part.OsPartNumber))
                        {
                            sw.WriteLine($"{Q(part.PartNumber)}{Q(part.OsPartNumber)}{Q("COMMON SET")}{Q("0")}1 0 2 1 1 1 20");
                        }
                        break;

                    case ErpPartType.Machined:
                        if (!string.IsNullOrEmpty(part.MpPartNumber))
                        {
                            sw.WriteLine($"{Q(part.PartNumber)}{Q(part.MpPartNumber)}{Q("COMMON SET")}{Q("01")}1 0 2 1 1 1 20");
                        }
                        break;

                    case ErpPartType.Purchased:
                        if (!string.IsNullOrEmpty(part.PurchasedPartNumber))
                        {
                            sw.WriteLine($"{Q(part.PartNumber)}{Q(part.PurchasedPartNumber)}{Q("COMMON SET")}{Q("01")}1 0 2 1 1 1 10");
                        }
                        break;

                    case ErpPartType.CustomerSupplied:
                        if (!string.IsNullOrEmpty(part.CustPartNumber))
                        {
                            string custNum = part.CustPartNumber;
                            if (custNum.Length >= 5 && !custNum.StartsWith("CUST-", StringComparison.OrdinalIgnoreCase))
                                custNum = "CUST-" + custNum;
                            sw.WriteLine($"{Q(part.PartNumber)}{Q(custNum)}{Q("COMMON SET")}{Q("01")}1 0 2 1 1 1 10");
                        }
                        break;
                }
            }
        }

        #endregion

        #region Routing (RT)

        private void WriteRouting(StreamWriter sw, ErpExportData data)
        {
            sw.WriteLine();
            sw.WriteLine("DECL(RT) ADD RT-ITEM-KEY RT-WORKCENTER-KEY RT-OP-NUM RT-SETUP RT-RUN-STD RT-REV RT-MULT-SEQ");
            sw.WriteLine("END");

            foreach (var part in data.Parts)
            {
                if (part.IsAssembly) continue;

                string location = part.Location ?? DefaultLocation;

                // OP20 - Cutting/Primary operation
                if (part.Operations.TryGetValue("OP20", out var op20))
                {
                    string wc = op20.WorkCenter;
                    if (wc.Length >= 4)
                    {
                        location = wc.Substring(0, 1);
                        wc = wc.Substring(1, 3);
                    }

                    // Special handling for certain work centers (add OP10 offline program)
                    if (wc == "105" || wc == "115" || wc == "120" || wc == "125" || wc == "135")
                    {
                        sw.WriteLine($"{Q(part.PartNumber)}{Q("O" + wc)}10 0 0 {Q("COMMON SET")}0");
                    }

                    sw.WriteLine($"{Q(part.PartNumber)}{Q(location + wc)}20 {op20.Setup:0.####} {op20.Run:0.####} {Q("COMMON SET")}0");
                }

                // OP30 - Deburr (F210)
                if (part.Operations.TryGetValue("F210", out var f210) && f210.Enabled)
                {
                    // Deburr not at N2
                    string deburrLocation = (location == "D") ? "F" : location;
                    sw.WriteLine($"{Q(part.PartNumber)}{Q(deburrLocation + "210")}30 {f210.Setup:0.####} {f210.Run:0.####} {Q("COMMON SET")}0");
                }

                // OP40 - Press Brake (F140)
                if (part.Operations.TryGetValue("F140", out var f140) && f140.Enabled)
                {
                    sw.WriteLine($"{Q(part.PartNumber)}{Q(location + "140")}40 {f140.Setup:0.####} {f140.Run:0.####} {Q("COMMON SET")}0");
                }

                // OP50 - Tapping (F220)
                if (part.Operations.TryGetValue("F220", out var f220) && f220.Enabled)
                {
                    sw.WriteLine($"{Q(part.PartNumber)}{Q(location + "220")}50 {f220.Setup:0.####} {f220.Run:0.####} {Q("COMMON SET")}0");
                }

                // OP60 - Forming (F325)
                if (part.Operations.TryGetValue("F325", out var f325) && f325.Enabled)
                {
                    sw.WriteLine($"{Q(part.PartNumber)}{Q(location + "325")}60 {f325.Setup:0.####} {f325.Run:0.####} {Q("COMMON SET")}0");
                }

                // Additional operations (F400 Weld, F385 Assembly, etc.)
                int opNum = 70;
                foreach (var op in part.Operations.Values.Where(o => o.Enabled && o.OpNumber == 0))
                {
                    sw.WriteLine($"{Q(part.PartNumber)}{Q(location + op.WorkCenter)}{ opNum} {op.Setup:0.####} {op.Run:0.####} {Q("COMMON SET")}0");
                    opNum += 10;
                }
            }
        }

        #endregion

        #region Routing Notes (RN)

        private void WriteRoutingNotes(StreamWriter sw, ErpExportData data)
        {
            var notesToWrite = data.Parts
                .SelectMany(p => p.Operations.Values)
                .Where(op => !string.IsNullOrEmpty(op.Note))
                .ToList();

            if (notesToWrite.Count == 0)
                return;

            sw.WriteLine();
            sw.WriteLine("DECL(RN) ADD RN-ITEM-KEY RN-OP-NUM RN-LINE-NO RN-NOTE");
            sw.WriteLine("END");

            foreach (var part in data.Parts)
            {
                foreach (var op in part.Operations.Values.Where(o => !string.IsNullOrEmpty(o.Note)))
                {
                    // Split long notes into multiple lines (max 30 chars per VBA)
                    var lines = SplitNote(op.Note, 30);
                    int lineNo = 1;
                    foreach (var line in lines)
                    {
                        sw.WriteLine($"{Q(part.PartNumber)}{op.OpNumber} {lineNo} {Q(line)}");
                        lineNo++;
                    }
                }
            }
        }

        private static IEnumerable<string> SplitNote(string note, int maxLength)
        {
            if (string.IsNullOrEmpty(note))
                yield break;

            for (int i = 0; i < note.Length; i += maxLength)
            {
                yield return note.Substring(i, Math.Min(maxLength, note.Length - i));
            }
        }

        #endregion

        #region Helpers

        /// <summary>
        /// Wraps string in quotes with trailing space (QuoteMe from VBA).
        /// </summary>
        private static string Q(string value)
        {
            return "\"" + (value ?? "") + "\" ";
        }

        #endregion
    }

    #region Data Models

    /// <summary>
    /// Container for all ERP export data.
    /// </summary>
    public sealed class ErpExportData
    {
        public string ParentPartNumber { get; set; }
        public string ParentDrawing { get; set; }
        public string ParentDescription { get; set; }
        public string ParentRevision { get; set; }
        public string ParentOsWorkCenter { get; set; }

        public List<ErpPartData> Parts { get; } = new List<ErpPartData>();
        public List<BomRelationship> BomRelationships { get; } = new List<BomRelationship>();
    }

    public sealed class ErpPartData
    {
        public string PartNumber { get; set; }
        public string Drawing { get; set; }
        public string Description { get; set; }
        public string Revision { get; set; }
        public string Customer { get; set; }
        public string Location { get; set; }
        public int Quantity { get; set; } = 1;

        public bool IsAssembly { get; set; }
        public ErpPartType PartType { get; set; } = ErpPartType.Standard;
        public MaterialType MaterialType { get; set; } = MaterialType.Generic;

        public string OptiMaterial { get; set; }
        public double RawWeight { get; set; }
        public double F300Length { get; set; }
        public double MatlSf { get; set; }

        // OS/MP/Purchased part info
        public bool RequiresOsPart { get; set; }
        public string OsPartNumber { get; set; }
        public string OsWorkCenter { get; set; }
        public string MpPartNumber { get; set; }
        public string PurchasedPartNumber { get; set; }
        public string CustPartNumber { get; set; }

        public Dictionary<string, RoutingOperation> Operations { get; } = new Dictionary<string, RoutingOperation>(StringComparer.OrdinalIgnoreCase);
    }

    public sealed class RoutingOperation
    {
        public string WorkCenter { get; set; }
        public int OpNumber { get; set; }
        public double Setup { get; set; }
        public double Run { get; set; }
        public string Note { get; set; }
        public bool Enabled { get; set; } = true;
    }

    public sealed class BomRelationship
    {
        public string ParentPartNumber { get; set; }
        public string ChildPartNumber { get; set; }
        public string PieceNumber { get; set; }
        public int Quantity { get; set; }
    }

    public enum ErpPartType
    {
        Standard = 0,
        Machined = 1,
        Outsourced = 2,
        Purchased = 3,
        CustomerSupplied = 4
    }

    public enum MaterialType
    {
        Generic = 0,
        SheetMetal = 1,
        Insulation = 2,
        Tube = 3
    }

    #endregion
}
