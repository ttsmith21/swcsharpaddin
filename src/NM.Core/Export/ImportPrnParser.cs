using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NM.Core.Export
{
    /// <summary>
    /// Parses VBA-produced Import.prn files for ERP integration.
    /// </summary>
    public static class ImportPrnParser
    {
        public static Dictionary<string, ImportPrnRecord> Parse(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Import.prn file not found: {filePath}");

            var lines = File.ReadAllLines(filePath);
            return ParseLines(lines);
        }

        public static Dictionary<string, ImportPrnRecord> ParseLines(string[] lines)
        {
            var records = new Dictionary<string, ImportPrnRecord>();
            var currentSection = "";
            var fieldNames = new List<string>();
            var inDataSection = false;

            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i].Trim();

                if (string.IsNullOrWhiteSpace(line))
                {
                    inDataSection = false;
                    continue;
                }

                if (line.StartsWith("DECL("))
                {
                    currentSection = line.Substring(5, 2); // Extract IM, PS, RT, RN, ML
                    fieldNames.Clear();
                    inDataSection = false;

                    // Field names are on the same line after "DECL(XX) ADD "
                    // e.g.: "DECL(IM) ADD IM-KEY IM-DRAWING IM-DESCR ..."
                    var tokens = Tokenize(line);
                    // Skip "DECL(XX)" and "ADD" tokens, keep the field names
                    for (int t = 0; t < tokens.Count; t++)
                    {
                        var tok = tokens[t];
                        if (tok.StartsWith("DECL(") || tok.Equals("ADD", StringComparison.OrdinalIgnoreCase))
                            continue;
                        fieldNames.Add(tok);
                    }
                    continue;
                }

                if (line == "END")
                {
                    inDataSection = true;
                    continue;
                }

                if (!inDataSection)
                    continue;

                // Data row
                var values = Tokenize(line);
                if (values.Count == 0) continue;

                var dataDict = new Dictionary<string, string>();
                for (int j = 0; j < Math.Min(fieldNames.Count, values.Count); j++)
                {
                    dataDict[fieldNames[j]] = values[j];
                }

                ProcessDataRow(currentSection, fieldNames.Count, dataDict, records);
            }

            return records;
        }

        private static void ProcessDataRow(string section, int fieldCount, Dictionary<string, string> data, Dictionary<string, ImportPrnRecord> records)
        {
            if (section == "IM")
            {
                ProcessImRecord(data, records);
            }
            else if (section == "PS")
            {
                // Differentiate by presence of PS-DIM-1 field: BOM (5 fields) vs Material (11+ fields)
                if (data.ContainsKey("PS-DIM-1"))
                {
                    ProcessMaterialRecord(data, records);
                }
                else
                {
                    ProcessBomRecord(data, records);
                }
            }
            else if (section == "RT")
            {
                ProcessRoutingRecord(data, records);
            }
            else if (section == "RN")
            {
                ProcessRoutingNoteRecord(data, records);
            }
        }

        private static void ProcessImRecord(Dictionary<string, string> data, Dictionary<string, ImportPrnRecord> records)
        {
            if (!data.ContainsKey("IM-KEY")) return;

            var partNumber = data["IM-KEY"];
            if (!records.ContainsKey(partNumber))
            {
                records[partNumber] = new ImportPrnRecord { PartNumber = partNumber };
            }

            var record = records[partNumber];
            record.Drawing = GetValue(data, "IM-DRAWING");
            record.Description = GetValue(data, "IM-DESCR");
            record.Revision = GetValue(data, "IM-REV");
            record.ImType = GetIntValue(data, "IM-TYPE");
            record.ImClass = GetIntValue(data, "IM-CLASS");
            record.Commodity = GetValue(data, "IM-COMMODITY");
            record.StdLot = GetIntValue(data, "IM-STD-LOT");
        }

        private static void ProcessBomRecord(Dictionary<string, string> data, Dictionary<string, ImportPrnRecord> records)
        {
            if (!data.ContainsKey("PS-SUBORD-KEY")) return;

            var childPartNumber = data["PS-SUBORD-KEY"];
            if (!records.ContainsKey(childPartNumber))
            {
                records[childPartNumber] = new ImportPrnRecord { PartNumber = childPartNumber };
            }

            var record = records[childPartNumber];
            record.ParentPartNumber = GetValue(data, "PS-PARENT-KEY");
            record.PieceNumber = GetValue(data, "PS-PIECE-NO");
            record.BomQuantity = GetDoubleValue(data, "PS-QTY-P");
        }

        private static void ProcessMaterialRecord(Dictionary<string, string> data, Dictionary<string, ImportPrnRecord> records)
        {
            if (!data.ContainsKey("PS-PARENT-KEY")) return;

            var partNumber = data["PS-PARENT-KEY"];
            if (!records.ContainsKey(partNumber))
            {
                records[partNumber] = new ImportPrnRecord { PartNumber = partNumber };
            }

            var record = records[partNumber];
            record.OptiMaterial = GetValue(data, "PS-SUBORD-KEY");
            record.RawWeight = GetDoubleValue(data, "PS-QTY-P");
            record.F300Length = GetDoubleValue(data, "PS-DIM-1");
        }

        private static void ProcessRoutingRecord(Dictionary<string, string> data, Dictionary<string, ImportPrnRecord> records)
        {
            if (!data.ContainsKey("RT-ITEM-KEY")) return;

            var partNumber = data["RT-ITEM-KEY"];
            if (!records.ContainsKey(partNumber))
            {
                records[partNumber] = new ImportPrnRecord { PartNumber = partNumber };
            }

            var record = records[partNumber];
            var routing = new RoutingStep
            {
                WorkCenter = GetValue(data, "RT-WORKCENTER-KEY"),
                OpNumber = GetIntValue(data, "RT-OP-NUM"),
                SetupHours = GetDoubleValue(data, "RT-SETUP"),
                RunHours = GetDoubleValue(data, "RT-RUN-STD")
            };

            record.Routing.Add(routing);
        }

        private static void ProcessRoutingNoteRecord(Dictionary<string, string> data, Dictionary<string, ImportPrnRecord> records)
        {
            if (!data.ContainsKey("RN-ITEM-KEY")) return;

            var partNumber = data["RN-ITEM-KEY"];
            if (!records.ContainsKey(partNumber))
            {
                records[partNumber] = new ImportPrnRecord { PartNumber = partNumber };
            }

            var record = records[partNumber];
            var opNumber = GetIntValue(data, "RN-OP-NUM");
            var noteText = GetValue(data, "RN-DESCR");

            if (!record.RoutingNotes.ContainsKey(opNumber))
            {
                record.RoutingNotes[opNumber] = new List<string>();
            }

            record.RoutingNotes[opNumber].Add(noteText);
        }

        private static List<string> Tokenize(string line)
        {
            var tokens = new List<string>();
            var currentToken = new StringBuilder();
            var inQuotes = false;
            var escapeNext = false;
            var hadQuotes = false; // Track quoted tokens to preserve empty strings ""

            for (int i = 0; i < line.Length; i++)
            {
                var ch = line[i];

                if (escapeNext)
                {
                    currentToken.Append(ch);
                    escapeNext = false;
                    continue;
                }

                if (ch == '\\' && i + 1 < line.Length && line[i + 1] == '"')
                {
                    escapeNext = true;
                    continue;
                }

                if (ch == '"')
                {
                    inQuotes = !inQuotes;
                    hadQuotes = true;
                    continue;
                }

                if (!inQuotes && (ch == ' ' || ch == '\t'))
                {
                    if (currentToken.Length > 0 || hadQuotes)
                    {
                        tokens.Add(currentToken.ToString());
                        currentToken.Clear();
                        hadQuotes = false;
                    }
                    continue;
                }

                currentToken.Append(ch);
            }

            if (currentToken.Length > 0 || hadQuotes)
            {
                tokens.Add(currentToken.ToString());
            }

            return tokens;
        }

        private static string GetValue(Dictionary<string, string> data, string key)
        {
            return data.ContainsKey(key) ? data[key] : "";
        }

        private static int GetIntValue(Dictionary<string, string> data, string key)
        {
            if (!data.ContainsKey(key)) return 0;
            int.TryParse(data[key], out int result);
            return result;
        }

        private static double GetDoubleValue(Dictionary<string, string> data, string key)
        {
            if (!data.ContainsKey(key)) return 0.0;
            double.TryParse(data[key], out double result);
            return result;
        }
    }

    /// <summary>
    /// Represents a single part record from Import.prn with all associated data.
    /// </summary>
    public class ImportPrnRecord
    {
        public string PartNumber { get; set; }
        public string Drawing { get; set; }
        public string Description { get; set; }
        public string Revision { get; set; }
        public int ImType { get; set; }
        public int ImClass { get; set; }
        public string Commodity { get; set; }
        public int StdLot { get; set; }

        // Material relationship (from second PS section)
        public string OptiMaterial { get; set; }
        public double RawWeight { get; set; }
        public double F300Length { get; set; }

        // BOM relationship (from first PS section)
        public double BomQuantity { get; set; }
        public string ParentPartNumber { get; set; }
        public string PieceNumber { get; set; }

        // Routing
        public List<RoutingStep> Routing { get; set; }
        public Dictionary<int, List<string>> RoutingNotes { get; set; }

        public ImportPrnRecord()
        {
            PartNumber = "";
            Drawing = "";
            Description = "";
            Revision = "";
            Commodity = "";
            OptiMaterial = "";
            ParentPartNumber = "";
            PieceNumber = "";
            Routing = new List<RoutingStep>();
            RoutingNotes = new Dictionary<int, List<string>>();
        }
    }

    /// <summary>
    /// Represents a single routing operation step.
    /// </summary>
    public class RoutingStep
    {
        public string WorkCenter { get; set; }
        public int OpNumber { get; set; }
        public double SetupHours { get; set; }
        public double RunHours { get; set; }

        public RoutingStep()
        {
            WorkCenter = "";
        }
    }
}
