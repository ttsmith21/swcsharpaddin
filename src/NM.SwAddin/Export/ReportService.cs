using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NM.Core;
using NM.Core.Export;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Export
{
    /// <summary>
    /// Service for generating reports from SolidWorks documents.
    /// Wraps ReportGenerator with SolidWorks-specific data extraction.
    /// Ported from VBA SP.bas Report() and ReportPart() functions.
    /// </summary>
    public sealed class ReportService
    {
        private readonly ISldWorks _swApp;
        private readonly ReportGenerator _generator;

        public ReportService(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
            _generator = new ReportGenerator();
        }

        /// <summary>
        /// Generates a report for an assembly, extracting data from BOM.
        /// Equivalent to VBA Report() function.
        /// </summary>
        public bool GenerateAssemblyReport(string assemblyPath, string outputPath)
        {
            const string proc = nameof(GenerateAssemblyReport);
            ErrorHandler.PushCallStack(proc);

            try
            {
                int errors = 0;
                int warnings = 0;

                // Determine document type
                var docType = assemblyPath.EndsWith(".sldasm", StringComparison.OrdinalIgnoreCase)
                    ? swDocumentTypes_e.swDocASSEMBLY
                    : swDocumentTypes_e.swDocPART;

                var model = _swApp.OpenDoc6(
                    assemblyPath,
                    (int)docType,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                    "",
                    ref errors,
                    ref warnings) as IModelDoc2;

                if (model == null)
                {
                    ErrorHandler.DebugLog($"{proc}: Failed to open {assemblyPath}");
                    return false;
                }

                var reportData = new ReportGenerator.AssemblyReportData
                {
                    AssemblyName = Path.GetFileNameWithoutExtension(assemblyPath),
                    FilePath = assemblyPath
                };

                if (docType == swDocumentTypes_e.swDocASSEMBLY)
                {
                    // Extract BOM data
                    ExtractBomData(model, reportData);
                }
                else
                {
                    // Single part - just add it
                    var partData = ExtractPartData(model);
                    if (partData != null)
                        reportData.Parts.Add(partData);
                }

                // Generate the report
                _generator.GenerateAssemblyReport(reportData, outputPath);

                _swApp.CloseAllDocuments(true);
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Generates a report for all parts in a folder.
        /// Equivalent to VBA ReportPart() function.
        /// </summary>
        public bool GenerateFolderReport(string folderPath, string outputPath)
        {
            const string proc = nameof(GenerateFolderReport);
            ErrorHandler.PushCallStack(proc);

            try
            {
                var partFiles = Directory.GetFiles(folderPath, "*.sldprt", SearchOption.TopDirectoryOnly);
                if (partFiles.Length == 0)
                {
                    ErrorHandler.DebugLog($"{proc}: No part files found in {folderPath}");
                    return false;
                }

                var reportData = new ReportGenerator.FolderReportData
                {
                    FolderPath = folderPath
                };

                foreach (var partFile in partFiles)
                {
                    int errors = 0;
                    int warnings = 0;

                    var model = _swApp.OpenDoc6(
                        partFile,
                        (int)swDocumentTypes_e.swDocPART,
                        (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                        "",
                        ref errors,
                        ref warnings) as IModelDoc2;

                    if (model == null) continue;

                    var partData = ExtractPartData(model);
                    if (partData != null)
                    {
                        partData.FilePath = partFile;
                        reportData.Parts.Add(partData);
                    }
                }

                _generator.GenerateFolderReport(reportData, outputPath);

                _swApp.CloseAllDocuments(true);
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        private void ExtractBomData(IModelDoc2 model, ReportGenerator.AssemblyReportData reportData)
        {
            var ext = model.Extension;
            if (ext == null) return;

            // Insert a temporary BOM table to get component list
            var bomTable = ext.InsertBomTable3(
                "",  // Use default template
                0, 0,
                (int)swBomType_e.swBomType_PartsOnly,
                "Default",
                false,
                (int)swNumberingType_e.swIndentedBOMNotSet,
                false) as IBomTableAnnotation;

            if (bomTable == null)
            {
                // Fallback: traverse assembly directly
                ExtractFromAssemblyTraversal(model, reportData);
                return;
            }

            try
            {
                var table = bomTable as ITableAnnotation;
                if (table == null) return;

                int rowCount = table.RowCount;
                string folderPath = Path.GetDirectoryName(model.GetPathName()) ?? "";

                var processedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (int row = 1; row < rowCount; row++)  // Skip header row
                {
                    string itemNumber = "";
                    string modelName = "";
                    int compCount = bomTable.GetComponentsCount2(row, "Default", out itemNumber, out modelName);

                    if (string.IsNullOrEmpty(modelName)) continue;
                    if (processedNames.Contains(modelName)) continue;
                    processedNames.Add(modelName);

                    // Try to open the part
                    string partPath = Path.Combine(folderPath, modelName + ".sldprt");
                    if (!File.Exists(partPath)) continue;

                    int errors = 0;
                    int warnings = 0;
                    var partModel = _swApp.OpenDoc6(
                        partPath,
                        (int)swDocumentTypes_e.swDocPART,
                        (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                        "",
                        ref errors,
                        ref warnings) as IModelDoc2;

                    if (partModel == null) continue;

                    var partData = ExtractPartData(partModel);
                    if (partData != null)
                    {
                        partData.Quantity = compCount > 0 ? compCount : 1;
                        partData.FilePath = partPath;
                        reportData.Parts.Add(partData);
                    }
                }

                // Delete the temporary BOM
                var feat = bomTable.BomFeature;
                if (feat != null)
                {
                    model.Extension.SelectByID2(feat.Name, "BOMFEATURE", 0, 0, 0, false, 0, null, 0);
                    model.EditDelete();
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"BOM extraction error: {ex.Message}");
            }
        }

        private void ExtractFromAssemblyTraversal(IModelDoc2 model, ReportGenerator.AssemblyReportData reportData)
        {
            var config = model.ConfigurationManager?.ActiveConfiguration;
            if (config == null) return;

            var rootComp = config.GetRootComponent3(true) as IComponent2;
            if (rootComp == null) return;

            var partCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            CollectComponents(rootComp, partCounts);

            string folderPath = Path.GetDirectoryName(model.GetPathName()) ?? "";

            foreach (var kvp in partCounts)
            {
                string partPath = kvp.Key;
                int qty = kvp.Value;

                if (!File.Exists(partPath)) continue;
                if (!partPath.EndsWith(".sldprt", StringComparison.OrdinalIgnoreCase)) continue;

                int errors = 0;
                int warnings = 0;
                var partModel = _swApp.OpenDoc6(
                    partPath,
                    (int)swDocumentTypes_e.swDocPART,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                    "",
                    ref errors,
                    ref warnings) as IModelDoc2;

                if (partModel == null) continue;

                var partData = ExtractPartData(partModel);
                if (partData != null)
                {
                    partData.Quantity = qty;
                    partData.FilePath = partPath;
                    reportData.Parts.Add(partData);
                }
            }
        }

        private void CollectComponents(IComponent2 comp, Dictionary<string, int> partCounts)
        {
            if (comp == null) return;

            var childrenObj = comp.GetChildren();
            if (childrenObj == null) return;

            var children = childrenObj as object[];
            if (children == null) return;

            foreach (var childObj in children)
            {
                var child = childObj as IComponent2;
                if (child == null || child.IsSuppressed()) continue;

                string childPath = child.GetPathName();
                if (string.IsNullOrEmpty(childPath)) continue;

                if (partCounts.ContainsKey(childPath))
                    partCounts[childPath]++;
                else
                    partCounts[childPath] = 1;

                // Recurse
                CollectComponents(child, partCounts);
            }
        }

        private ReportGenerator.PartReportData ExtractPartData(IModelDoc2 model)
        {
            if (model == null) return null;

            string partName = model.GetTitle() ?? "";
            if (partName.Contains("."))
                partName = Path.GetFileNameWithoutExtension(partName);

            // Extract custom properties
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var propMgr = model.Extension?.CustomPropertyManager[""];

            if (propMgr != null)
            {
                var names = propMgr.GetNames() as string[];
                if (names != null)
                {
                    foreach (var name in names)
                    {
                        string valOut = "";
                        string resolvedOut = "";
                        bool wasResolved = false;
                        propMgr.Get5(name, true, out valOut, out resolvedOut, out wasResolved);
                        props[name] = resolvedOut ?? valOut ?? "";
                    }
                }
            }

            // Also get config-specific properties
            string configName = model.ConfigurationManager?.ActiveConfiguration?.Name ?? "";
            if (!string.IsNullOrEmpty(configName))
            {
                var configPropMgr = model.Extension?.CustomPropertyManager[configName];
                if (configPropMgr != null)
                {
                    var names = configPropMgr.GetNames() as string[];
                    if (names != null)
                    {
                        foreach (var name in names)
                        {
                            if (props.ContainsKey(name)) continue;  // File-level takes precedence
                            string valOut = "";
                            string resolvedOut = "";
                            bool wasResolved = false;
                            configPropMgr.Get5(name, true, out valOut, out resolvedOut, out wasResolved);
                            props[name] = resolvedOut ?? valOut ?? "";
                        }
                    }
                }
            }

            return ReportGenerator.FromCustomProperties(partName, props);
        }
    }
}
