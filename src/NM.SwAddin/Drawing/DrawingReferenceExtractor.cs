using System;
using System.Collections.Generic;
using NM.Core;
using NM.Core.Models;
using NM.SwAddin.AssemblyProcessing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Drawing
{
    /// <summary>
    /// Extracts model references from drawing views.
    /// Ported from VBA drawing traversal logic.
    /// </summary>
    public sealed class DrawingReferenceExtractor
    {
        private readonly ISldWorks _swApp;

        public DrawingReferenceExtractor(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Extracts all unique models referenced by views in a drawing.
        /// For assemblies, recursively extracts all part components.
        /// </summary>
        /// <param name="drawingDoc">The drawing document.</param>
        /// <returns>List of unique SwModelInfo for all referenced parts.</returns>
        public List<SwModelInfo> ExtractReferences(IModelDoc2 drawingDoc)
        {
            const string proc = nameof(DrawingReferenceExtractor) + ".ExtractReferences";
            ErrorHandler.PushCallStack(proc);

            var models = new List<SwModelInfo>();
            var seenPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var drawing = drawingDoc as IDrawingDoc;
                if (drawing == null)
                {
                    ErrorHandler.DebugLog("ExtractReferences: Not a drawing document");
                    return models;
                }

                // Get all sheets
                var sheetNames = drawing.GetSheetNames() as string[];
                if (sheetNames == null || sheetNames.Length == 0)
                {
                    ErrorHandler.DebugLog("ExtractReferences: No sheets found");
                    return models;
                }

                foreach (var sheetName in sheetNames)
                {
                    var sheet = drawing.Sheet[sheetName] as ISheet;
                    if (sheet == null) continue;

                    // Get views on this sheet
                    var viewsRaw = sheet.GetViews();
                    if (viewsRaw == null) continue;

                    var views = viewsRaw as object[];
                    if (views == null) continue;

                    foreach (var viewObj in views)
                    {
                        var view = viewObj as IView;
                        if (view == null) continue;

                        ExtractFromView(view, models, seenPaths);
                    }
                }

                ErrorHandler.DebugLog($"ExtractReferences: Found {models.Count} unique models from drawing");
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Failed to extract drawing references", ex, ErrorHandler.LogLevel.Error);
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }

            return models;
        }

        private void ExtractFromView(IView view, List<SwModelInfo> models, HashSet<string> seenPaths)
        {
            try
            {
                // Get referenced document from view
                var refDoc = view.ReferencedDocument as IModelDoc2;
                if (refDoc == null) return;

                string refPath = refDoc.GetPathName();
                if (string.IsNullOrWhiteSpace(refPath)) return;

                string refConfig = view.ReferencedConfiguration ?? string.Empty;
                var docType = (swDocumentTypes_e)refDoc.GetType();

                if (docType == swDocumentTypes_e.swDocPART)
                {
                    // Direct part reference
                    string key = BuildKey(refPath, refConfig);
                    if (!seenPaths.Contains(key))
                    {
                        seenPaths.Add(key);
                        models.Add(new SwModelInfo(refPath, refConfig) { ModelDoc = refDoc });
                    }
                }
                else if (docType == swDocumentTypes_e.swDocASSEMBLY)
                {
                    // Assembly - extract all components
                    ExtractAssemblyComponents(refDoc, refConfig, models, seenPaths);
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"ExtractFromView failed: {ex.Message}");
            }
        }

        private void ExtractAssemblyComponents(IModelDoc2 assyDoc, string configName, List<SwModelInfo> models, HashSet<string> seenPaths)
        {
            try
            {
                var assyDocTyped = assyDoc as IAssemblyDoc;
                if (assyDocTyped == null) return;

                // Use ComponentCollector to get all unique part components from assembly
                var collector = new ComponentCollector();
                var result = collector.CollectUniqueComponents(assyDocTyped);

                foreach (var modelInfo in result.ValidComponents)
                {
                    if (modelInfo == null) continue;
                    if (string.IsNullOrWhiteSpace(modelInfo.FilePath)) continue;

                    // Only include parts
                    if (modelInfo.Type != SwModelInfo.ModelType.Part) continue;

                    string key = BuildKey(modelInfo.FilePath, modelInfo.Configuration);

                    if (!seenPaths.Contains(key))
                    {
                        seenPaths.Add(key);
                        models.Add(modelInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"ExtractAssemblyComponents failed: {ex.Message}");
            }
        }

        private static string BuildKey(string path, string config)
        {
            return $"{path}|{config}".ToLowerInvariant();
        }
    }
}
