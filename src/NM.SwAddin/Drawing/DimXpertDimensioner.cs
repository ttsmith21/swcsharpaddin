using System;
using System.Collections.Generic;
using System.Linq;
using NM.Core;
using NM.SwAddin.Geometry;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swdimxpert;

namespace NM.SwAddin.Drawing
{
    /// <summary>
    /// Runs DimXpert auto-dimensioning on a part model to recognize manufacturing features
    /// (holes, slots, pockets, bosses) from B-Rep geometry. Works on STEP imports because
    /// DimXpert analyzes solid body topology, not the feature tree.
    /// </summary>
    public sealed class DimXpertDimensioner
    {
        private readonly ISldWorks _swApp;

        public DimXpertDimensioner(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Result of running DimXpert auto-dimension scheme on a part.
        /// </summary>
        public sealed class DimXpertResult
        {
            public bool Success { get; set; }
            public int FeaturesRecognized { get; set; }
            public int AnnotationsCreated { get; set; }
            public int AnnotationsImported { get; set; }
            public string Message { get; set; }
            public List<string> RecognizedFeatureTypes { get; set; } = new List<string>();
        }

        /// <summary>
        /// Runs the DimXpert Auto Dimension Scheme on the part model.
        /// Call BEFORE creating the drawing so annotations exist on the 3D model.
        /// </summary>
        public DimXpertResult RunAutoScheme(IModelDoc2 partModel)
        {
            const string proc = nameof(RunAutoScheme);
            ErrorHandler.PushCallStack(proc);
            PerformanceTracker.Instance.StartTimer("DimXpert_RunAutoScheme");

            var result = new DimXpertResult();

            try
            {
                if (partModel == null)
                {
                    result.Message = "Part model is null";
                    return result;
                }

                // Get the active configuration name
                string configName = partModel.ConfigurationManager?.ActiveConfiguration?.Name ?? "Default";

                // Access DimXpertManager — the indexer takes (configName, createIfMissing)
                var dimXpertMgr = partModel.Extension.get_DimXpertManager(configName, true) as DimXpertManager;
                if (dimXpertMgr == null)
                {
                    result.Message = "DimXpertManager not available (swdimxpert DLL may be missing)";
                    return result;
                }

                var dxPart = dimXpertMgr.DimXpertPart as IDimXpertPart;
                if (dxPart == null)
                {
                    result.Message = "IDimXpertPart not available";
                    return result;
                }

                // Select datum faces
                var datumSelector = new DatumFaceSelector();
                var datums = datumSelector.SelectDatums(partModel);
                if (!datums.Success)
                {
                    result.Message = "Could not auto-select datum faces: " + datums.Message;
                    return result;
                }

                // Pre-select datum entities with marks for DimXpert
                var selMgr = partModel.SelectionManager as ISelectionMgr;
                if (selMgr == null)
                {
                    result.Message = "SelectionManager not available";
                    return result;
                }

                partModel.ClearSelection2(true);
                var selData = selMgr.CreateSelectData() as ISelectData;
                if (selData == null)
                {
                    result.Message = "Cannot create SelectData";
                    return result;
                }

                // Primary datum: Mark 1
                selData.Mark = 1;
                bool sel = ((IEntity)datums.PrimaryFace).Select4(false, selData);
                if (!sel)
                {
                    result.Message = "Failed to select primary datum face";
                    return result;
                }

                // Secondary datum: Mark 2 (face or edge fallback)
                selData.Mark = 2;
                if (datums.SecondaryFace != null)
                    ((IEntity)datums.SecondaryFace).Select4(true, selData);
                else if (datums.SecondaryEdge != null)
                    ((IEntity)datums.SecondaryEdge).Select4(true, selData);

                // Tertiary datum: Mark 4 (face or edge fallback)
                selData.Mark = 4;
                if (datums.TertiaryFace != null)
                    ((IEntity)datums.TertiaryFace).Select4(true, selData);
                else if (datums.TertiaryEdge != null)
                    ((IEntity)datums.TertiaryEdge).Select4(true, selData);

                // Configure and run auto-dimension scheme
                var schemeOption = dxPart.GetAutoDimSchemeOption() as IDimXpertAutoDimSchemeOption;
                if (schemeOption == null)
                {
                    result.Message = "Cannot get auto-dimension scheme options";
                    return result;
                }

                ErrorHandler.DebugLog($"[DimXpert] Running auto-dimension scheme on '{configName}' " +
                    $"(datums: {(datums.UsedEdgeFallback ? "edge fallback" : "3 faces")})");

                bool created = dxPart.CreateAutoDimensionScheme(schemeOption);
                if (!created)
                {
                    result.Message = "CreateAutoDimensionScheme returned false (no features recognized or selection error)";
                    // This is not necessarily a failure — simple flat blanks may have no recognizable features
                    result.Success = true;
                    result.FeaturesRecognized = 0;
                    return result;
                }

                // Count recognized features
                var featuresRaw = dxPart.GetFeatures() as object[];
                result.FeaturesRecognized = featuresRaw?.Length ?? 0;

                // Enumerate feature types for logging
                if (featuresRaw != null)
                {
                    var typeCounts = new Dictionary<string, int>();
                    foreach (var featObj in featuresRaw)
                    {
                        var dxFeat = featObj as IDimXpertFeature;
                        if (dxFeat == null) continue;

                        string typeStr = dxFeat.GetFeatureType().ToString();
                        if (typeCounts.ContainsKey(typeStr))
                            typeCounts[typeStr]++;
                        else
                            typeCounts[typeStr] = 1;

                        // Count annotations on this feature
                        var annsRaw = dxFeat.GetAppliedAnnotations() as object[];
                        if (annsRaw != null)
                            result.AnnotationsCreated += annsRaw.Length;
                    }

                    result.RecognizedFeatureTypes = typeCounts
                        .Select(kv => $"{kv.Key}={kv.Value}")
                        .ToList();
                }

                result.Success = true;
                result.Message = $"DimXpert recognized {result.FeaturesRecognized} features, " +
                    $"{result.AnnotationsCreated} annotations";

                ErrorHandler.DebugLog($"[DimXpert] {result.Message} [{string.Join(", ", result.RecognizedFeatureTypes)}]");

                return result;
            }
            catch (Exception ex)
            {
                result.Message = "Exception: " + ex.Message;
                ErrorHandler.DebugLog($"[DimXpert] Error in RunAutoScheme: {ex.Message}");
                return result;
            }
            finally
            {
                partModel?.ClearSelection2(true);
                PerformanceTracker.Instance.StopTimer("DimXpert_RunAutoScheme");
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Imports DimXpert annotations from the 3D model into a drawing view.
        /// Uses InsertModelAnnotations3 (available SW 2022+).
        /// </summary>
        /// <returns>Number of annotations imported into the view.</returns>
        public int ImportAnnotationsToView(IDrawingDoc drawDoc, IView view)
        {
            const string proc = nameof(ImportAnnotationsToView);
            PerformanceTracker.Instance.StartTimer("DimXpert_ImportAnnotations");

            try
            {
                if (drawDoc == null || view == null) return 0;

                string viewName = view.GetName2();
                drawDoc.ActivateView(viewName);

                // Count annotations before import
                var annotsBefore = view.GetAnnotations() as object[];
                int countBefore = annotsBefore?.Length ?? 0;

                // InsertModelAnnotations3: imports model annotations into current active view
                // Parameters (booleans for what to import):
                //   source: swImportModelItemsFromEntireModel (0)
                //   importDimensions: true
                //   importNotes: false
                //   importGTols: false
                //   importDatumFeatures: false
                //   importDatums: false
                //   importSurfaceFinish: false
                //   importWeldSymbols: false
                //   importCosmeticThreads: false
                //   importOther: false
                //   eliminateDuplicates: true
                //   alignDimsTolerance: true
                bool success = drawDoc.InsertModelAnnotations3(
                    0,      // current view
                    (int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel,
                    true,   // import dimensions
                    true,   // import notes
                    false,  // import GTol
                    false,  // import datum features
                    false,  // import datums
                    false,  // import surface finish symbols
                    false,  // import weld symbols
                    false,  // import cosmetic threads
                    false,  // import other
                    true    // eliminate duplicates
                );

                // Count annotations after import
                var annotsAfter = view.GetAnnotations() as object[];
                int countAfter = annotsAfter?.Length ?? 0;
                int imported = countAfter - countBefore;

                ErrorHandler.DebugLog($"[DimXpert] ImportAnnotations to '{viewName}': {imported} new annotations " +
                    $"(before={countBefore}, after={countAfter}, success={success})");

                return imported;
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[DimXpert] ImportAnnotations error: {ex.Message}");
                return 0;
            }
            finally
            {
                PerformanceTracker.Instance.StopTimer("DimXpert_ImportAnnotations");
            }
        }
    }
}
