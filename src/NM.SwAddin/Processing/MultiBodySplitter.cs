using System;
using System.IO;
using System.Runtime.InteropServices;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Processing
{
    /// <summary>
    /// Splits a multi-body part into individual part files and a containing assembly
    /// using SolidWorks Save Bodies feature. Each body becomes PartName-A, PartName-B, etc.
    /// The generated assembly preserves spatial position of all bodies.
    /// </summary>
    public sealed class MultiBodySplitter
    {
        private readonly ISldWorks _swApp;

        public MultiBodySplitter(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        public sealed class SplitResult
        {
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }

            /// <summary>Path to the generated assembly containing all sub-parts.</summary>
            public string AssemblyPath { get; set; }

            /// <summary>Paths to the individual part files (one per body).</summary>
            public string[] PartPaths { get; set; } = Array.Empty<string>();

            /// <summary>Number of solid bodies that were split.</summary>
            public int BodyCount { get; set; }

            public static SplitResult Fail(string message)
            {
                return new SplitResult { Success = false, ErrorMessage = message };
            }
        }

        /// <summary>
        /// Splits a multi-body part into individual part files and an assembly.
        /// Uses SolidWorks CreateSaveBodyFeature (Insert > Features > Save Bodies).
        /// </summary>
        /// <param name="doc">The open multi-body part document.</param>
        /// <returns>Result with paths to generated files.</returns>
        public SplitResult SplitToAssembly(IModelDoc2 doc)
        {
            const string proc = "MultiBodySplitter.SplitToAssembly";
            ErrorHandler.PushCallStack(proc);
            PerformanceTracker.Instance.StartTimer("MultiBodySplit");
            try
            {
                if (doc == null)
                    return SplitResult.Fail("Document is null");

                var partDoc = doc as IPartDoc;
                if (partDoc == null)
                    return SplitResult.Fail("Document is not a part");

                // Get all solid bodies
                var bodiesObj = partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
                if (bodiesObj == null)
                    return SplitResult.Fail("No solid bodies found");

                var bodies = (object[])bodiesObj;
                if (bodies.Length < 2)
                    return SplitResult.Fail($"Part has {bodies.Length} body - not a multi-body part");

                // Derive output file paths from the source part
                string sourcePath = doc.GetPathName();
                if (string.IsNullOrWhiteSpace(sourcePath))
                    return SplitResult.Fail("Part must be saved to disk before splitting");

                string dir = Path.GetDirectoryName(sourcePath);
                string baseName = Path.GetFileNameWithoutExtension(sourcePath);

                // Generate suffixed names: PartName-A.SLDPRT, PartName-B.SLDPRT, ...
                string[] partPaths = new string[bodies.Length];
                for (int i = 0; i < bodies.Length; i++)
                {
                    string suffix = GetSuffix(i);
                    partPaths[i] = Path.Combine(dir, $"{baseName}-{suffix}.SLDPRT");
                }

                string assyPath = Path.Combine(dir, $"{baseName}-ASSY.SLDASM");

                ErrorHandler.DebugLog($"[SPLIT] Splitting {baseName} into {bodies.Length} bodies");
                for (int i = 0; i < partPaths.Length; i++)
                    ErrorHandler.DebugLog($"[SPLIT]   Body {i} -> {Path.GetFileName(partPaths[i])}");
                ErrorHandler.DebugLog($"[SPLIT]   Assembly -> {Path.GetFileName(assyPath)}");

                // CRITICAL: Wrap each IBody2 in DispatchWrapper for COM marshaling.
                // Passing raw object[] to CreateSaveBodyFeature causes AccessViolationException in C#.
                var wrappedBodies = new DispatchWrapper[bodies.Length];
                for (int i = 0; i < bodies.Length; i++)
                {
                    wrappedBodies[i] = new DispatchWrapper(bodies[i]);
                }

                // Call CreateSaveBodyFeature - the API equivalent of Insert > Features > Save Bodies
                var featMgr = doc.FeatureManager;
                IFeature saveFeat = (IFeature)featMgr.CreateSaveBodyFeature(
                    wrappedBodies,    // Bodies to save (DispatchWrapper[])
                    partPaths,        // Output file paths
                    assyPath,         // Assembly output path
                    true,             // Create assembly
                    false             // Don't consume original bodies
                );

                if (saveFeat == null)
                {
                    ErrorHandler.DebugLog("[SPLIT] CreateSaveBodyFeature returned null - trying fallback");
                    return TryManualSplit(doc, partDoc, bodies, partPaths, assyPath);
                }

                ErrorHandler.DebugLog($"[SPLIT] CreateSaveBodyFeature succeeded: {saveFeat.Name}");

                // Save the source part (it now has the Save Bodies feature)
                int saveErrors = 0, saveWarnings = 0;
                doc.Save3(
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    ref saveErrors, ref saveWarnings);

                // Verify output files exist
                int verified = 0;
                foreach (var p in partPaths)
                {
                    if (File.Exists(p)) verified++;
                    else ErrorHandler.DebugLog($"[SPLIT] WARNING: Expected file not found: {p}");
                }

                if (verified == 0)
                    return SplitResult.Fail("Save Bodies feature created but no output files found on disk");

                return new SplitResult
                {
                    Success = true,
                    AssemblyPath = assyPath,
                    PartPaths = partPaths,
                    BodyCount = bodies.Length
                };
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Split failed", ex, ErrorHandler.LogLevel.Error);
                return SplitResult.Fail($"Exception: {ex.Message}");
            }
            finally
            {
                PerformanceTracker.Instance.StopTimer("MultiBodySplit");
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Fallback: manually copy each body into a new part file and create an assembly.
        /// Used when CreateSaveBodyFeature fails (e.g., imported geometry issues).
        /// </summary>
        private SplitResult TryManualSplit(IModelDoc2 doc, IPartDoc partDoc, object[] bodies, string[] partPaths, string assyPath)
        {
            ErrorHandler.DebugLog("[SPLIT] Attempting manual body copy fallback");

            string templatePart = _swApp.GetUserPreferenceStringValue(
                (int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            string templateAssy = _swApp.GetUserPreferenceStringValue(
                (int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);

            if (string.IsNullOrEmpty(templatePart) || string.IsNullOrEmpty(templateAssy))
                return SplitResult.Fail("Cannot find SolidWorks part/assembly templates for manual split");

            // Copy each body into a new part file
            for (int i = 0; i < bodies.Length; i++)
            {
                var srcBody = (IBody2)bodies[i];
                try
                {
                    // Copy body geometry into a temporary body
                    var copiedBody = (IBody2)srcBody.Copy();
                    if (copiedBody == null)
                    {
                        ErrorHandler.DebugLog($"[SPLIT] Body {i} copy returned null");
                        return SplitResult.Fail($"Failed to copy body {i}");
                    }

                    // Create new part from template
                    var newDoc = (IModelDoc2)_swApp.NewDocument(templatePart, 0, 0, 0);
                    if (newDoc == null)
                        return SplitResult.Fail($"Failed to create new part document for body {i}");

                    var newPart = (IPartDoc)newDoc;

                    // Insert body as a feature in the new part
                    // swCreateFeatureBodyOpts_e: 0=Default, 1=Check, 4=Simplify
                    var feat = newPart.CreateFeatureFromBody3(copiedBody, false, 0);

                    if (feat == null)
                        ErrorHandler.DebugLog($"[SPLIT] CreateFeatureFromBody3 returned null for body {i} (may still work)");

                    // Copy material from source part
                    var srcPart = (IPartDoc)doc;
                    string srcConfig = doc.ConfigurationManager?.ActiveConfiguration?.Name ?? "";
                    string matDb;
                    string matName = srcPart.GetMaterialPropertyName2(srcConfig, out matDb);
                    if (!string.IsNullOrEmpty(matName))
                    {
                        newPart.SetMaterialPropertyName2("", matDb ?? "", matName);
                    }

                    // Save the new part
                    int errors = 0, warnings = 0;
                    newDoc.Extension.SaveAs(
                        partPaths[i],
                        (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                        null, ref errors, ref warnings);

                    _swApp.CloseDoc(newDoc.GetTitle());
                }
                catch (Exception ex)
                {
                    ErrorHandler.DebugLog($"[SPLIT] Manual copy failed for body {i}: {ex.Message}");
                    return SplitResult.Fail($"Manual copy failed for body {i}: {ex.Message}");
                }
            }

            // Create assembly with all sub-parts at origin (bodies already carry their coordinates)
            try
            {
                var assyDoc = (IModelDoc2)_swApp.NewDocument(templateAssy, 0, 0, 0);
                if (assyDoc == null)
                    return SplitResult.Fail("Failed to create assembly document");

                var assy = (IAssemblyDoc)assyDoc;

                // Add each sub-part at identity transform (origin)
                // Bodies were copied with their original coordinates intact,
                // so placing each component at origin preserves spatial relationships.
                foreach (var partPath in partPaths)
                {
                    // swAddComponentConfigOptions_e: 0=CurrentSelectedConfig
                    var comp = assy.AddComponent5(partPath, 0, "", false, "", 0, 0, 0);

                    if (comp == null)
                        ErrorHandler.DebugLog($"[SPLIT] AddComponent5 returned null for {Path.GetFileName(partPath)}");
                }

                // Save assembly
                int aErrors = 0, aWarnings = 0;
                assyDoc.Extension.SaveAs(
                    assyPath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    null, ref aErrors, ref aWarnings);

                _swApp.CloseDoc(assyDoc.GetTitle());
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[SPLIT] Assembly creation failed: {ex.Message}");
                return SplitResult.Fail($"Assembly creation failed: {ex.Message}");
            }

            return new SplitResult
            {
                Success = true,
                AssemblyPath = assyPath,
                PartPaths = partPaths,
                BodyCount = bodies.Length
            };
        }

        /// <summary>
        /// Generates a letter suffix: 0->A, 1->B, ..., 25->Z, 26->AA, 27->AB, etc.
        /// </summary>
        private static string GetSuffix(int index)
        {
            if (index < 26)
                return ((char)('A' + index)).ToString();

            // For >26 bodies: AA, AB, AC, ...
            int high = (index / 26) - 1;
            int low = index % 26;
            return ((char)('A' + high)).ToString() + ((char)('A' + low)).ToString();
        }
    }
}
