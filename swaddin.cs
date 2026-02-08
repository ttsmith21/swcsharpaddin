using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.Reflection;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;
using SolidWorksTools;
using SolidWorksTools.File;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using NM.Core;
using NM.Core.Config;
using NM.SwAddin;
using NM.SwAddin.UI;


namespace swcsharpaddin
{
    /// <summary>
    /// Summary description for swcsharpaddin.
    /// </summary>
    [Guid("d5355548-9569-4381-9939-5d14252a3e47"), ComVisible(true)]
    [SwAddin(
        Description = "NM AutoPilot - Automated sheet metal and tube processing",
        Title = "NM AutoPilot",
        LoadAtStartup = true
        )]
    public class SwAddin : ISwAddin
    {
        #region Local Variables
        ISldWorks iSwApp = null;
        ICommandManager iCmdMgr = null;
        int addinID = 0;
        BitmapHandler iBmp;
        IconResourceHelper _iconHelper;

        public const int mainCmdGroupID = 5;
        public const int mainItemID1 = 0;
        public const int mainItemID2 = 1;
        public const int mainItemID3 = 2;   // Smoke Tests (DEBUG)
        public const int mainItemID4 = 3;   // Run Pipeline
        public const int mainItemID5 = 4;   // QA Runner (DEBUG)
        public const int mainItemID6 = 5;   // Review Problems

        #region Event Handler Variables
        Hashtable openDocs = new Hashtable();
        SolidWorks.Interop.sldworks.SldWorks SwEventPtr = null;
        #endregion



        // Public Properties
        public ISldWorks SwApp
        {
            get { return iSwApp; }
        }
        public ICommandManager CmdMgr
        {
            get { return iCmdMgr; }
        }

        public Hashtable OpenDocs
        {
            get { return openDocs; }
        }

        #endregion

        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute
            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion
            
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }

            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
                
                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        #region ISwAddin Implementation
        public SwAddin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            // Initialize configuration from JSON files (config/nm-config.json, config/nm-tables.json)
            try
            {
                NmConfigProvider.Initialize();
                if (NmConfigProvider.LoadedConfigPath != null)
                    ErrorHandler.DebugLog($"[Config] Loaded from: {NmConfigProvider.LoadedConfigPath}");
                else
                    ErrorHandler.DebugLog("[Config] Using compiled defaults (no JSON files found)");

                foreach (var msg in NmConfigProvider.LastValidation)
                    ErrorHandler.DebugLog($"[Config] {msg}");
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[Config] Initialization failed (using defaults): {ex.Message}");
            }

            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;

            //Setup callbacks
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            #region Setup the Command Manager
            iCmdMgr = iSwApp.GetCommandManager(cookie);
            AddCommandMgr();
            #endregion

            #region Setup the Event Handlers
            SwEventPtr = (SolidWorks.Interop.sldworks.SldWorks)iSwApp;
            openDocs = new Hashtable();
            AttachEventHandlers();
            #endregion

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();
            DetachEventHandlers();

	    System.Runtime.InteropServices.Marshal.ReleaseComObject(iCmdMgr);
            iCmdMgr = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp);
            iSwApp = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers
            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }
        #endregion

        #region UI Methods
        public void AddCommandMgr()
        {
            ICommandGroup cmdGroup;
            if (iBmp == null)
                iBmp = new BitmapHandler();
            if (_iconHelper == null)
                _iconHelper = new IconResourceHelper();

            Assembly thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());

            string Title = "NM AutoPilot", ToolTip = "NM AutoPilot - Northern Manufacturing";

            int[] docTypes = new int[] {
                (int)swDocumentTypes_e.swDocASSEMBLY,
                (int)swDocumentTypes_e.swDocDRAWING,
                (int)swDocumentTypes_e.swDocPART
            };

            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            // All command IDs that will be registered - must match exactly
#if DEBUG
            int[] knownIDs = new int[] { mainItemID4, mainItemID6, mainItemID5, mainItemID3 };
#else
            int[] knownIDs = new int[] { mainItemID4, mainItemID6 };
#endif

            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs))
                {
                    ignorePrevious = true;
                }
            }

            cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);

            // Modern PNG icon API (SolidWorks 2022+)
            try
            {
                cmdGroup.IconList = _iconHelper.GetToolbarIconList(thisAssembly);
                cmdGroup.MainIconList = _iconHelper.GetMainIconList(thisAssembly);
            }
            catch (Exception ex)
            {
                // Fallback to legacy BMP if PNG extraction fails
                NM.Core.ErrorHandler.DebugLog("[Icons] PNG icon load failed, falling back to BMP: " + ex.Message);
                cmdGroup.LargeIconList = iBmp.CreateFileFromResourceBitmap("swcsharpaddin.ToolbarLarge.bmp", thisAssembly);
                cmdGroup.SmallIconList = iBmp.CreateFileFromResourceBitmap("swcsharpaddin.ToolbarSmall.bmp", thisAssembly);
                cmdGroup.LargeMainIcon = iBmp.CreateFileFromResourceBitmap("swcsharpaddin.MainIconLarge.bmp", thisAssembly);
                cmdGroup.SmallMainIcon = iBmp.CreateFileFromResourceBitmap("swcsharpaddin.MainIconSmall.bmp", thisAssembly);
            }

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);

            // Icon strip index: 0=RunPipeline, 1=ReviewProblems, 2=RunQA, 3=RunSmokeTests
            int cmdIndexPipeline = cmdGroup.AddCommandItem2("Run Pipeline", -1,
                "Run unified processing pipeline on active document", "Run Pipeline",
                0, "RunPipeline", "", mainItemID4, menuToolbarOption);

            int cmdIndexReview = cmdGroup.AddCommandItem2("Review Problems", -1,
                "Review and fix problem parts from the last pipeline run", "Review Problems",
                1, "ReviewProblems", "ReviewProblemsEnable", mainItemID6, menuToolbarOption);

#if DEBUG
            int cmdIndexQA = cmdGroup.AddCommandItem2("Run QA", -1,
                "Run Gold Standard QA tests", "Run QA",
                2, "RunQA", "", mainItemID5, menuToolbarOption);

            int cmdIndexSmoke = cmdGroup.AddCommandItem2("Run Smoke Tests", -1,
                "Run automated smoke tests", "Run Tests",
                3, "RunSmokeTests", "", mainItemID3, menuToolbarOption);
#endif

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();

            // Set up the command tab for each document type
            foreach (int type in docTypes)
            {
                CommandTab cmdTab = iCmdMgr.GetCommandTab(type, Title);

                if (cmdTab != null && (!getDataResult || ignorePrevious))
                {
                    iCmdMgr.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                if (cmdTab == null)
                {
                    cmdTab = iCmdMgr.AddCommandTab(type, Title);
                    CommandTabBox cmdBox = cmdTab.AddCommandTabBox();

#if DEBUG
                    int[] cmdIDs = new int[4];
                    int[] TextType = new int[4];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndexPipeline);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndexReview);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndexQA);
                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdIDs[3] = cmdGroup.get_CommandID(cmdIndexSmoke);
                    TextType[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
#else
                    int[] cmdIDs = new int[2];
                    int[] TextType = new int[2];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndexPipeline);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndexReview);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
#endif

                    cmdBox.AddCommands(cmdIDs, TextType);
                }
            }

            thisAssembly = null;
        }

        public void RemoveCommandMgr()
        {
            iBmp.Dispose();
            _iconHelper?.Dispose();
            iCmdMgr.RemoveCommandGroup(mainCmdGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }
            else
            {

                for (int i = 0; i < addinList.Count; i++)
                {
                    if (addinList[i] != storedList[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        #endregion

        #region UI Callbacks
        #endregion

        #region Event Methods
        public bool AttachEventHandlers()
        {
            AttachSwEvents();
            //Listen for events on all currently open docs
            AttachEventsToAllDocuments();
            return true;
        }

        private bool AttachSwEvents()
        {
            try
            {
                SwEventPtr.FileNewNotify2 += new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.FileOpenPostNotify += new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }



        private bool DetachSwEvents()
        {
            try
            {
                SwEventPtr.FileNewNotify2 -= new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.FileOpenPostNotify -= new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }

        }

        public void AttachEventsToAllDocuments()
        {
            ModelDoc2 modDoc = (ModelDoc2)iSwApp.GetFirstDocument();
            while (modDoc != null)
            {
                if (!openDocs.Contains(modDoc))
                {
                    AttachModelDocEventHandler(modDoc);
                }
                modDoc = (ModelDoc2)modDoc.GetNext();
            }
        }

        public bool AttachModelDocEventHandler(ModelDoc2 modDoc)
        {
            if (modDoc == null)
                return false;

            DocumentEventHandler docHandler = null;

            if (!openDocs.Contains(modDoc))
            {
                switch (modDoc.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        {
                            docHandler = new PartEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        {
                            docHandler = new AssemblyEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        {
                            docHandler = new DrawingEventHandler(modDoc, this);
                            break;
                        }
                    default:
                        {
                            return false; //Unsupported document type
                        }
                }
                docHandler.AttachEventHandlers();
                openDocs.Add(modDoc, docHandler);
            }
            return true;
        }

        public bool DetachModelEventHandler(ModelDoc2 modDoc)
        {
            DocumentEventHandler docHandler;
            docHandler = (DocumentEventHandler)openDocs[modDoc];
            openDocs.Remove(modDoc);
            modDoc = null;
            docHandler = null;
            return true;
        }

        public bool DetachEventHandlers()
        {
            DetachSwEvents();

            //Close events on all currently open docs
            DocumentEventHandler docHandler;
            int numKeys = openDocs.Count;
            object[] keys = new Object[numKeys];

            //Remove all document event handlers
            openDocs.Keys.CopyTo(keys, 0);
            foreach (ModelDoc2 key in keys)
            {
                docHandler = (DocumentEventHandler)openDocs[key];
                docHandler.DetachEventHandlers(); //This also removes the pair from the hash
                docHandler = null;
            }
            return true;
        }
        #endregion

        #region Event Handlers
        int FileOpenPostNotify(string FileName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnFileNew(object newDoc, int docType, string templateName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        #endregion

        /// <summary>
        /// Runs the unified two-pass workflow: validate all → show problems → process good.
        /// Shows the settings form first to let user choose material, bend options, and output settings.
        /// </summary>
        public void RunPipeline()
        {
            NM.Core.ErrorHandler.PushCallStack("RunPipeline");
            try
            {
                // Show settings form first (like the VBA UIX)
                NM.Core.ProcessingOptions options;
                using (var settingsForm = new MainSelectionForm(iSwApp))
                {
                    if (settingsForm.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    {
                        NM.Core.ErrorHandler.DebugLog("[RunPipeline] User canceled settings dialog");
                        return;
                    }
                    options = settingsForm.Options;
                    NM.Core.ErrorHandler.DebugLog($"[RunPipeline] Options: Material={options.Material}, BendTable={options.BendTable}, KFactor={options.KFactor}");
                }

                // Run workflow with user-selected options
                var dispatcher = new NM.SwAddin.Pipeline.WorkflowDispatcher(iSwApp);
                dispatcher.Run(options);
            }
            catch (System.Exception ex)
            {
                NM.Core.ErrorHandler.HandleError("RunPipeline", ex.Message, ex, NM.Core.ErrorHandler.LogLevel.Error);
                System.Windows.Forms.MessageBox.Show($"Pipeline error: {ex.Message}", "Error");
            }
            finally
            {
                NM.Core.ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Opens the Review Problems UI to review and fix problem parts from the last pipeline run.
        /// Called by SolidWorks when user clicks "Review Problems" button.
        /// </summary>
        public void ReviewProblems()
        {
            NM.Core.ErrorHandler.PushCallStack("ReviewProblems");
            try
            {
                var problems = NM.Core.ProblemParts.ProblemPartManager.Instance.GetProblemParts();

                if (problems.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "No problems to review.\n\nRun the pipeline first, and any problem parts will appear here for review.",
                        "NM AutoPilot - Review Problems",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // Open the wizard form (non-modal so user can interact with SolidWorks)
                var wizard = new ProblemWizardForm(problems, iSwApp, 0);
                wizard.Show();

                while (wizard.Visible)
                {
                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(50);
                }

                int fixedCount = wizard.FixedProblems.Count;
                NM.Core.ErrorHandler.DebugLog($"[ReviewProblems] Wizard closed. Fixed: {fixedCount}, Action: {wizard.SelectedAction}");

                if (fixedCount > 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        $"Fixed {fixedCount} problem part(s).\n\n" +
                        $"Remaining problems: {NM.Core.ProblemParts.ProblemPartManager.Instance.GetProblemParts().Count}",
                        "Review Complete",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }

                wizard.Dispose();
            }
            catch (System.Exception ex)
            {
                NM.Core.ErrorHandler.HandleError("ReviewProblems", ex.Message, ex, NM.Core.ErrorHandler.LogLevel.Error);
                System.Windows.Forms.MessageBox.Show($"Review Problems error: {ex.Message}", "Error");
            }
            finally
            {
                NM.Core.ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Enable callback for Review Problems button.
        /// SolidWorks calls this to determine if the button should be grayed out.
        /// </summary>
        /// <returns>1 = enabled, 0 = disabled (grayed out)</returns>
        public int ReviewProblemsEnable()
        {
            try
            {
                bool hasProblems = NM.Core.ProblemParts.ProblemPartManager.Instance
                    .GetProblemParts().Count > 0;
                return hasProblems ? 1 : 0;
            }
            catch
            {
                return 1; // On error, keep enabled so user can click and see a message
            }
        }

        /// <summary>
        /// Runs the Gold Standard QA tests by processing files from C:\Temp\nm_qa_config.json
        /// </summary>
        public void RunQA()
        {
            NM.Core.ErrorHandler.PushCallStack("RunQA");
            try
            {
                var runner = new NM.SwAddin.Pipeline.QARunner(iSwApp);
                var summary = runner.Run();

                var msg = $"QA Complete:\n" +
                          $"  Total: {summary.TotalFiles}\n" +
                          $"  Passed: {summary.Passed}\n" +
                          $"  Failed: {summary.Failed}\n" +
                          $"  Errors: {summary.Errors}\n" +
                          $"  Time: {summary.TotalElapsedMs:F0}ms";

                System.Windows.Forms.MessageBox.Show(msg, "QA Results",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    summary.Errors > 0 || summary.Failed > 0
                        ? System.Windows.Forms.MessageBoxIcon.Warning
                        : System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                NM.Core.ErrorHandler.HandleError("RunQA", ex.Message, ex, NM.Core.ErrorHandler.LogLevel.Error);
                System.Windows.Forms.MessageBox.Show($"QA error: {ex.Message}", "Error");
            }
            finally
            {
                NM.Core.ErrorHandler.PopCallStack();
            }
        }

#if DEBUG
        public void RunSmokeTests()
        {
            NM.Core.ErrorHandler.PushCallStack("RunSmokeTests");
            var sb = new System.Text.StringBuilder();
            try
            {
                // Perf: total smoke timer
                NM.Core.PerformanceTracker.Instance.StartTimer("SmokeTests");

                if (iSwApp == null)
                {
                    System.Windows.Forms.MessageBox.Show("SolidWorks is not available.");
                    return;
                }

                var doc = iSwApp.ActiveDoc as IModelDoc2;
                if (doc == null)
                {
                    sb.AppendLine("No active document. Open a part and retry.");
                }
                else
                {
                    // Zoom to fit
                    try
                    {
                        SolidWorksApiWrapper.ZoomToFit(doc);
                        sb.AppendLine("ZoomToFit: OK");
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine("ZoomToFit: FAIL - " + ex.Message);
                    }

                    // Set view display mode
                    try
                    {
                        SolidWorksApiWrapper.SetViewDisplayMode(doc, SwViewDisplayMode.SwViewDisplayMode_Shaded);
                        sb.AppendLine("SetViewDisplayMode: OK");
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine("SetViewDisplayMode: FAIL - " + ex.Message);
                    }

                    // Material set/get (best-effort; skip if not a part)
                    try
                    {
                        bool matSet = SolidWorksApiWrapper.SetMaterialName(doc, "Plain Carbon Steel");
                        string matName = SolidWorksApiWrapper.GetMaterialName(doc);
                        sb.AppendLine($"Material: set={(matSet ? "OK" : "FAIL")}, get='{matName}'");
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine("Material: FAIL - " + ex.Message);
                    }

                    // Custom property add/get/delete
                    try
                    {
                        var added = SwPropertyHelper.AddCustomProperty(doc, "TestProp", swCustomInfoType_e.swCustomInfoText, "Hello", "");
                        SwPropertyHelper.GetCustomProperties(doc, "", out var names, out var types, out var values);
                        var idx = Array.FindIndex(names, n => string.Equals(n, "TestProp", StringComparison.OrdinalIgnoreCase));
                        bool found = idx >= 0;
                        var deleted = SwPropertyHelper.DeleteCustomProperty(doc, "TestProp", "");
                        sb.AppendLine($"CustomProps: add={(added ? "OK" : "FAIL")}, found={(found ? "YES" : "NO")}, delete={(deleted ? "OK" : "FAIL")}");
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine("CustomProps: FAIL - " + ex.Message);
                    }

                    // Rebuild
                    try
                    {
                        bool rebuilt = SwDocumentHelper.ForceRebuildDoc(doc);
                        sb.AppendLine("Rebuild: " + (rebuilt ? "OK" : "SKIPPED/NO-OP"));
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine("Rebuild: FAIL - " + ex.Message);
                    }

                    // Sketch: On Top Plane, draw a short line, then exit
                    try
                    {
                        if (SwSketchHelper.StartSketchOnPlane(doc, "Top Plane"))
                        {
                            var seg = SwSketchHelper.CreateSketchLine(doc, 0, 0, 0.02, 0.02);
                            bool ended = SwSketchHelper.EndSketch(doc);
                            sb.AppendLine($"Sketch: line={(seg != null ? "OK" : "FAIL")}, end={(ended ? "OK" : "FAIL")}");
                        }
                        else
                        {
                            sb.AppendLine("Sketch: could not start on Top Plane");
                        }
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine("Sketch: FAIL - " + ex.Message);
                    }

                    // ModelInfo sync using SolidWorksModel
                    try
                    {
                        NM.Core.PerformanceTracker.Instance.StartTimer("ModelSync");
                        var info = new NM.Core.ModelInfo();
                        var svc = new NM.SwAddin.SolidWorksModel(info, iSwApp);
                        svc.Attach(doc);
                        bool loaded = svc.LoadPropertiesFromSolidWorks();
                        // For visibility in UI, write at document-level (Custom tab), not config-specific
                        info.ConfigurationName = string.Empty;
                        info.CustomProperties.SetPropertyValue("SyncTest", "42", NM.Core.CustomPropertyType.Text);
                        bool syncOk = svc.SavePropertiesToSolidWorks();
                        sb.AppendLine("ModelSync: " + (syncOk ? "OK" : "FAIL") + " (target=Custom tab)");
                        if (!syncOk)
                        {
                            sb.AppendLine("Info: " + info.ProblemDescription);
                        }

                        // DEBUG: SolidWorksFileOperations smoke - SaveAs/Activate/Close
                        try
                        {
                            NM.Core.PerformanceTracker.Instance.StartTimer("FileOps");
                            var fileSvc = new NM.SwAddin.SolidWorksFileOperations(iSwApp);
                            var title = doc.GetTitle();
                            var tempDir = Path.GetTempPath();
                            var newPath = Path.Combine(tempDir, $"{title}_SmokeCopy.sldprt");
                            // If drawing/assembly, extension will differ; derive from existing path if available
                            var ext = Path.GetExtension(doc.GetPathName());
                            if (!string.IsNullOrEmpty(ext)) newPath = Path.ChangeExtension(newPath, ext);

                            if (fileSvc.SaveAs(doc, newPath))
                            {
                                sb.AppendLine($"FileOps: SaveAs OK -> {newPath}");
                                var reopened = fileSvc.OpenSWDocument(newPath, silent: true, readOnly: true);
                                sb.AppendLine("FileOps: Open RO " + (reopened != null ? "OK" : "FAIL"));
                                if (reopened != null)
                                {
                                    // Activate by title
                                    int errs = 0; iSwApp.ActivateDoc3(Path.GetFileName(newPath), true, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref errs);
                                    sb.AppendLine("FileOps: Activate " + (errs == 0 ? "OK" : $"ERR={errs}"));
                                    fileSvc.CloseSWDocument(reopened);
                                    sb.AppendLine("FileOps: Close OK");
                                }
                            }
                            else
                            {
                                sb.AppendLine("FileOps: SaveAs FAIL");
                            }
                        }
                        catch (Exception fx)
                        {
                            sb.AppendLine("FileOps: EX - " + fx.Message);
                        }
                        finally
                        {
                            NM.Core.PerformanceTracker.Instance.StopTimer("FileOps");
                        }
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine("ModelSync: FAIL - " + ex.Message);
                    }
                    finally
                    {
                        NM.Core.PerformanceTracker.Instance.StopTimer("ModelSync");
                    }
                }

                // Perf: finalize and export CSV
                NM.Core.PerformanceTracker.Instance.StopTimer("SmokeTests");
                var csv = Path.Combine(@"C:\SolidWorksMacroLogs", "perf.csv");
                NM.Core.PerformanceTracker.Instance.ExportToCsv(csv);
                var count = NM.Core.PerformanceTracker.Instance.GetTimerCount();
                sb.AppendLine($"Perf: timers={count}, csv={csv}");

                System.Windows.Forms.MessageBox.Show(sb.ToString(), "Smoke Tests");
            }
            finally
            {
                NM.Core.ErrorHandler.PopCallStack();
            }
        }
#endif

    }

}

namespace NM.Core
{
    /// <summary>
    /// Configuration facade — backward-compatible static API that delegates to NmConfigProvider.
    /// Existing code can continue using Configuration.Manufacturing.F115CostPerHour etc.
    /// New code should prefer NmConfigProvider.Current directly.
    /// </summary>
    public static class Configuration
    {
        /// <summary>Logging and debugging settings.</summary>
        public static class Logging
        {
            public static bool LogEnabled => NM.Core.Config.NmConfigProvider.Current.Logging.LogEnabled;
            public static string ErrorLogPath => NM.Core.Config.NmConfigProvider.Current.Paths.ErrorLogPath;
            public static bool ShowWarnings => NM.Core.Config.NmConfigProvider.Current.Logging.ShowWarnings;
            public static bool ProductionMode => NM.Core.Config.NmConfigProvider.Current.Logging.ProductionMode;
            public static bool EnablePerformanceMonitoring => NM.Core.Config.NmConfigProvider.Current.Logging.PerformanceMonitoring;
            public static bool EnableDebugMode => NM.Core.Config.NmConfigProvider.Current.Logging.DebugMode;
            public static bool IsProductionMode => ProductionMode;
            public static bool IsPerformanceMonitoringEnabled => EnablePerformanceMonitoring;
        }

        /// <summary>File system locations used by the application.</summary>
        public static class FilePaths
        {
            private static NM.Core.Config.Sections.PathConfig P => NM.Core.Config.NmConfigProvider.Current.Paths;
            public static string MaterialFilePath => P.MaterialDataPaths != null && P.MaterialDataPaths.Length > 0 ? P.MaterialDataPaths[0] : "";
            public static string LaserDataFilePath => MaterialFilePath;
            public static string MaterialPropertyFilePath => P.MaterialPropertyFilePath;
            public static string BendTableSs => P.BendTables.StainlessSteel != null && P.BendTables.StainlessSteel.Length > 0 ? P.BendTables.StainlessSteel[0] : "";
            public static string BendTableCs => P.BendTables.CarbonSteel != null && P.BendTables.CarbonSteel.Length > 0 ? P.BendTables.CarbonSteel[0] : "";
            public static string BendTableSsLocal => P.BendTables.StainlessSteel != null && P.BendTables.StainlessSteel.Length > 1 ? P.BendTables.StainlessSteel[1] : "";
            public static string BendTableCsLocal => P.BendTables.CarbonSteel != null && P.BendTables.CarbonSteel.Length > 1 ? P.BendTables.CarbonSteel[1] : "";
            public static string BendTableNone => P.BendTables.NoneValue;
            public const string ExcelLookupFile = "NewLaser.xls";
            public static string ExtractDataAddInPath => P.ExtractDataAddInPath;
            public static string ErrorLogPath => P.ErrorLogPath;
        }

        /// <summary>Manufacturing rates, costs, and processing parameters.</summary>
        public static class Manufacturing
        {
            private static NM.Core.Config.NmConfig C => NM.Core.Config.NmConfigProvider.Current;
            private static NM.Core.Config.Sections.ManufacturingParams M => C.Manufacturing;

            // Press brake rate tiers
            public static double Rate1Seconds => M.PressBrake.RateSeconds[0];
            public static double Rate2Seconds => M.PressBrake.RateSeconds[1];
            public static double Rate3Seconds => M.PressBrake.RateSeconds[2];
            public static double Rate4Seconds => M.PressBrake.RateSeconds[3];
            public static double Rate5Seconds => M.PressBrake.RateSeconds[4];
            public static double SetupRateMinutesPerFoot => M.PressBrake.SetupMinutesPerFoot;
            public static double BrakeSetupMinutes => M.PressBrake.SetupFixedMinutes;
            public static double Rate3MaxWeightLbs => M.PressBrake.WeightThresholdsLbs[2];
            public static double Rate2MaxWeightLbs => M.PressBrake.WeightThresholdsLbs[1];
            public static double Rate1MaxWeightLbs => M.PressBrake.WeightThresholdsLbs[0];
            public static double Rate1MaxLengthIn => M.PressBrake.LengthThresholdsIn[0];
            public static double Rate2MaxLengthIn => M.PressBrake.LengthThresholdsIn[1];

            // Laser/waterjet setup
            public static double LaserSetupRateMinutesPerSheet => M.Laser.SetupMinutesPerSheet;
            public static double LaserSetupFixedMinutes => M.Laser.SetupFixedMinutes;
            public static double WaterJetSetupFixedMinutes => M.Waterjet.SetupFixedMinutes;
            public static double WaterJetSetupRateMinutesPerLoad => M.Waterjet.SetupMinutesPerLoad;
            public static double StandardSheetWidthIn => M.StandardSheet.WidthIn;
            public static double StandardSheetLengthIn => M.StandardSheet.LengthIn;

            // Work center hourly costs
            public static double F115CostPerHour => C.WorkCenters.F115_LaserCutting;
            public static double F300CostPerHour => C.WorkCenters.F300_MaterialHandling;
            public static double F210CostPerHour => C.WorkCenters.F210_Deburring;
            public static double F140CostPerHour => C.WorkCenters.F140_PressBrake;
            public static double F145CostPerHour => C.WorkCenters.F145_CncBending;
            public static double F155CostPerHour => C.WorkCenters.F155_Waterjet;
            public static double F220CostPerHour => C.WorkCenters.F220_Tapping;
            public static double F325CostPerHour => C.WorkCenters.F325_RollForming;
            public static double F400CostPerHour => C.WorkCenters.F400_Welding;
            public static double F385CostPerHour => C.WorkCenters.F385_Assembly;
            public static double F500CostPerHour => C.WorkCenters.F500_Finishing;
            public static double F525CostPerHour => C.WorkCenters.F525_Packaging;
            public static double EngCostPerHour => C.WorkCenters.ENG_Engineering;

            // Pricing modifiers
            public static double MaterialMarkup => C.Pricing.MaterialMarkup;
            public static double TightPercent => C.Pricing.TightTolerance;
            public static double NormalPercent => C.Pricing.NormalTolerance;
            public static double LoosePercent => C.Pricing.LooseTolerance;

            // Cutting
            public static double PierceConstant => M.Cutting.PierceConstant;
            public static int TabSpacing => M.Cutting.TabSpacing;
        }

        /// <summary>Material-related constants (unit conversions delegate to UnitConversions; properties delegate to config).</summary>
        public static class Materials
        {
            // Unit conversions — true physics constants
            public const double MetersToInches = NM.Core.Constants.UnitConversions.MetersToInches;
            public const double InchesToMeters = NM.Core.Constants.UnitConversions.InchesToMeters;
            public const double KgToLbs = NM.Core.Constants.UnitConversions.KgToLbs;
            public static double SteelDensityLbsPerIn3 => NM.Core.Config.NmConfigProvider.Current.MaterialDensities.Steel_General;

            public static System.Collections.Generic.IReadOnlyList<string> InitialCustomProperties
                => NM.Core.Config.NmConfigProvider.Current.CustomProperties.InitialProperties.AsReadOnly();

            public static System.Collections.Generic.IReadOnlyList<string> GetInitialCustomProperties() => InitialCustomProperties;
        }

        /// <summary>General application defaults and user preferences.</summary>
        public static class Defaults
        {
            private static NM.Core.Config.Sections.ProcessingDefaults D => NM.Core.Config.NmConfigProvider.Current.Defaults;
            private static NM.Core.Config.Sections.LoggingConfig L => NM.Core.Config.NmConfigProvider.Current.Logging;

            public static int MaxRetries => D.MaxRetries;
            public static string DefaultSheetName => D.DefaultSheetName;
            public static bool AutoCloseExcel => D.AutoCloseExcel;
            public static double DefaultCostPerLb => NM.Core.Config.NmConfigProvider.Current.MaterialPricing.DefaultCostPerLb;
            public static int DefaultQuantity => D.DefaultQuantity;
            public static double DefaultKFactor => D.DefaultKFactor;
            public static bool LogEnabledDefault => L.LogEnabled;
            public static bool ShowWarningsDefault => L.ShowWarnings;
            public static bool ProductionModeDefault => L.ProductionMode;
            public static bool EnableDebugModeDefault => L.DebugMode;
            public static bool EnablePerformanceMonitoringDefault => L.PerformanceMonitoring;
            public static bool SolidWorksVisibleDefault => L.SolidWorksVisible;
        }
    }
}
