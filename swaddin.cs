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
using NM.SwAddin;


namespace swcsharpaddin
{
    /// <summary>
    /// Summary description for swcsharpaddin.
    /// </summary>
    [Guid("d5355548-9569-4381-9939-5d14252a3e47"), ComVisible(true)]
    [SwAddin(
        Description = "swcsharpaddin description",
        Title = "swcsharpaddin",
        LoadAtStartup = true
        )]
    public class SwAddin : ISwAddin
    {
        #region Local Variables
        ISldWorks iSwApp = null;
        ICommandManager iCmdMgr = null;
        int addinID = 0;
        BitmapHandler iBmp;

        public const int mainCmdGroupID = 5;
        public const int mainItemID1 = 0;
        public const int mainItemID2 = 1;
        public const int mainItemID3 = 2;
        public const int mainItemID4 = 3;

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

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }
        #endregion

        #region UI Methods
        public void AddCommandMgr()
        {
            ICommandGroup cmdGroup;
            if(iBmp == null)
                iBmp = new BitmapHandler();
            Assembly thisAssembly;
            int cmdIndex2 = -1;
            string Title = "C# Addin", ToolTip = "C# Addin";


            int[] docTypes = new int[]{(int)swDocumentTypes_e.swDocASSEMBLY,
                                       (int)swDocumentTypes_e.swDocDRAWING,
                                       (int)swDocumentTypes_e.swDocPART};

            thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());


            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;
            //get the ID information stored in the registry
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            int[] knownIDs = new int[2] { mainItemID1, mainItemID2};
            
            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs)) //if the IDs don't match, reset the commandGroup
            {
                    ignorePrevious = true;
                }
            }

            cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);
            cmdGroup.LargeIconList = iBmp.CreateFileFromResourceBitmap("swcsharpaddin.ToolbarLarge.bmp", thisAssembly);
            cmdGroup.SmallIconList = iBmp.CreateFileFromResourceBitmap("swcsharpaddin.ToolbarSmall.bmp", thisAssembly);
            cmdGroup.LargeMainIcon = iBmp.CreateFileFromResourceBitmap("swcsharpaddin.MainIconLarge.bmp", thisAssembly);
            cmdGroup.SmallMainIcon = iBmp.CreateFileFromResourceBitmap("swcsharpaddin.MainIconSmall.bmp", thisAssembly);

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            cmdGroup.AddCommandItem2("Run Pipeline", -1, "Run unified processing pipeline", "Run Pipeline", 1, "RunPipeline", "", mainItemID4, menuToolbarOption);
#if DEBUG
            cmdIndex2 = cmdGroup.AddCommandItem2("Run Smoke Tests", -1, "Run automated smoke tests", "Run Tests", 3, "RunSmokeTests", "", mainItemID3, menuToolbarOption);
#endif

                cmdGroup.HasToolbar = true;
                cmdGroup.HasMenu = true;
                cmdGroup.Activate();

                bool bResult;
            
            
                foreach (int type in docTypes)
                {
                    CommandTab cmdTab;

                    cmdTab = iCmdMgr.GetCommandTab(type, Title);

                if (cmdTab != null & !getDataResult | ignorePrevious)//if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
                {
                    bool res = iCmdMgr.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                //if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                    if (cmdTab == null)
                    {
                        cmdTab = iCmdMgr.AddCommandTab(type, Title);

                        CommandTabBox cmdBox = cmdTab.AddCommandTabBox();

#if DEBUG
                        int[] cmdIDs = new int[2];
                        int[] TextType = new int[2];

                        cmdIDs[0] = cmdGroup.ToolbarId;
                        TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                        cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex2);
                        TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
#else
                        int[] cmdIDs = new int[1];
                        int[] TextType = new int[1];

                        cmdIDs[0] = cmdGroup.ToolbarId;
                        TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
#endif

                        bResult = cmdBox.AddCommands(cmdIDs, TextType);
                    }

                }
                thisAssembly = null;
              
            }

        public void RemoveCommandMgr()
        {
            iBmp.Dispose();
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
                SwEventPtr.ActiveDocChangeNotify += new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 += new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 += new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify += new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
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
                SwEventPtr.ActiveDocChangeNotify -= new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 -= new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 -= new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify -= new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
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
        //Events
        public int OnDocChange()
        {
            return 0;
        }

        public int OnDocLoad(string docTitle, string docPath)
        {
            return 0;
        }

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

        public int OnModelChange()
        {
            return 0;
        }

        #endregion

        /// <summary>
        /// Runs the unified two-pass workflow: validate all → show problems → process good.
        /// </summary>
        public void RunPipeline()
        {
            NM.Core.ErrorHandler.PushCallStack("RunPipeline");
            try
            {
                var dispatcher = new NM.SwAddin.Pipeline.WorkflowDispatcher(iSwApp);
                dispatcher.Run();
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
                        var added = SolidWorksApiWrapper.AddCustomProperty(doc, "TestProp", swCustomInfoType_e.swCustomInfoText, "Hello", "");
                        SolidWorksApiWrapper.GetCustomProperties(doc, "", out var names, out var types, out var values);
                        var idx = Array.FindIndex(names, n => string.Equals(n, "TestProp", StringComparison.OrdinalIgnoreCase));
                        bool found = idx >= 0;
                        var deleted = SolidWorksApiWrapper.DeleteCustomProperty(doc, "TestProp", "");
                        sb.AppendLine($"CustomProps: add={(added ? "OK" : "FAIL")}, found={(found ? "YES" : "NO")}, delete={(deleted ? "OK" : "FAIL")}");
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine("CustomProps: FAIL - " + ex.Message);
                    }

                    // Rebuild
                    try
                    {
                        bool rebuilt = SolidWorksApiWrapper.ForceRebuildDoc(doc);
                        sb.AppendLine("Rebuild: " + (rebuilt ? "OK" : "SKIPPED/NO-OP"));
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine("Rebuild: FAIL - " + ex.Message);
                    }

                    // Sketch: On Top Plane, draw a short line, then exit
                    try
                    {
                        if (SolidWorksApiWrapper.StartSketchOnPlane(doc, "Top Plane"))
                        {
                            var seg = SolidWorksApiWrapper.CreateSketchLine(doc, 0, 0, 0.02, 0.02);
                            bool ended = SolidWorksApiWrapper.EndSketch(doc);
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
    /// <summary>Configuration settings for Northern Manufacturing SolidWorks Automator.</summary>
    public static class Configuration
    {
        /// <summary>Logging and debugging settings.</summary>
        public static class Logging
        {
            /// <summary>Enable or disable logging.</summary>
            public const bool LogEnabled = true;

            /// <summary>Log file location for errors.</summary>
            public const string ErrorLogPath = @"C:\SolidWorksMacroLogs\ErrorLog.txt";

            /// <summary>Show pop-ups for warnings.</summary>
            public const bool ShowWarnings = false;

            /// <summary>Production mode suppresses verbose logging.</summary>
            public const bool ProductionMode = false;

            /// <summary>Enable performance monitoring/timing.</summary>
            public const bool EnablePerformanceMonitoring = true;

            /// <summary>Set to true to enable verbose debugging.</summary>
            public const bool EnableDebugMode = true;

            /// <summary>Returns whether production mode is enabled.</summary>
            public static bool IsProductionMode => ProductionMode;

            /// <summary>Returns whether performance monitoring is enabled.</summary>
            public static bool IsPerformanceMonitoringEnabled => EnablePerformanceMonitoring;
        }

        /// <summary>File system locations used by the application.</summary>
        public static class FilePaths
        {
            /// <summary>Path to the material Excel file.</summary>
            public const string MaterialFilePath = @"O:\\Engineering Department\\Solidworks\\Macros\\(Semi)Autopilot\\Laser2022v4.xlsx";

            /// <summary>Path to the laser data Excel file.</summary>
            public const string LaserDataFilePath = @"O:\\Engineering Department\\Solidworks\\Macros\\(Semi)Autopilot\\Laser2022v4.xlsx";

            /// <summary>Path to the SolidWorks materials database.</summary>
            public const string MaterialPropertyFilePath = @"C:\\Program Files\\SolidWorks Corp\\SolidWorks\\lang\\english\\sldmaterials\\SolidWorks Materials.sldmat";

            /// <summary>Network bend table for Stainless Steel.</summary>
            public const string BendTableSs = @"O:\\Engineering Department\\Solidworks\\Bend Tables\\StainlessSteel.xlsx";

            /// <summary>Network bend table for Carbon Steel.</summary>
            public const string BendTableCs = @"O:\\Engineering Department\\Solidworks\\Bend Tables\\CarbonSteel.xlsx";

            /// <summary>Local fallback bend table for Stainless Steel.</summary>
            public const string BendTableSsLocal = @"C:\\Program Files\\SolidWorks Corp\\SolidWorks\\lang\\english\\Sheet Metal Bend Tables\\Stainless Steel.xls";

            /// <summary>Local fallback bend table for Carbon Steel.</summary>
            public const string BendTableCsLocal = @"C:\\Program Files\\SolidWorks Corp\\SolidWorks\\lang\\english\\Sheet Metal Bend Tables\\Steel - Mild Steel.xls";

            /// <summary>Special value indicating no bend table should be used (use K-factor).</summary>
            public const string BendTableNone = "-1";

            /// <summary>Excel lookup file used for cutting information.</summary>
            public const string ExcelLookupFile = "NewLaser.xls";

            /// <summary>Path to the ExtractData add-in DLL for external start.</summary>
            public const string ExtractDataAddInPath = @"C:\Program Files\SolidWorks Corp\SolidWorks\Toolbox\data collector\ExtractData.dll";

            // TODO(vNext): Validate the paths at startup and provide user-friendly guidance if missing.
        }

        /// <summary>Manufacturing rates, costs, and processing parameters.</summary>
        public static class Manufacturing
        {
            // CalculateBendInfo constants
            /// <summary>Processing rate 1 in seconds.</summary>
            public const double Rate1Seconds = 10; // seconds
            /// <summary>Processing rate 2 in seconds.</summary>
            public const double Rate2Seconds = 30; // seconds
            /// <summary>Processing rate 3 in seconds.</summary>
            public const double Rate3Seconds = 45; // seconds
            /// <summary>Processing rate 4 in seconds.</summary>
            public const double Rate4Seconds = 200; // seconds
            /// <summary>Processing rate 5 in seconds.</summary>
            public const double Rate5Seconds = 400; // seconds
            /// <summary>Minutes per foot for brake setup.</summary>
            public const double SetupRateMinutesPerFoot = 1.25; // minutes/ft
            /// <summary>Brake setup constant in minutes.</summary>
            public const double BrakeSetupMinutes = 10; // minutes
            /// <summary>Max weight for rate 3 (lbs).</summary>
            public const double Rate3MaxWeightLbs = 100; // lbs
            /// <summary>Max weight for rate 2 (lbs).</summary>
            public const double Rate2MaxWeightLbs = 40; // lbs
            /// <summary>Max weight for rate 1 (lbs).</summary>
            public const double Rate1MaxWeightLbs = 5; // lbs
            /// <summary>Max length for rate 1 (inches).</summary>
            public const double Rate1MaxLengthIn = 12; // in
            /// <summary>Max length for rate 2 (inches).</summary>
            public const double Rate2MaxLengthIn = 60; // in
            /// <summary>Laser setup time per sheet in minutes.</summary>
            public const double LaserSetupRateMinutesPerSheet = 5; // minutes
            /// <summary>Laser setup fixed time in minutes.</summary>
            public const double LaserSetupFixedMinutes = 0.5; // minutes
            /// <summary>Waterjet setup fixed time in minutes.</summary>
            public const double WaterJetSetupFixedMinutes = 15; // minutes
            /// <summary>Waterjet setup time per sheet load in minutes.</summary>
            public const double WaterJetSetupRateMinutesPerLoad = 30; // minutes
            /// <summary>Standard sheet width in inches.</summary>
            public const double StandardSheetWidthIn = 60; // in
            /// <summary>Standard sheet length in inches.</summary>
            public const double StandardSheetLengthIn = 120; // in

            // Standard Costs $/hr
            public const double F115CostPerHour = 120;
            public const double F300CostPerHour = 44;
            public const double F210CostPerHour = 42;
            public const double F140CostPerHour = 80;
            public const double F145CostPerHour = 175;
            public const double F155CostPerHour = 120;
            public const double F325CostPerHour = 65;
            public const double F400CostPerHour = 48;
            public const double F385CostPerHour = 37;
            public const double F500CostPerHour = 48;
            public const double F525CostPerHour = 47;
            public const double EngCostPerHour = 50;

            /// <summary>Material markup multiplier.</summary>
            public const double MaterialMarkup = 1.05;  // 5%
            /// <summary>Tight tolerance multiplier.</summary>
            public const double TightPercent = 1.15; // 15%
            /// <summary>Normal tolerance multiplier.</summary>
            public const double NormalPercent = 1.0;  // 0%
            /// <summary>Loose tolerance multiplier.</summary>
            public const double LoosePercent = 0.95; // -5%

            // CalculateCutInfo constants
            /// <summary>Constant added to calculated pierce total.</summary>
            public const double PierceConstant = 2;
            /// <summary>Tab spacing in units consistent with process planning (typically mm or in).</summary>
            public const int TabSpacing = 30;
        }

        /// <summary>Material-related constants such as unit conversions and default property sets.</summary>
        public static class Materials
        {
            /// <summary>Conversion factor from meters to inches.</summary>
            public const double MetersToInches = 39.3701;
            /// <summary>Conversion factor from inches to meters.</summary>
            public const double InchesToMeters = 1.0 / 39.3701;
            /// <summary>Conversion factor from kilograms to pounds.</summary>
            public const double KgToLbs = 2.20462;
            /// <summary>Density of steel in pounds per cubic inch.</summary>
            public const double SteelDensityLbsPerIn3 = 0.284;

            /// <summary>Initial custom property names used on models.</summary>
            public static readonly IReadOnlyList<string> InitialCustomProperties = new[]
            {
                "IsSheetMetal", "IsTube", "Thickness", "Description", "Customer",
                "CustPartNumber", "CuttingType", "Drawing", "ExportDate", "F115_Hours",
                "F115_Price", "F210_Hours", "F210_Price", "Length", "MaterialCostPerLB",
                "Model", "MPNumber", "RawWeight", "Revision", "Total_Weight"
            };

            /// <summary>Returns the initial custom property names sequence.</summary>
            public static IReadOnlyList<string> GetInitialCustomProperties() => InitialCustomProperties;
        }

        /// <summary>General application defaults and user preferences.</summary>
        public static class Defaults
        {
            /// <summary>Number of retries for Excel file reads.</summary>
            public const int MaxRetries = 3;
            /// <summary>Default worksheet name to use when reading from Excel.</summary>
            public const string DefaultSheetName = "Sheet1";
            /// <summary>Whether Excel should be automatically closed after operations.</summary>
            public const bool AutoCloseExcel = true;

            /// <summary>Conservative default cost per pound when quoting.</summary>
            public const double DefaultCostPerLb = 3.5;
            /// <summary>Default order quantity.</summary>
            public const int DefaultQuantity = 1;
            /// <summary>Typical steel K-factor when not using a bend table.</summary>
            public const double DefaultKFactor = 0.44;

            // TODO(vNext): Externalize defaults to a JSON config with environment overrides.
        }
    }
}
