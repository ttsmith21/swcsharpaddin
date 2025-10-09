using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwCSharpAddin1
{
    [Guid("5b7c6d57-06b3-4a66-b899-2b1b0b7f2c55")]
    public class SwAddin : ISwAddin
    {
        // ...existing fields...

        #region UI Callbacks
        public void CreateCube()
        {
            // ...existing code...
        }

        public void ShowProblemPartsUI()
        {
            try
            {
                var form = new NM.SwAddin.UI.ProblemPartsForm(NM.Core.ProblemParts.ProblemPartManager.Instance);
                form.StartPosition = FormStartPosition.CenterScreen;
                form.Show();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Problem Parts UI error: " + ex.Message);
            }
        }

        public void RunFolderProcessing()
        {
            try
            {
                using (var dlg = new FolderBrowserDialog())
                {
                    dlg.Description = "Select a root folder to process";
                    if (dlg.ShowDialog() != DialogResult.OK) return;

                    var processor = new NM.SwAddin.Processing.FolderProcessor(iSwApp);
                    var res = processor.ProcessFolder(dlg.SelectedPath, recursive: true);

                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine("Folder Processing Summary:");
                    sb.AppendLine($"Root: {dlg.SelectedPath}");
                    sb.AppendLine($"Discovered: {res.TotalDiscovered}");
                    sb.AppendLine($"Imported: {res.ImportedCount}");
                    sb.AppendLine($"Opened: {res.OpenedOk}");
                    sb.AppendLine($"Failed: {res.FailedOpen}");
                    sb.AppendLine($"Skipped: {res.Skipped}");
                    if (res.Errors.Count > 0)
                    {
                        sb.AppendLine();
                        sb.AppendLine("Errors:");
                        foreach (var e in res.Errors)
                            sb.AppendLine(" - " + e);
                    }
                    MessageBox.Show(sb.ToString(), "Process Folder");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Folder processing error: " + ex.Message);
            }
        }

        // NEW: Auto workflow callback
        public void RunAutoWorkflow()
        {
            try
            {
                NM.SwAddin.Pipeline.AutoWorkflow.Run(iSwApp);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Auto workflow error: " + ex.Message);
            }
        }
        #endregion

        // ...existing code for CommandManager setup...
        private void AddCommandMgr()
        {
            // ...existing command group creation code...

            // After existing group is activated, register the Auto group
            AddAutoWorkflowCommandGroup();
        }

        // NEW: separate CommandGroup for the Auto Run button
        private void AddAutoWorkflowCommandGroup()
        {
            try
            {
                int err = 0;
                const int AutoCmdGroupId = 0x5A17; // unique group id
                bool ignorePrev = true;

                var autoGrp = iCmdMgr.CreateCommandGroup2(
                    AutoCmdGroupId,
                    "NM Auto",
                    "Auto-run workflow (Assembly/Part/Folder)",
                    "",
                    -1,
                    ignorePrev,
                    ref err);

                int menuToolbar = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
                autoGrp.AddCommandItem2(
                    "Auto Run",
                    -1,
                    "Auto dispatch: assembly ? assembly flow; part ? single-part; none ? folder batch",
                    "Auto Run",
                    0,
                    nameof(RunAutoWorkflow),
                    "",
                    1,
                    menuToolbar);

                autoGrp.HasToolbar = true;
                autoGrp.HasMenu = true;
                autoGrp.Activate();

                foreach (int dt in new[] { (int)swDocumentTypes_e.swDocPART, (int)swDocumentTypes_e.swDocASSEMBLY })
                {
                    var tab = iCmdMgr.GetCommandTab(dt, "NM Auto");
                    if (tab != null)
                    {
                        try { iCmdMgr.RemoveCommandTab(tab); } catch { }
                        tab = null;
                    }
                    tab = iCmdMgr.AddCommandTab(dt, "NM Auto");
                    var box = tab.AddCommandTabBox();

                    int cmdId = autoGrp.get_CommandID(0);
                    int disp = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    int[] ids = { cmdId };
                    int[] text = { disp };
                    box.AddCommands(ids, text);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Failed to add NM Auto command group: " + ex.Message);
            }
        }
    }
}
