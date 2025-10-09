using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;
using System.Windows.Forms;
using NM.StepClassifierAddin.Interop;
using NM.StepClassifierAddin.Utils;
using NM.StepClassifierAddin.Classification;
using NM.Core;
using NM.Core.Models;
using NM.SwAddin.Services;
using NM.SwAddin.Pipeline;
using NM.SwAddin.Validation;

namespace NM.StepClassifierAddin
{
    /// <summary>
    /// NM Step Classifier Addin entry point implementing ISwAddin.
    /// </summary>
    [ComVisible(true)]
    [Guid("3C8B9F88-73F3-4C5A-9A3E-8E8C4E7F7B7D")] // NOTE: Generate a new GUID for production
    [ProgId("NM.StepClassifierAddin")]
    public class Addin : ISwAddin
    {
        private ISldWorks _sw; private int _cookie; private ICommandManager _cmdMgr;
        private const int GroupId = 41; private const int BtnClassifyId = 0; private const int BtnPreflightId = 1; private const int BtnListProblemsId = 2; private const int BtnValidateId = 3; private const int BtnRunId = 4;
        private readonly ProblemPartTracker _problemTracker = new ProblemPartTracker();

        [ComRegisterFunction]
        public static void Register(Type t) => ComRegistration.RegisterAddin(t, "NM Step Classifier", "Classifies STEP bodies: Stick / Sheet / Other");
        [ComUnregisterFunction]
        public static void Unregister(Type t) => ComRegistration.UnregisterAddin(t);

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            _sw = (ISldWorks)ThisSW; _cookie = cookie; _sw.SetAddinCallbackInfo2(0, this, _cookie);
            _cmdMgr = _sw.GetCommandManager(_cookie);
            AddCommandMgr();
            return true;
        }
        public bool DisconnectFromSW()
        {
            try { _cmdMgr.RemoveCommandGroup(GroupId); } catch { }
            _cmdMgr = null; _sw = null; return true;
        }

        private void AddCommandMgr()
        {
            int err = 0; bool ignorePrev = true; // force recreate
            var grp = _cmdMgr.CreateCommandGroup2(GroupId, "NM Classifier", "NM Classifier", "", -1, ignorePrev, ref err);
            grp.AddCommandItem2("Classify Selected Body", -1, "Classify first solid body in active part", "Classify", 0, nameof(ClassifySelectedBody), "", BtnClassifyId, (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));
            grp.AddCommandItem2("Preflight Check", -1, "Run preflight problem checks on the active part", "Preflight", 0, nameof(RunPreflight), "", BtnPreflightId, (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));
            grp.AddCommandItem2("List Problem Parts", -1, "Show recorded problem parts from this session", "Problems", 0, nameof(ShowProblemParts), "", BtnListProblemsId, (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));
            grp.AddCommandItem2("Validate Active Part", -1, "Run core validation pipeline on active part", "Validate", 0, nameof(ValidateActivePart), "", BtnValidateId, (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));
            grp.AddCommandItem2("Run Single-Part Pipeline", -1, "Validate ? Convert to Sheet Metal ? Save (optional)", "Run", 0, nameof(RunSinglePartPipeline), "", BtnRunId, (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));
            grp.HasToolbar = true; grp.HasMenu = true; grp.Activate();

            int[] doctypes = { (int)swDocumentTypes_e.swDocPART };
            foreach (int dt in doctypes)
            {
                var tab = _cmdMgr.GetCommandTab(dt, "NM Classifier");
                if (tab != null) { _cmdMgr.RemoveCommandTab(tab); tab = null; }
                tab = _cmdMgr.AddCommandTab(dt, "NM Classifier");
                var box = tab.AddCommandTabBox();
                int[] ids = { grp.get_CommandID(0), grp.get_CommandID(1), grp.get_CommandID(2), grp.get_CommandID(3), grp.get_CommandID(4) };
                int disp = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                int[] text = { disp, disp, disp, disp, disp };
                box.AddCommands(ids, text);
            }
        }

        public void ClassifySelectedBody()
        {
            try
            {
                var model = _sw?.ActiveDoc as IModelDoc2;
                var part = model as IPartDoc;
                if (part == null) { MessageBox.Show("Open a part document."); return; }
                var body = part.GetFirstSolidBody();
                if (body == null) { MessageBox.Show("No solid body found."); return; }

                var result = PartClassifier.Classify(_sw, model, body);
                MessageBox.Show($"Classification: {result}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        public void RunPreflight()
        {
            try
            {
                var model = _sw?.ActiveDoc as IModelDoc2;
                if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocPART)
                { MessageBox.Show("Open a part document."); return; }

                string cfg = string.Empty;
                try { cfg = model.ConfigurationManager?.ActiveConfiguration?.Name ?? string.Empty; } catch { }

                var mi = new ModelInfo();
                try { mi.Initialize(model.GetPathName() ?? string.Empty, cfg); } catch { mi.Initialize(string.Empty, cfg); }

                PartPreflight.Result res;
                bool ok = PartPreflight.Evaluate(mi, model, _problemTracker, out res);
                if (ok)
                {
                    MessageBox.Show("Preflight OK: single solid body and no obvious complexity issues.");
                }
                else
                {
                    MessageBox.Show("Problem detected: " + (res.Reason ?? "unknown"));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        public void ShowProblemParts()
        {
            try
            {
                var summary = _problemTracker.GetDisplaySummary();
                MessageBox.Show(summary);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        public void ValidateActivePart()
        {
            try
            {
                var docSvc = new DocumentService(_sw);
                var active = docSvc.GetActiveModel();
                if (active == null || active.Type != SwModelInfo.ModelType.Part) { MessageBox.Show("Open a part document."); return; }

                var validator = new PartValidationAdapter();
                var result = validator.Validate(active, active.ModelDoc);

                // Log performance summary and show result + counters
                ErrorHandler.DebugLog(ValidationStats.BuildSummary());
                ErrorHandler.DebugLog(PerformanceTracker.Instance.BuildSummary());
                _sw.SendMsgToUser2(
                    $"File: {active.FileName}\nState: {active.State}\nResult: {result.Summary}\n{ValidationStats.BuildSummary()}",
                    (int)swMessageBoxIcon_e.swMbInformation,
                    (int)swMessageBoxBtn_e.swMbOk);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        public void RunSinglePartPipeline()
        {
            try
            {
                var runner = new MainRunner(_sw);
                var result = runner.ProcessActivePart(flatten: true, saveOnSuccess: true);
                var text = result.Success ? $"Success: {result.Message}" : $"Failed: {result.Message}";
                _sw.SendMsgToUser2(
                    text + "\n" + ValidationStats.BuildSummary(),
                    result.Success ? (int)swMessageBoxIcon_e.swMbInformation : (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk);
                if (NM.Core.Configuration.Logging.EnableDebugMode)
                {
                    MessageBox.Show(text + "\n" + ValidationStats.BuildSummary(), "NM Pipeline", MessageBoxButtons.OK,
                        result.Success ? MessageBoxIcon.Information : MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
    }
}