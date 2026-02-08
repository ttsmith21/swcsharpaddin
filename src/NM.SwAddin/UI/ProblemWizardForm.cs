using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using NM.Core.DataModel;
using NM.Core.Models;
using NM.Core.Processing;
using NM.Core.ProblemParts;
using NM.SwAddin.Processing;
using NM.SwAddin.Geometry;
using NM.SwAddin.Pipeline;
using NM.SwAddin.Validation;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using static NM.Core.Constants.UnitConversions;

namespace NM.SwAddin.UI
{
    /// <summary>
    /// Wizard-style form for stepping through problem parts one at a time.
    /// Non-modal so user can interact with SolidWorks while fixing issues.
    /// </summary>
    public sealed class ProblemWizardForm : Form
    {
        private readonly List<ProblemPartManager.ProblemItem> _problems;
        private readonly List<ProblemPartManager.ProblemItem> _fixedProblems = new List<ProblemPartManager.ProblemItem>();
        private readonly HashSet<ProblemPartManager.ProblemItem> _reviewedItems = new HashSet<ProblemPartManager.ProblemItem>();
        private readonly ISldWorks _swApp;
        private int _currentIndex;
        private int _initialGoodCount;
        private IModelDoc2 _currentDoc;
        private TubeDiagnosticInfo _tubeDiagnostics;

        // Navigation
        private Label _lblProgress;
        private Button _btnPrevious;
        private Button _btnNext;

        // Problem info
        private Label _lblFileName;
        private Label _lblPath;
        private Label _lblConfig;
        private Label _lblCategory;
        private TextBox _txtError;
        private ListBox _lstSuggestions;

        // Part type classification
        private GroupBox _grpPartType;
        private Button _btnMarkPUR;
        private Button _btnMarkMACH;
        private Button _btnMarkCUST;
        private Button _btnMarkSM;
        private Button _btnMarkTUBE;
        private Button _btnSplit;

        // Tube diagnostics
        private GroupBox _grpTubeDiag;
        private Button _btnShowCutLength;
        private Button _btnShowHoles;
        private Button _btnShowBoundary;
        private Button _btnShowProfile;
        private Button _btnShowAll;
        private Button _btnClearSelection;
        private Label _lblDiagStatus;

        // Actions
        private Button _btnRetry;
        private Button _btnSkip;
        private Button _btnRevertNesting;
        private Button _btnFinish;
        private Button _btnCancel;

        // Progress
        private ProgressBar _progressBar;
        private Label _lblFixedCount;
        private Label _lblStatus;

        public ProblemAction SelectedAction { get; private set; } = ProblemAction.Cancel;
        public List<ProblemPartManager.ProblemItem> FixedProblems => _fixedProblems;

        public ProblemWizardForm(List<ProblemPartManager.ProblemItem> problems, ISldWorks swApp, int initialGoodCount)
        {
            _problems = problems ?? new List<ProblemPartManager.ProblemItem>();
            _swApp = swApp;
            _initialGoodCount = initialGoodCount;
            _currentIndex = 0;

            Text = "Problem Part Wizard";
            Width = 600;
            Height = 720;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = true;
            ShowInTaskbar = true;
            TopMost = true;

            BuildUI();
            LoadCurrentProblem();
            UpdateProgress();
        }

        private void BuildUI()
        {
            // Navigation header
            _lblProgress = new Label
            {
                Left = 20, Top = 15, Width = 200, Height = 25,
                Font = new Font(Font.FontFamily, 12, FontStyle.Bold),
                Text = "Problem 1 of X"
            };

            _btnPrevious = new Button { Left = 380, Top = 12, Width = 90, Height = 28, Text = "<< Previous" };
            _btnPrevious.Click += (s, e) => NavigatePrevious();

            _btnNext = new Button { Left = 480, Top = 12, Width = 90, Height = 28, Text = "Next >>" };
            _btnNext.Click += (s, e) => NavigateNext();

            // Separator
            var separator1 = new Label { Left = 20, Top = 45, Width = 550, Height = 2, BorderStyle = BorderStyle.Fixed3D };

            // File info
            _lblFileName = new Label { Left = 20, Top = 55, Width = 550, Height = 22, Font = new Font(Font.FontFamily, 10, FontStyle.Bold) };
            _lblPath = new Label { Left = 20, Top = 77, Width = 550, Height = 18, ForeColor = Color.DimGray };
            _lblConfig = new Label { Left = 20, Top = 95, Width = 300, Height = 18 };
            _lblCategory = new Label { Left = 320, Top = 95, Width = 250, Height = 18 };

            // Error box
            var lblError = new Label { Left = 20, Top = 120, Width = 100, Text = "Error:", Font = new Font(Font.FontFamily, 9, FontStyle.Bold) };
            _txtError = new TextBox
            {
                Left = 20, Top = 140, Width = 550, Height = 60,
                Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
                BackColor = Color.MistyRose, ForeColor = Color.DarkRed
            };

            // Suggestions
            var lblSuggestions = new Label { Left = 20, Top = 205, Width = 150, Text = "Suggestions:", Font = new Font(Font.FontFamily, 9, FontStyle.Bold) };
            _lstSuggestions = new ListBox { Left = 20, Top = 225, Width = 550, Height = 80 };

            // Part type classification group (PUR/MACH/CUST + SM/TUBE)
            _grpPartType = new GroupBox { Left = 20, Top = 310, Width = 550, Height = 90, Text = "Classify As" };

            _btnMarkPUR = new Button { Left = 10, Top = 20, Width = 100, Height = 28, Text = "PUR (Purchased)", BackColor = Color.LightYellow };
            _btnMarkPUR.Click += (s, e) => MarkAsPartType(ProblemPartManager.PartTypeOverride.Purchased);

            _btnMarkMACH = new Button { Left = 115, Top = 20, Width = 105, Height = 28, Text = "MACH (Machined)", BackColor = Color.LightYellow };
            _btnMarkMACH.Click += (s, e) => MarkAsPartType(ProblemPartManager.PartTypeOverride.Machined);

            _btnMarkCUST = new Button { Left = 225, Top = 20, Width = 135, Height = 28, Text = "CUST (Customer)", BackColor = Color.LightYellow };
            _btnMarkCUST.Click += (s, e) => MarkAsPartType(ProblemPartManager.PartTypeOverride.CustomerSupplied);

            _btnMarkSM = new Button { Left = 10, Top = 55, Width = 160, Height = 28, Text = "Sheet Metal (convert)", BackColor = Color.LightSteelBlue };
            _btnMarkSM.Click += (s, e) => RunAsClassification("SheetMetal");

            _btnMarkTUBE = new Button { Left = 180, Top = 55, Width = 160, Height = 28, Text = "Tube (convert)", BackColor = Color.LightSteelBlue };
            _btnMarkTUBE.Click += (s, e) => RunAsClassification("Tube");

            _btnSplit = new Button { Left = 350, Top = 55, Width = 160, Height = 28, Text = "Split \u2192 Assy", BackColor = Color.LightCoral };
            _btnSplit.Click += (s, e) => RunSplitToAssembly();

            _grpPartType.Controls.AddRange(new Control[] { _btnMarkPUR, _btnMarkMACH, _btnMarkCUST, _btnMarkSM, _btnMarkTUBE, _btnSplit });

            // Tube diagnostics group
            _grpTubeDiag = new GroupBox { Left = 20, Top = 410, Width = 550, Height = 90, Text = "Tube Diagnostics (for tube/structural parts)" };

            _btnShowCutLength = new Button { Left = 10, Top = 22, Width = 80, Height = 26, Text = "Cut Length" };
            _btnShowCutLength.Click += (s, e) => ShowTubeDiagnostic(TubeDiagnosticKind.CutLength);

            _btnShowHoles = new Button { Left = 95, Top = 22, Width = 80, Height = 26, Text = "Holes" };
            _btnShowHoles.Click += (s, e) => ShowTubeDiagnostic(TubeDiagnosticKind.Holes);

            _btnShowBoundary = new Button { Left = 180, Top = 22, Width = 80, Height = 26, Text = "Boundary" };
            _btnShowBoundary.Click += (s, e) => ShowTubeDiagnostic(TubeDiagnosticKind.Boundary);

            _btnShowProfile = new Button { Left = 265, Top = 22, Width = 80, Height = 26, Text = "Profile" };
            _btnShowProfile.Click += (s, e) => ShowTubeDiagnostic(TubeDiagnosticKind.Profile);

            _btnShowAll = new Button { Left = 350, Top = 22, Width = 80, Height = 26, Text = "Show All" };
            _btnShowAll.Click += (s, e) => ShowTubeDiagnostic(TubeDiagnosticKind.All);

            _btnClearSelection = new Button { Left = 435, Top = 22, Width = 80, Height = 26, Text = "Clear" };
            _btnClearSelection.Click += (s, e) => ShowTubeDiagnostic(TubeDiagnosticKind.Clear);

            _lblDiagStatus = new Label { Left = 10, Top = 55, Width = 530, Height = 30, Text = "Part opens automatically. Use buttons to highlight geometry." };

            _grpTubeDiag.Controls.AddRange(new Control[] { _btnShowCutLength, _btnShowHoles, _btnShowBoundary, _btnShowProfile, _btnShowAll, _btnClearSelection, _lblDiagStatus });

            // Separator
            var separator2 = new Label { Left = 20, Top = 510, Width = 550, Height = 2, BorderStyle = BorderStyle.Fixed3D };

            // Action buttons
            _btnRetry = new Button { Left = 20, Top = 520, Width = 130, Height = 35, Text = "Retry && Validate", BackColor = Color.LightGreen };
            _btnRetry.Click += OnRetry;

            _btnSkip = new Button { Left = 160, Top = 520, Width = 90, Height = 35, Text = "Skip" };
            _btnSkip.Click += OnSkip;

            _btnRevertNesting = new Button { Left = 460, Top = 520, Width = 110, Height = 35, Text = "Revert to 80%", BackColor = Color.LightSalmon, Visible = false };
            _btnRevertNesting.Click += OnRevertNestingOverride;

            _btnFinish = new Button { Left = 260, Top = 520, Width = 90, Height = 35, Text = "Finish" };
            _btnFinish.Click += OnFinish;

            _btnCancel = new Button { Left = 360, Top = 520, Width = 90, Height = 35, Text = "Cancel" };
            _btnCancel.Click += OnCancel;

            // Progress section
            _progressBar = new ProgressBar { Left = 20, Top = 570, Width = 400, Height = 22 };
            _lblFixedCount = new Label { Left = 430, Top = 573, Width = 140, Height = 20 };

            _lblStatus = new Label { Left = 20, Top = 600, Width = 550, Height = 40, ForeColor = Color.DarkBlue };

            // Add all controls
            Controls.AddRange(new Control[]
            {
                _lblProgress, _btnPrevious, _btnNext, separator1,
                _lblFileName, _lblPath, _lblConfig, _lblCategory,
                lblError, _txtError, lblSuggestions, _lstSuggestions,
                _grpPartType, _grpTubeDiag, separator2,
                _btnRetry, _btnSkip, _btnRevertNesting, _btnFinish, _btnCancel,
                _progressBar, _lblFixedCount, _lblStatus
            });
        }

        private void LoadCurrentProblem()
        {
            // Save & close the previous part before loading the next one
            SaveAndCloseCurrentPart();

            if (_problems.Count == 0)
            {
                _lblFileName.Text = "(No problems to display)";
                _lblPath.Text = "";
                _lblConfig.Text = "";
                _lblCategory.Text = "";
                _txtError.Text = "";
                _lstSuggestions.Items.Clear();
                DisableActions();
                return;
            }

            if (_currentIndex < 0) _currentIndex = 0;
            if (_currentIndex >= _problems.Count) _currentIndex = _problems.Count - 1;

            var item = _problems[_currentIndex];

            _lblProgress.Text = $"Problem {_currentIndex + 1} of {_problems.Count}";
            _lblFileName.Text = item.DisplayName ?? "(Unknown)";
            _lblPath.Text = item.FilePath ?? "";
            _lblConfig.Text = $"Config: {item.Configuration ?? "(default)"}";
            _lblCategory.Text = $"Category: {item.Category}";
            _txtError.Text = item.ProblemDescription ?? "Unknown error";
            _txtError.BackColor = Color.MistyRose;

            LoadSuggestions(item);
            UpdateNavigationButtons();
            UpdatePartTypeHint(item);

            // Show/hide nesting efficiency revert button
            _btnRevertNesting.Visible = (item.Category == ProblemPartManager.ProblemCategory.NestingEfficiencyOverride);

            // Reset tube diagnostics
            _tubeDiagnostics = null;
            _lblDiagStatus.Text = "Part opens automatically. Use buttons to highlight geometry.";

            _lblStatus.Text = "Part opened. Fix the issue, then click 'Retry & Validate'.";

            // Auto-open the part in SolidWorks
            OpenCurrentPart();
        }

        private void LoadSuggestions(ProblemPartManager.ProblemItem item)
        {
            _lstSuggestions.Items.Clear();
            var suggestions = ProblemSuggestionProvider.GetSuggestions(item.Category, item.ProblemDescription);
            foreach (var s in suggestions)
            {
                _lstSuggestions.Items.Add(s);
            }
        }

        private void UpdateNavigationButtons()
        {
            _btnPrevious.Enabled = _currentIndex > 0;
            _btnNext.Enabled = _currentIndex < _problems.Count - 1;
        }

        private void UpdateProgress()
        {
            int total = _problems.Count;
            int reviewed = _reviewedItems.Count;
            int fixedCount = _fixedProblems.Count;
            int skipped = reviewed - fixedCount;

            _progressBar.Maximum = total > 0 ? total : 1;
            _progressBar.Value = Math.Min(reviewed, _progressBar.Maximum);

            if (reviewed >= total && total > 0)
            {
                _lblFixedCount.Text = $"Done: {fixedCount} fixed, {skipped} skipped";
                _lblStatus.Text = $"All {total} problems reviewed. Click Finish to continue.";
                _lblStatus.ForeColor = Color.DarkGreen;
                _btnFinish.Text = "Finish";
                _btnFinish.BackColor = Color.LightGreen;
                _btnFinish.Font = new Font(_btnFinish.Font, FontStyle.Bold);
            }
            else
            {
                _lblFixedCount.Text = $"{reviewed} / {total} reviewed ({fixedCount} fixed)";
                _btnFinish.Text = $"Finish ({_initialGoodCount + fixedCount} ready)";
            }
        }

        private void DisableActions()
        {
            _btnRetry.Enabled = false;
            _btnSkip.Enabled = false;
        }

        private void NavigatePrevious()
        {
            if (_currentIndex > 0)
            {
                _currentIndex--;
                LoadCurrentProblem();
            }
        }

        private void NavigateNext()
        {
            if (_currentIndex < _problems.Count - 1)
            {
                _currentIndex++;
                LoadCurrentProblem();
            }
        }

        private void SaveAndCloseCurrentPart()
        {
            if (_currentDoc == null) return;
            ProblemPartActions.SaveAndClosePart(_currentDoc, _swApp);
            _currentDoc = null;
            _tubeDiagnostics = null;
        }

        private void OpenCurrentPart()
        {
            if (_swApp == null) return;
            if (_problems.Count == 0 || _currentIndex >= _problems.Count) return;

            var item = _problems[_currentIndex];
            _currentDoc = ProblemPartActions.OpenPart(item, _swApp);

            if (_currentDoc == null)
                _lblStatus.Text = $"Failed to open: {item.DisplayName}";
            else
                _lblDiagStatus.Text = "Part is open. Use buttons above to highlight geometry for debugging.";
        }

        private void UpdatePartTypeHint(ProblemPartManager.ProblemItem item)
        {
            object hintObj;
            if (item.Metadata.TryGetValue("PurchasedHint", out hintObj) && hintObj is string hint && !string.IsNullOrEmpty(hint))
            {
                _grpPartType.Text = $"Classify As (HINT: Likely Purchased - {hint})";
                _grpPartType.ForeColor = Color.DarkOrange;
                _btnMarkPUR.BackColor = Color.Gold;
            }
            else
            {
                _grpPartType.Text = "Classify As (skips processing)";
                _grpPartType.ForeColor = SystemColors.ControlText;
                _btnMarkPUR.BackColor = Color.LightYellow;
            }
            _btnMarkMACH.BackColor = Color.LightYellow;
            _btnMarkCUST.BackColor = Color.LightYellow;
        }

        private void MarkAsPartType(ProblemPartManager.PartTypeOverride typeOverride)
        {
            if (_problems.Count == 0 || _currentIndex >= _problems.Count) return;
            var item = _problems[_currentIndex];

            var result = ProblemPartActions.ClassifyPart(item, typeOverride, _currentDoc, _fixedProblems);
            _reviewedItems.Add(item);

            _lblStatus.Text = result.Message;
            _txtError.BackColor = Color.LightGoldenrodYellow;
            _txtError.Text = $"Classified as {typeOverride} (rbPartType=1, rbPartTypeSub={(int)typeOverride}) - will skip classification pipeline";

            UpdateProgress();
            AutoAdvance();
        }

        private void RunAsClassification(string classification)
        {
            if (_currentDoc == null || _problems.Count == 0 || _currentIndex >= _problems.Count) return;
            var item = _problems[_currentIndex];

            _lblStatus.Text = $"Processing as {classification}...";
            Application.DoEvents();

            var result = ProblemPartActions.RunAsClassification(item, classification, _swApp, _currentDoc, _fixedProblems);

            if (result.Success)
            {
                _reviewedItems.Add(item);
                _lblStatus.Text = result.Message;
                _txtError.BackColor = Color.LightGreen;
                _txtError.Text = $"Successfully processed as {classification}";
                UpdateProgress();
                AutoAdvance();
            }
            else
            {
                _lblStatus.Text = result.Message;
                _txtError.BackColor = Color.LightSalmon;
                _txtError.Text = result.Message;
            }
        }

        private void RunSplitToAssembly()
        {
            if (_currentDoc == null || _problems.Count == 0 || _currentIndex >= _problems.Count) return;
            var item = _problems[_currentIndex];

            _lblStatus.Text = "Splitting multi-body part...";
            Application.DoEvents();

            var result = ProblemPartActions.RunSplitToAssembly(
                item, _swApp, _currentDoc, _fixedProblems,
                status => { _lblStatus.Text = status; Application.DoEvents(); });

            if (result.Success)
            {
                _reviewedItems.Add(item);
                _lblStatus.Text = result.Message;
                _txtError.BackColor = Color.LightGreen;
                _txtError.Text = result.Message;
                UpdateProgress();
                AutoAdvance();
            }
            else
            {
                _lblStatus.Text = result.Message;
                _txtError.BackColor = Color.LightSalmon;
                _txtError.Text = result.Message;
            }
        }

        private void ShowTubeDiagnostic(TubeDiagnosticKind kind)
        {
            if (kind != TubeDiagnosticKind.Clear)
            {
                string statusMsg;
                if (!ProblemPartActions.EnsureTubeDiagnostics(_swApp, _currentDoc, ref _tubeDiagnostics, out statusMsg))
                {
                    _lblDiagStatus.Text = statusMsg;
                    return;
                }
            }

            _lblDiagStatus.Text = ProblemPartActions.SelectTubeDiagnostic(_swApp, _currentDoc, _tubeDiagnostics, kind);
        }

        private void AutoAdvance()
        {
            if (_currentIndex < _problems.Count - 1)
            {
                _lblStatus.Text += " - Moving to next problem...";
                Application.DoEvents();
                System.Threading.Thread.Sleep(500);
                _currentIndex++;
                LoadCurrentProblem();
            }
            else
            {
                UpdateProgress();
            }
        }

        private void OnRetry(object sender, EventArgs e)
        {
            if (_swApp == null || _problems.Count == 0) return;

            var item = _problems[_currentIndex];
            _lblStatus.Text = $"Validating {item.DisplayName}...";
            Application.DoEvents();

            var result = ProblemPartActions.RetryAndValidate(item, _swApp, _currentDoc, _fixedProblems);

            if (result.Success)
            {
                _reviewedItems.Add(item);
                _lblStatus.Text = result.Message;
                _txtError.BackColor = Color.LightGreen;
                _txtError.Text = "Validation passed!";
                UpdateProgress();

                if (_currentIndex < _problems.Count - 1)
                {
                    _lblStatus.Text += " - Moving to next problem...";
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(800);
                    _currentIndex++;
                    LoadCurrentProblem();
                }
                else
                {
                    UpdateProgress();
                }
            }
            else
            {
                _txtError.Text = item.ProblemDescription;
                _txtError.BackColor = Color.MistyRose;
                LoadSuggestions(item);
                _lblStatus.Text = result.Message;
            }
        }

        private void OnSkip(object sender, EventArgs e)
        {
            if (_problems.Count > 0 && _currentIndex < _problems.Count)
            {
                var item = _problems[_currentIndex];
                ProblemPartActions.MarkAsSkipped(item, _currentDoc);
                _reviewedItems.Add(item);
            }

            if (_currentIndex < _problems.Count - 1)
            {
                UpdateProgress();
                _currentIndex++;
                LoadCurrentProblem();
            }
            else
            {
                SaveAndCloseCurrentPart();
                UpdateProgress();
            }
        }

        /// <summary>
        /// Reverts a nesting efficiency auto-override back to the default 80% efficiency mode.
        /// Writes rbWeightCalc=0 and NestEfficiency=80 to the part's custom properties.
        /// </summary>
        private void OnRevertNestingOverride(object sender, EventArgs e)
        {
            if (_problems.Count == 0 || _currentIndex >= _problems.Count) return;
            var item = _problems[_currentIndex];

            if (item.Category != ProblemPartManager.ProblemCategory.NestingEfficiencyOverride)
                return;

            try
            {
                if (_currentDoc != null)
                {
                    // Revert to efficiency mode with 80% default
                    SwPropertyHelper.AddCustomProperty(_currentDoc, "rbWeightCalc",
                        swCustomInfoType_e.swCustomInfoText, "0", "");
                    SwPropertyHelper.AddCustomProperty(_currentDoc, "NestEfficiency",
                        swCustomInfoType_e.swCustomInfoNumber, "80", "");
                    SwDocumentHelper.SaveDocument(_currentDoc);

                    NM.Core.ErrorHandler.DebugLog($"[Wizard] Reverted nesting override: {item.DisplayName} -> 80% efficiency mode");
                }

                // Mark as resolved
                _fixedProblems.Add(item);
                _reviewedItems.Add(item);
                ProblemPartManager.Instance.RemoveResolvedPart(item);

                _lblStatus.Text = $"Reverted to 80% efficiency: {item.DisplayName}";
                _txtError.BackColor = Color.LightGoldenrodYellow;
                _txtError.Text = "Reverted to default 80% nesting efficiency mode. Part will use standard material estimate.";

                UpdateProgress();
                AutoAdvance();
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Revert failed: {ex.Message}";
                NM.Core.ErrorHandler.DebugLog($"[Wizard] Revert nesting override failed: {ex.Message}");
            }
        }

        private void OnFinish(object sender, EventArgs e)
        {
            SaveAndCloseCurrentPart();
            SelectedAction = ProblemAction.ContinueWithGood;
            Close();
        }

        private void OnCancel(object sender, EventArgs e)
        {
            SaveAndCloseCurrentPart();
            SelectedAction = ProblemAction.Cancel;
            Close();
        }
    }
}
