using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NM.Core.Models;
using NM.Core.Processing;
using NM.Core.ProblemParts;
using NM.SwAddin.Geometry;
using NM.SwAddin.Pipeline;
using NM.SwAddin.Validation;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

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
            Height = 620;
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

            // Tube diagnostics group
            _grpTubeDiag = new GroupBox { Left = 20, Top = 315, Width = 550, Height = 90, Text = "Tube Diagnostics (for tube/structural parts)" };

            _btnShowCutLength = new Button { Left = 10, Top = 22, Width = 80, Height = 26, Text = "Cut Length" };
            _btnShowCutLength.Click += OnShowCutLengthEdges;

            _btnShowHoles = new Button { Left = 95, Top = 22, Width = 80, Height = 26, Text = "Holes" };
            _btnShowHoles.Click += OnShowHoleEdges;

            _btnShowBoundary = new Button { Left = 180, Top = 22, Width = 80, Height = 26, Text = "Boundary" };
            _btnShowBoundary.Click += OnShowBoundaryEdges;

            _btnShowProfile = new Button { Left = 265, Top = 22, Width = 80, Height = 26, Text = "Profile" };
            _btnShowProfile.Click += OnShowProfileFaces;

            _btnShowAll = new Button { Left = 350, Top = 22, Width = 80, Height = 26, Text = "Show All" };
            _btnShowAll.Click += OnShowAllDiagnostics;

            _btnClearSelection = new Button { Left = 435, Top = 22, Width = 80, Height = 26, Text = "Clear" };
            _btnClearSelection.Click += OnClearSelection;

            _lblDiagStatus = new Label { Left = 10, Top = 55, Width = 530, Height = 30, Text = "Part opens automatically. Use buttons to highlight geometry." };

            _grpTubeDiag.Controls.AddRange(new Control[] { _btnShowCutLength, _btnShowHoles, _btnShowBoundary, _btnShowProfile, _btnShowAll, _btnClearSelection, _lblDiagStatus });

            // Separator
            var separator2 = new Label { Left = 20, Top = 415, Width = 550, Height = 2, BorderStyle = BorderStyle.Fixed3D };

            // Action buttons
            _btnRetry = new Button { Left = 20, Top = 425, Width = 130, Height = 35, Text = "Retry && Validate", BackColor = Color.LightGreen };
            _btnRetry.Click += OnRetry;

            _btnSkip = new Button { Left = 160, Top = 425, Width = 90, Height = 35, Text = "Skip" };
            _btnSkip.Click += OnSkip;

            _btnFinish = new Button { Left = 260, Top = 425, Width = 90, Height = 35, Text = "Finish" };
            _btnFinish.Click += OnFinish;

            _btnCancel = new Button { Left = 360, Top = 425, Width = 90, Height = 35, Text = "Cancel" };
            _btnCancel.Click += OnCancel;

            // Progress section
            _progressBar = new ProgressBar { Left = 20, Top = 475, Width = 400, Height = 22 };
            _lblFixedCount = new Label { Left = 430, Top = 478, Width = 140, Height = 20 };

            _lblStatus = new Label { Left = 20, Top = 505, Width = 550, Height = 40, ForeColor = Color.DarkBlue };

            // Add all controls
            Controls.AddRange(new Control[]
            {
                _lblProgress, _btnPrevious, _btnNext, separator1,
                _lblFileName, _lblPath, _lblConfig, _lblCategory,
                lblError, _txtError, lblSuggestions, _lstSuggestions,
                _grpTubeDiag, separator2,
                _btnRetry, _btnSkip, _btnFinish, _btnCancel,
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
            var suggestions = GetSuggestions(item.Category, item.ProblemDescription);
            foreach (var s in suggestions)
            {
                _lstSuggestions.Items.Add(s);
            }
        }

        private List<string> GetSuggestions(ProblemPartManager.ProblemCategory category, string error)
        {
            var list = new List<string>();
            var errLower = (error ?? "").ToLowerInvariant();

            switch (category)
            {
                case ProblemPartManager.ProblemCategory.GeometryValidation:
                    if (errLower.Contains("multi-body"))
                    {
                        list.Add("Insert > Features > Combine to merge bodies");
                        list.Add("Delete unwanted bodies from FeatureManager");
                        list.Add("Check if bodies came from imported geometry");
                    }
                    else if (errLower.Contains("surface") || errLower.Contains("no solid"))
                    {
                        list.Add("Part may be surface-only - needs solid geometry");
                        list.Add("Use Insert > Surface > Thicken to create solid");
                        list.Add("Check if part file is corrupted or empty");
                    }
                    else
                    {
                        list.Add("Review part geometry in SolidWorks");
                        list.Add("Check for invalid features or geometry errors");
                    }
                    break;

                case ProblemPartManager.ProblemCategory.MaterialMissing:
                    list.Add("Right-click part in FeatureManager > Material > Edit Material");
                    list.Add("Or: Click the material icon in FeatureManager tree");
                    list.Add("Assign appropriate material (e.g., AISI 304, Plain Carbon Steel)");
                    break;

                case ProblemPartManager.ProblemCategory.SheetMetalConversion:
                    list.Add("Check that part has uniform wall thickness");
                    list.Add("Verify parallel faces exist for sheet metal conversion");
                    list.Add("Ensure no complex features block conversion");
                    list.Add("Try using Insert > Sheet Metal > Insert Bends manually");
                    break;

                case ProblemPartManager.ProblemCategory.FileAccess:
                    list.Add("Verify file exists at the shown path");
                    list.Add("Check file is not read-only or locked");
                    list.Add("Ensure file is not open in another application");
                    break;

                case ProblemPartManager.ProblemCategory.Suppressed:
                    list.Add("Component is suppressed in assembly");
                    list.Add("Right-click component > Set to Resolved");
                    list.Add("Check if component was intentionally excluded");
                    break;

                case ProblemPartManager.ProblemCategory.Lightweight:
                    list.Add("Component is in Lightweight mode");
                    list.Add("Right-click > Set to Resolved");
                    list.Add("Or: Tools > Options > Assemblies > uncheck Lightweight");
                    break;

                case ProblemPartManager.ProblemCategory.ThicknessExtraction:
                    list.Add("Could not determine sheet metal thickness");
                    list.Add("Check that part has uniform wall thickness");
                    list.Add("May need manual thickness specification");
                    break;

                default:
                    list.Add("Review part in SolidWorks");
                    list.Add("Check error message for specific guidance");
                    break;
            }

            return list;
        }

        private void UpdateNavigationButtons()
        {
            _btnPrevious.Enabled = _currentIndex > 0;
            _btnNext.Enabled = _currentIndex < _problems.Count - 1;
        }

        private void UpdateProgress()
        {
            int total = _problems.Count;
            int remaining = _problems.Count(p => !_fixedProblems.Contains(p));
            int fixedCount = _fixedProblems.Count;

            _progressBar.Maximum = total > 0 ? total : 1;
            _progressBar.Value = Math.Min(fixedCount, _progressBar.Maximum);

            _lblFixedCount.Text = $"{fixedCount} / {total} fixed";

            _btnFinish.Text = remaining == 0 ? "Finish" : $"Finish ({_initialGoodCount + fixedCount} ready)";
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

        /// <summary>
        /// Saves and closes the currently-open document, if any.
        /// </summary>
        private void SaveAndCloseCurrentPart()
        {
            if (_currentDoc == null) return;

            try
            {
                SolidWorksApiWrapper.SaveDocument(_currentDoc);
            }
            catch (Exception ex)
            {
                NM.Core.ErrorHandler.DebugLog($"[Wizard] Save failed: {ex.Message}");
            }

            try
            {
                SolidWorksApiWrapper.CloseDocument(_swApp, _currentDoc);
            }
            catch (Exception ex)
            {
                NM.Core.ErrorHandler.DebugLog($"[Wizard] Close failed: {ex.Message}");
            }

            _currentDoc = null;
            _tubeDiagnostics = null;
        }

        /// <summary>
        /// Opens the part for the current problem. Called automatically on navigation.
        /// </summary>
        private void OpenCurrentPart()
        {
            if (_swApp == null) return;
            if (_problems.Count == 0 || _currentIndex >= _problems.Count) return;

            var item = _problems[_currentIndex];
            if (string.IsNullOrEmpty(item.FilePath)) return;

            try
            {
                int errs = 0, warns = 0;
                int docType = GuessDocType(item.FilePath);

                _currentDoc = _swApp.OpenDoc6(item.FilePath, docType, 0, item.Configuration ?? "", ref errs, ref warns) as IModelDoc2;
                if (_currentDoc == null || errs != 0)
                {
                    _lblStatus.Text = $"Failed to open: error {errs}";
                    return;
                }

                // Activate to bring to front
                int activateErr = 0;
                _swApp.ActivateDoc3(_currentDoc.GetTitle(), true, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref activateErr);

                _lblDiagStatus.Text = "Part is open. Use buttons above to highlight geometry for debugging.";
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Open failed: {ex.Message}";
            }
        }

        private void OnRetry(object sender, EventArgs e)
        {
            if (_swApp == null || _problems.Count == 0) return;

            var item = _problems[_currentIndex];
            _lblStatus.Text = $"Validating {item.DisplayName}...";
            Application.DoEvents();

            try
            {
                // Try to use already-open doc, or open it
                var model = _currentDoc;
                if (model == null || !string.Equals(model.GetPathName(), item.FilePath, StringComparison.OrdinalIgnoreCase))
                {
                    int errs = 0, warns = 0;
                    model = _swApp.OpenDoc6(item.FilePath, GuessDocType(item.FilePath), 0, item.Configuration ?? "", ref errs, ref warns) as IModelDoc2;
                }

                if (model == null)
                {
                    _lblStatus.Text = "Could not open part for validation.";
                    return;
                }

                // Revalidate
                var swInfo = new SwModelInfo(item.FilePath) { Configuration = item.Configuration ?? "" };
                var validator = new PartValidationAdapter();
                var vr = validator.Validate(swInfo, model);

                if (vr.Success)
                {
                    // Fixed!
                    _fixedProblems.Add(item);
                    ProblemPartManager.Instance.RemoveResolvedPart(item);
                    _lblStatus.Text = $"FIXED: {item.DisplayName}";
                    _txtError.BackColor = Color.LightGreen;
                    _txtError.Text = "Validation passed!";

                    UpdateProgress();

                    // Auto-advance after short delay if there are more
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
                        _lblStatus.Text += " - All problems reviewed!";
                    }
                }
                else
                {
                    // Still failing
                    item.ProblemDescription = vr.Summary;
                    item.RetryCount++;
                    _txtError.Text = vr.Summary;
                    _txtError.BackColor = Color.MistyRose;
                    LoadSuggestions(item);
                    _lblStatus.Text = $"Still failing: {vr.Summary}";
                }
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Validation error: {ex.Message}";
            }
        }

        private void OnSkip(object sender, EventArgs e)
        {
            MarkCurrentPartAsSkipped();

            if (_currentIndex < _problems.Count - 1)
            {
                _currentIndex++;
                LoadCurrentProblem(); // saves/closes skipped part, opens next
            }
            else
            {
                SaveAndCloseCurrentPart();
                _lblStatus.Text = "No more problems. Click Finish to continue.";
            }
        }

        /// <summary>
        /// Writes NM_SkippedReason custom property to the current part.
        /// </summary>
        private void MarkCurrentPartAsSkipped()
        {
            if (_currentDoc == null) return;
            if (_currentIndex < 0 || _currentIndex >= _problems.Count) return;

            var item = _problems[_currentIndex];

            try
            {
                string reason = item.ProblemDescription ?? "Unknown";
                string value = $"{reason} [Skipped {DateTime.Now:yyyy-MM-dd HH:mm}]";

                SolidWorksApiWrapper.AddCustomProperty(
                    _currentDoc,
                    "NM_SkippedReason",
                    swCustomInfoType_e.swCustomInfoText,
                    value,
                    ""); // file-level property

                NM.Core.ErrorHandler.DebugLog($"[Wizard] Marked skipped: {item.DisplayName} - {reason}");
            }
            catch (Exception ex)
            {
                NM.Core.ErrorHandler.DebugLog($"[Wizard] Failed to mark skipped: {ex.Message}");
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

        private static int GuessDocType(string path)
        {
            var ext = (Path.GetExtension(path) ?? "").ToLowerInvariant();
            if (ext == ".sldprt") return (int)swDocumentTypes_e.swDocPART;
            if (ext == ".sldasm") return (int)swDocumentTypes_e.swDocASSEMBLY;
            if (ext == ".slddrw") return (int)swDocumentTypes_e.swDocDRAWING;
            return (int)swDocumentTypes_e.swDocPART;
        }

        #region Tube Diagnostics

        private bool EnsurePartOpenAndExtractDiagnostics()
        {
            if (_swApp == null)
            {
                _lblDiagStatus.Text = "SolidWorks instance not available.";
                return false;
            }

            if (_currentDoc == null)
            {
                _lblDiagStatus.Text = "Part is not open. Navigate to a problem to auto-open it.";
                return false;
            }

            if (_tubeDiagnostics == null)
            {
                try
                {
                    var extractor = new TubeGeometryExtractor(_swApp);
                    var (profile, diagnostics) = extractor.ExtractWithDiagnostics(_currentDoc);
                    _tubeDiagnostics = diagnostics;

                    if (profile != null)
                    {
                        _lblDiagStatus.Text = $"Profile: {profile.Shape}, OD={profile.OuterDiameterMeters * 39.3701:F3}in | {diagnostics.GetSummary()}";
                    }
                    else
                    {
                        _lblDiagStatus.Text = "No tube profile detected. " + diagnostics.GetSummary();
                    }
                }
                catch (Exception ex)
                {
                    _lblDiagStatus.Text = "Extraction failed: " + ex.Message;
                    return false;
                }
            }

            return true;
        }

        private void OnShowCutLengthEdges(object sender, EventArgs e)
        {
            if (!EnsurePartOpenAndExtractDiagnostics()) return;
            try
            {
                var extractor = new TubeGeometryExtractor(_swApp);
                extractor.SelectCutLengthEdges(_currentDoc, _tubeDiagnostics, 0x00FF00);
                _lblDiagStatus.Text = $"Selected {_tubeDiagnostics.CutLengthEdges.Count} cut length edges (green).";
            }
            catch (Exception ex) { _lblDiagStatus.Text = "Select failed: " + ex.Message; }
        }

        private void OnShowHoleEdges(object sender, EventArgs e)
        {
            if (!EnsurePartOpenAndExtractDiagnostics()) return;
            try
            {
                var extractor = new TubeGeometryExtractor(_swApp);
                extractor.SelectHoleEdges(_currentDoc, _tubeDiagnostics);
                _lblDiagStatus.Text = $"Selected {_tubeDiagnostics.HoleEdges.Count} hole edges (red).";
            }
            catch (Exception ex) { _lblDiagStatus.Text = "Select failed: " + ex.Message; }
        }

        private void OnShowBoundaryEdges(object sender, EventArgs e)
        {
            if (!EnsurePartOpenAndExtractDiagnostics()) return;
            try
            {
                var extractor = new TubeGeometryExtractor(_swApp);
                extractor.SelectBoundaryEdges(_currentDoc, _tubeDiagnostics);
                _lblDiagStatus.Text = $"Selected {_tubeDiagnostics.BoundaryEdges.Count} boundary edges (blue).";
            }
            catch (Exception ex) { _lblDiagStatus.Text = "Select failed: " + ex.Message; }
        }

        private void OnShowProfileFaces(object sender, EventArgs e)
        {
            if (!EnsurePartOpenAndExtractDiagnostics()) return;
            try
            {
                var extractor = new TubeGeometryExtractor(_swApp);
                extractor.SelectProfileFaces(_currentDoc, _tubeDiagnostics);
                _lblDiagStatus.Text = $"Selected {_tubeDiagnostics.ProfileFaces.Count} profile faces (cyan).";
            }
            catch (Exception ex) { _lblDiagStatus.Text = "Select failed: " + ex.Message; }
        }

        private void OnShowAllDiagnostics(object sender, EventArgs e)
        {
            if (!EnsurePartOpenAndExtractDiagnostics()) return;
            try
            {
                var extractor = new TubeGeometryExtractor(_swApp);
                extractor.SelectAllDiagnostics(_currentDoc, _tubeDiagnostics);
                _lblDiagStatus.Text = _tubeDiagnostics.GetSummary();
            }
            catch (Exception ex) { _lblDiagStatus.Text = "Select failed: " + ex.Message; }
        }

        private void OnClearSelection(object sender, EventArgs e)
        {
            if (_currentDoc == null)
            {
                _lblDiagStatus.Text = "No model open.";
                return;
            }
            try
            {
                _currentDoc.ClearSelection2(true);
                _lblDiagStatus.Text = "Selection cleared.";
            }
            catch (Exception ex) { _lblDiagStatus.Text = "Clear failed: " + ex.Message; }
        }

        #endregion
    }
}
