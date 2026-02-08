using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using NM.Core.Processing;
using NM.Core.ProblemParts;
using NM.SwAddin.Geometry;
using NM.SwAddin.Pipeline;
using NM.SwAddin.Validation;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.UI
{
    public sealed class ProblemReviewDialog : Form
    {
        private readonly ProblemPartManager.ProblemItem _item;
        private readonly ISldWorks _swApp;

        private Label _lblFile;
        private Label _lblPath;
        private Label _lblConfig;
        private Label _lblCategory;
        private Label _lblAttempts;
        private TextBox _txtReason;
        private ListBox _lstSuggestions;
        private Button _btnOpen;
        private Button _btnMark;
        private Button _btnConfirmFix;
        private Button _btnClose;

        // Tube diagnostic controls
        private GroupBox _grpTubeDiagnostics;
        private Button _btnShowCutLength;
        private Button _btnShowHoles;
        private Button _btnShowBoundary;
        private Button _btnShowProfile;
        private Button _btnShowAll;
        private Button _btnClearSelection;
        private Label _lblDiagStatus;

        // Diagnostic state
        private TubeDiagnosticInfo _tubeDiagnostics;
        private IModelDoc2 _openedModel;

        public bool Resolved { get; private set; }

        public ProblemReviewDialog(ProblemPartManager.ProblemItem item, ISldWorks swApp = null)
        {
            _item = item ?? throw new ArgumentNullException(nameof(item));
            _swApp = swApp;

            Text = "Review Problem Part";
            Width = 720;
            Height = 640;
            StartPosition = FormStartPosition.CenterParent;

            BuildUi();
            LoadData();
        }

        private void BuildUi()
        {
            _lblFile = new Label { Left = 10, Top = 10, Width = 680, Text = "File: " };
            _lblPath = new Label { Left = 10, Top = 30, Width = 680, Text = "Path: " };
            _lblConfig = new Label { Left = 10, Top = 50, Width = 680, Text = "Config: " };
            _lblCategory = new Label { Left = 10, Top = 70, Width = 680, Text = "Category: " };
            _lblAttempts = new Label { Left = 10, Top = 90, Width = 680, Text = "Attempts: " };

            _txtReason = new TextBox { Left = 10, Top = 120, Width = 680, Height = 80, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
            _lstSuggestions = new ListBox { Left = 10, Top = 210, Width = 680, Height = 140 };

            // Tube Diagnostics GroupBox
            _grpTubeDiagnostics = new GroupBox { Left = 10, Top = 360, Width = 680, Height = 100, Text = "Tube Diagnostics (requires part open in SolidWorks)" };

            _btnShowCutLength = new Button { Left = 10, Top = 20, Width = 100, Height = 25, Text = "Cut Length" };
            _btnShowCutLength.Click += OnShowCutLengthEdges;

            _btnShowHoles = new Button { Left = 115, Top = 20, Width = 100, Height = 25, Text = "Hole Edges" };
            _btnShowHoles.Click += OnShowHoleEdges;

            _btnShowBoundary = new Button { Left = 220, Top = 20, Width = 100, Height = 25, Text = "Boundary" };
            _btnShowBoundary.Click += OnShowBoundaryEdges;

            _btnShowProfile = new Button { Left = 325, Top = 20, Width = 100, Height = 25, Text = "Profile Faces" };
            _btnShowProfile.Click += OnShowProfileFaces;

            _btnShowAll = new Button { Left = 430, Top = 20, Width = 100, Height = 25, Text = "Show All" };
            _btnShowAll.Click += OnShowAllDiagnostics;

            _btnClearSelection = new Button { Left = 535, Top = 20, Width = 100, Height = 25, Text = "Clear" };
            _btnClearSelection.Click += OnClearSelection;

            _lblDiagStatus = new Label { Left = 10, Top = 50, Width = 660, Height = 40, Text = "Open part first, then click a button to highlight edges/faces." };

            _grpTubeDiagnostics.Controls.AddRange(new Control[] { _btnShowCutLength, _btnShowHoles, _btnShowBoundary, _btnShowProfile, _btnShowAll, _btnClearSelection, _lblDiagStatus });

            _btnOpen = new Button { Left = 10, Top = 520, Width = 160, Text = "Open in SolidWorks" };
            _btnOpen.Click += OnOpenInSolidWorks;

            _btnMark = new Button { Left = 180, Top = 520, Width = 160, Text = "Mark Reviewed" };
            _btnMark.Click += (s, e) => { DialogResult = DialogResult.OK; Close(); };

            _btnConfirmFix = new Button { Left = 350, Top = 520, Width = 160, Text = "Confirm Fix" };
            _btnConfirmFix.Click += OnConfirmFix;

            _btnClose = new Button { Left = 520, Top = 520, Width = 100, Text = "Close" };
            _btnClose.Click += (s, e) => Close();

            Controls.AddRange(new Control[] { _lblFile, _lblPath, _lblConfig, _lblCategory, _lblAttempts, _txtReason, _lstSuggestions, _grpTubeDiagnostics, _btnOpen, _btnMark, _btnConfirmFix, _btnClose });
        }

        private void LoadData()
        {
            _lblFile.Text = "File: " + (_item.DisplayName ?? "");
            _lblPath.Text = "Path: " + (_item.FilePath ?? "");
            _lblConfig.Text = "Config: " + (_item.Configuration ?? "");
            _lblCategory.Text = "Category: " + _item.Category.ToString();
            _lblAttempts.Text = "Attempts: " + _item.RetryCount.ToString();
            _txtReason.Text = _item.ProblemDescription ?? string.Empty;
            LoadSuggestions();
        }

        private void LoadSuggestions()
        {
            var suggestions = ProblemSuggestionProvider.GetSuggestions(_item.Category, _item.ProblemDescription);
            _lstSuggestions.Items.Clear();
            _lstSuggestions.Items.AddRange(suggestions.Cast<object>().ToArray());
        }

        private void OnOpenInSolidWorks(object sender, EventArgs e)
        {
            try
            {
                var sw = _swApp;
                if (sw == null)
                {
                    MessageBox.Show(this, "SolidWorks instance not available.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                int errs = 0, warns = 0;
                int docType = SwDocumentHelper.GuessDocType(_item.FilePath);

                // Open without Silent flag so it displays properly
                var model = sw.OpenDoc6(_item.FilePath, docType, 0, _item.Configuration ?? "", ref errs, ref warns) as IModelDoc2;
                if (model == null || errs != 0)
                {
                    MessageBox.Show(this, "Open failed (" + errs + ")", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Activate the document to bring it to front
                int activateErr = 0;
                sw.ActivateDoc3(model.GetTitle(), true, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref activateErr);

                // Store reference for tube diagnostics
                _openedModel = model;

                // Update button text to show it's open
                _btnOpen.Text = "Part Opened ✓";
                _btnOpen.Enabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Open failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnConfirmFix(object sender, EventArgs e)
        {
            if (_swApp == null)
            {
                MessageBox.Show(this, "SolidWorks instance not available.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                int errs = 0, warns = 0;
                int docType = SwDocumentHelper.GuessDocType(_item.FilePath);
                if (docType != (int)swDocumentTypes_e.swDocPART)
                {
                    MessageBox.Show(this, "Confirm Fix supports parts only. Open individual part and retry.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var model = _swApp.OpenDoc6(_item.FilePath, docType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, _item.Configuration, ref errs, ref warns) as IModelDoc2;
                if (model == null || errs != 0)
                {
                    MessageBox.Show(this, "Open failed (" + errs + ")", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Revalidate
                var validator = new PartValidationAdapter();
                var swInfo = new NM.Core.Models.SwModelInfo(_item.FilePath) { Configuration = _item.Configuration ?? string.Empty };
                var vr = validator.Validate(swInfo, model);
                if (!vr.Success)
                {
                    MessageBox.Show(this, "Still failing: " + (vr.Summary ?? "Unknown"), "Not Fixed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Process via single-part pipeline
                var res = MainRunner.RunSinglePart(_swApp, model, options: null);
                if (!res.Success)
                {
                    MessageBox.Show(this, "Processing failed: " + (res.Message ?? "Unknown"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Remove from problem list and mark resolved
                ProblemPartManager.Instance.RemoveResolvedPart(_item);
                Resolved = true;
                MessageBox.Show(this, "Fixed and processed.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Confirm Fix failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Optional: keep document open for user; else close
            }
        }

        #region Tube Diagnostic Event Handlers

        private bool EnsurePartOpenAndExtractDiagnostics()
        {
            if (_swApp == null)
            {
                _lblDiagStatus.Text = "SolidWorks instance not available.";
                return false;
            }

            // Check if model is already open
            if (_openedModel == null)
            {
                // Try to get active document
                _openedModel = _swApp.ActiveDoc as IModelDoc2;

                // If not the right file, try to open it
                if (_openedModel == null || !string.Equals(_openedModel.GetPathName(), _item.FilePath, StringComparison.OrdinalIgnoreCase))
                {
                    int errs = 0, warns = 0;
                    int docType = SwDocumentHelper.GuessDocType(_item.FilePath);
                    if (docType != (int)swDocumentTypes_e.swDocPART)
                    {
                        _lblDiagStatus.Text = "Tube diagnostics only work on parts.";
                        return false;
                    }

                    _openedModel = _swApp.OpenDoc6(_item.FilePath, docType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, _item.Configuration, ref errs, ref warns) as IModelDoc2;
                    if (_openedModel == null || errs != 0)
                    {
                        _lblDiagStatus.Text = "Failed to open part: error " + errs;
                        return false;
                    }
                }
            }

            // Extract diagnostics if not already done
            if (_tubeDiagnostics == null)
            {
                try
                {
                    var extractor = new TubeGeometryExtractor(_swApp);
                    var (profile, diagnostics) = extractor.ExtractWithDiagnostics(_openedModel);
                    _tubeDiagnostics = diagnostics;

                    if (profile != null)
                    {
                        _lblDiagStatus.Text = $"Profile: {profile.Shape}, OD={profile.OuterDiameterMeters:F3}m, CutLen={profile.CutLengthMeters:F3}m | " + diagnostics.GetSummary();
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
                extractor.SelectCutLengthEdges(_openedModel, _tubeDiagnostics);
                _lblDiagStatus.Text = $"Selected {_tubeDiagnostics.CutLengthEdges.Count} cut length edges (green).";
            }
            catch (Exception ex)
            {
                _lblDiagStatus.Text = "Select failed: " + ex.Message;
            }
        }

        private void OnShowHoleEdges(object sender, EventArgs e)
        {
            if (!EnsurePartOpenAndExtractDiagnostics()) return;

            try
            {
                var extractor = new TubeGeometryExtractor(_swApp);
                extractor.SelectHoleEdges(_openedModel, _tubeDiagnostics);
                _lblDiagStatus.Text = $"Selected {_tubeDiagnostics.HoleEdges.Count} hole edges (red).";
            }
            catch (Exception ex)
            {
                _lblDiagStatus.Text = "Select failed: " + ex.Message;
            }
        }

        private void OnShowBoundaryEdges(object sender, EventArgs e)
        {
            if (!EnsurePartOpenAndExtractDiagnostics()) return;

            try
            {
                var extractor = new TubeGeometryExtractor(_swApp);
                extractor.SelectBoundaryEdges(_openedModel, _tubeDiagnostics);
                _lblDiagStatus.Text = $"Selected {_tubeDiagnostics.BoundaryEdges.Count} boundary edges (blue).";
            }
            catch (Exception ex)
            {
                _lblDiagStatus.Text = "Select failed: " + ex.Message;
            }
        }

        private void OnShowProfileFaces(object sender, EventArgs e)
        {
            if (!EnsurePartOpenAndExtractDiagnostics()) return;

            try
            {
                var extractor = new TubeGeometryExtractor(_swApp);
                extractor.SelectProfileFaces(_openedModel, _tubeDiagnostics);
                _lblDiagStatus.Text = $"Selected {_tubeDiagnostics.ProfileFaces.Count} profile faces (cyan).";
            }
            catch (Exception ex)
            {
                _lblDiagStatus.Text = "Select failed: " + ex.Message;
            }
        }

        private void OnShowAllDiagnostics(object sender, EventArgs e)
        {
            if (!EnsurePartOpenAndExtractDiagnostics()) return;

            try
            {
                var extractor = new TubeGeometryExtractor(_swApp);
                extractor.SelectAllDiagnostics(_openedModel, _tubeDiagnostics);
                _lblDiagStatus.Text = _tubeDiagnostics.GetSummary();
            }
            catch (Exception ex)
            {
                _lblDiagStatus.Text = "Select failed: " + ex.Message;
            }
        }

        private void OnClearSelection(object sender, EventArgs e)
        {
            if (_openedModel == null)
            {
                _lblDiagStatus.Text = "No model open.";
                return;
            }

            try
            {
                _openedModel.ClearSelection2(true);
                _lblDiagStatus.Text = "Selection cleared.";
            }
            catch (Exception ex)
            {
                _lblDiagStatus.Text = "Clear failed: " + ex.Message;
            }
        }

        #endregion
    }
}
