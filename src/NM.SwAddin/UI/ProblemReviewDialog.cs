using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NM.Core.ProblemParts;
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

        public bool Resolved { get; private set; }

        public ProblemReviewDialog(ProblemPartManager.ProblemItem item, ISldWorks swApp = null)
        {
            _item = item ?? throw new ArgumentNullException(nameof(item));
            _swApp = swApp;

            Text = "Review Problem Part";
            Width = 720;
            Height = 560;
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
            _lstSuggestions = new ListBox { Left = 10, Top = 210, Width = 680, Height = 220 };

            _btnOpen = new Button { Left = 10, Top = 440, Width = 160, Text = "Open in SolidWorks" };
            _btnOpen.Click += OnOpenInSolidWorks;

            _btnMark = new Button { Left = 180, Top = 440, Width = 160, Text = "Mark Reviewed" };
            _btnMark.Click += (s, e) => { DialogResult = DialogResult.OK; Close(); };

            _btnConfirmFix = new Button { Left = 350, Top = 440, Width = 160, Text = "Confirm Fix" };
            _btnConfirmFix.Click += OnConfirmFix;

            _btnClose = new Button { Left = 520, Top = 440, Width = 100, Text = "Close" };
            _btnClose.Click += (s, e) => Close();

            Controls.AddRange(new Control[] { _lblFile, _lblPath, _lblConfig, _lblCategory, _lblAttempts, _txtReason, _lstSuggestions, _btnOpen, _btnMark, _btnConfirmFix, _btnClose });
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
            var suggestions = new List<string>();
            switch (_item.Category)
            {
                case ProblemPartManager.ProblemCategory.SheetMetalConversion:
                    suggestions.Add("� Check if part has uniform thickness");
                    suggestions.Add("� Verify parallel faces exist");
                    suggestions.Add("� Ensure no complex features prevent conversion");
                    break;
                case ProblemPartManager.ProblemCategory.MaterialMissing:
                    suggestions.Add("� Open part and assign material");
                    suggestions.Add("� Check material database is accessible");
                    break;
                case ProblemPartManager.ProblemCategory.FileAccess:
                    suggestions.Add("� Verify file exists at path");
                    suggestions.Add("� Check file is not read-only");
                    suggestions.Add("� Ensure file is not open in another session");
                    break;
                case ProblemPartManager.ProblemCategory.Lightweight:
                    suggestions.Add("� Open assembly and resolve component");
                    suggestions.Add("� Check if Large Assembly Mode is enabled");
                    break;
                default:
                    suggestions.Add("� Review part manually and adjust settings");
                    break;
            }
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
                int docType = GuessDocType(_item.FilePath);
                var model = sw.OpenDoc6(_item.FilePath, docType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, _item.Configuration, ref errs, ref warns);
                if (model == null || errs != 0)
                {
                    MessageBox.Show(this, "Open failed (" + errs + ")", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show(this, "Opened in SolidWorks.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
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
                int docType = GuessDocType(_item.FilePath);
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

        private static int GuessDocType(string path)
        {
            try
            {
                var ext = (Path.GetExtension(path) ?? string.Empty).ToLowerInvariant();
                if (ext == ".sldprt") return (int)swDocumentTypes_e.swDocPART;
                if (ext == ".sldasm") return (int)swDocumentTypes_e.swDocASSEMBLY;
                if (ext == ".slddrw") return (int)swDocumentTypes_e.swDocDRAWING;
            }
            catch { }
            return (int)swDocumentTypes_e.swDocPART;
        }
    }
}
