using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using NM.Core.Reconciliation.Models;
using NM.Core.Writeback.Models;

namespace NM.SwAddin.UI
{
    /// <summary>
    /// WinForms wizard for reviewing AI-extracted property suggestions before writing to the model.
    /// Shows conflicts, property suggestions (with checkboxes), and file rename options.
    /// Returns the list of approved suggestions when the user clicks Apply.
    /// </summary>
    public sealed class PropertyReviewWizard : Form
    {
        // Header
        private Label _lblFile;
        private Label _lblDrawing;
        private Label _lblSummary;

        // Conflicts section
        private GroupBox _grpConflicts;
        private DataGridView _gridConflicts;

        // Suggestions grid
        private GroupBox _grpSuggestions;
        private DataGridView _gridSuggestions;

        // Rename section
        private GroupBox _grpRename;
        private CheckBox _chkRename;
        private Label _lblRenameDetails;

        // Action buttons
        private Button _btnSelectAll;
        private Button _btnDeselectAll;
        private Button _btnApply;
        private Button _btnCancel;

        // Status
        private Label _lblStatus;

        // Data
        private readonly List<PropertySuggestion> _suggestions;
        private readonly List<DataConflict> _conflicts;
        private readonly RenameSuggestion _rename;
        private readonly FileRenameValidation _renameValidation;
        private readonly string _fileName;
        private readonly string _drawingName;
        private readonly string _analysisSummary;

        /// <summary>Suggestions approved by the user (populated on Apply).</summary>
        public List<PropertySuggestion> ApprovedSuggestions { get; private set; }

        /// <summary>Conflict resolutions chosen by the user.</summary>
        public Dictionary<string, string> ConflictResolutions { get; private set; }

        /// <summary>Whether the user approved the file rename.</summary>
        public bool RenameApproved { get; private set; }

        public PropertyReviewWizard(
            List<PropertySuggestion> suggestions,
            List<DataConflict> conflicts = null,
            RenameSuggestion rename = null,
            FileRenameValidation renameValidation = null,
            string fileName = null,
            string drawingName = null,
            string analysisSummary = null)
        {
            _suggestions = suggestions ?? new List<PropertySuggestion>();
            _conflicts = conflicts ?? new List<DataConflict>();
            _rename = rename;
            _renameValidation = renameValidation;
            _fileName = fileName ?? "";
            _drawingName = drawingName ?? "";
            _analysisSummary = analysisSummary ?? "";

            Text = "AI Drawing Analysis - Property Suggestions";
            Width = 860;
            Height = 720;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = true;
            FormBorderStyle = FormBorderStyle.Sizable;

            BuildUi();
            LoadData();
        }

        private void BuildUi()
        {
            // Header
            _lblFile = new Label { Left = 12, Top = 10, Width = 820, Font = new Font(Font.FontFamily, 9f, FontStyle.Bold) };
            _lblDrawing = new Label { Left = 12, Top = 30, Width = 820 };
            _lblSummary = new Label { Left = 12, Top = 50, Width = 820, ForeColor = Color.DarkBlue };

            // Conflicts
            _grpConflicts = new GroupBox { Left = 12, Top = 75, Width = 820, Height = 120, Text = "Conflicts (require resolution)", Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            _gridConflicts = new DataGridView
            {
                Left = 8, Top = 18, Width = 800, Height = 90,
                AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ReadOnly = false, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _gridConflicts.Columns.Add("Field", "Field");
            _gridConflicts.Columns.Add("ModelValue", "3D Model Value");
            _gridConflicts.Columns.Add("DrawingValue", "Drawing Value");
            var resolveCol = new DataGridViewComboBoxColumn
            {
                Name = "Resolution",
                HeaderText = "Use",
                Items = { "Model", "Drawing", "Skip" },
                Width = 100
            };
            _gridConflicts.Columns.Add(resolveCol);
            _gridConflicts.Columns.Add("Reason", "Reason");

            // Make only Resolution column editable
            _gridConflicts.Columns["Field"].ReadOnly = true;
            _gridConflicts.Columns["ModelValue"].ReadOnly = true;
            _gridConflicts.Columns["DrawingValue"].ReadOnly = true;
            _gridConflicts.Columns["Reason"].ReadOnly = true;

            _grpConflicts.Controls.Add(_gridConflicts);

            // Suggestions grid
            _grpSuggestions = new GroupBox { Left = 12, Top = 200, Width = 820, Height = 320, Text = "Property Suggestions", Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom };
            _gridSuggestions = new DataGridView
            {
                Left = 8, Top = 18, Width = 800, Height = 290,
                AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ReadOnly = false, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };

            var chkCol = new DataGridViewCheckBoxColumn { Name = "Apply", HeaderText = "", Width = 30 };
            _gridSuggestions.Columns.Add(chkCol);
            _gridSuggestions.Columns.Add("Property", "Property");
            _gridSuggestions.Columns.Add("CurrentValue", "Current");
            _gridSuggestions.Columns.Add("SuggestedValue", "Suggested");
            _gridSuggestions.Columns.Add("Source", "Source");
            _gridSuggestions.Columns.Add("Confidence", "Conf");
            _gridSuggestions.Columns.Add("ReasonCol", "Reason");

            // Only the checkbox is editable
            _gridSuggestions.Columns["Property"].ReadOnly = true;
            _gridSuggestions.Columns["CurrentValue"].ReadOnly = true;
            _gridSuggestions.Columns["SuggestedValue"].ReadOnly = true;
            _gridSuggestions.Columns["Source"].ReadOnly = true;
            _gridSuggestions.Columns["Confidence"].ReadOnly = true;
            _gridSuggestions.Columns["ReasonCol"].ReadOnly = true;

            // Column widths
            _gridSuggestions.Columns["Confidence"].Width = 45;
            _gridSuggestions.Columns["Source"].Width = 70;

            _grpSuggestions.Controls.Add(_gridSuggestions);

            // Rename section
            _grpRename = new GroupBox { Left = 12, Top = 528, Width = 820, Height = 55, Text = "File Rename", Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right };
            _chkRename = new CheckBox { Left = 12, Top = 20, Width = 20, Height = 20 };
            _lblRenameDetails = new Label { Left = 36, Top = 22, Width = 770, AutoSize = false };
            _grpRename.Controls.AddRange(new Control[] { _chkRename, _lblRenameDetails });

            // Buttons
            int btnY = 592;
            _btnSelectAll = new Button { Left = 12, Top = btnY, Width = 100, Height = 28, Text = "Select All", Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            _btnSelectAll.Click += OnSelectAll;

            _btnDeselectAll = new Button { Left = 120, Top = btnY, Width = 100, Height = 28, Text = "Deselect All", Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            _btnDeselectAll.Click += OnDeselectAll;

            _btnApply = new Button { Left = 610, Top = btnY, Width = 110, Height = 28, Text = "Apply Selected", Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            _btnApply.Click += OnApply;

            _btnCancel = new Button { Left = 728, Top = btnY, Width = 100, Height = 28, Text = "Cancel", Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            _btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            _lblStatus = new Label { Left = 230, Top = btnY + 5, Width = 370, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.DarkGreen, Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right };

            Controls.AddRange(new Control[]
            {
                _lblFile, _lblDrawing, _lblSummary,
                _grpConflicts, _grpSuggestions, _grpRename,
                _btnSelectAll, _btnDeselectAll, _btnApply, _btnCancel, _lblStatus
            });
        }

        private void LoadData()
        {
            _lblFile.Text = $"File: {_fileName}";
            _lblDrawing.Text = $"Drawing: {_drawingName}";
            _lblSummary.Text = _analysisSummary;

            // Load conflicts
            if (_conflicts.Count == 0)
            {
                _grpConflicts.Text = "Conflicts (none)";
                _grpConflicts.Height = 40;
                // Shift suggestions up
                _grpSuggestions.Top = 120;
                _grpSuggestions.Height += 80;
            }
            else
            {
                _grpConflicts.Text = $"Conflicts ({_conflicts.Count})";
                foreach (var c in _conflicts)
                {
                    string defaultResolution = c.Recommendation == ConflictResolution.UseModel ? "Model"
                        : c.Recommendation == ConflictResolution.UseDrawing ? "Drawing"
                        : "Skip";

                    _gridConflicts.Rows.Add(
                        c.Field,
                        c.ModelValue ?? "(empty)",
                        c.DrawingValue ?? "(empty)",
                        defaultResolution,
                        c.Reason ?? ""
                    );
                }
            }

            // Load suggestions — pre-check high-confidence items
            foreach (var s in _suggestions)
            {
                bool autoCheck = s.Confidence >= 0.80;
                int idx = _gridSuggestions.Rows.Add(
                    autoCheck,
                    s.PropertyName,
                    s.CurrentValue ?? "(empty)",
                    s.Value,
                    s.Source.ToString().Replace("Drawing", "Dwg "),
                    s.Confidence.ToString("0.00"),
                    s.Reason ?? ""
                );

                // Color-code: green for gap fill, yellow for override
                if (s.IsGapFill)
                    _gridSuggestions.Rows[idx].DefaultCellStyle.BackColor = Color.FromArgb(230, 255, 230);
                else if (s.IsOverride)
                    _gridSuggestions.Rows[idx].DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 210);
            }

            // Rename
            if (_rename == null)
            {
                _grpRename.Visible = false;
                _grpSuggestions.Height += 55;
            }
            else
            {
                string renameText = $"{System.IO.Path.GetFileName(_rename.OldPath)} → {System.IO.Path.GetFileName(_rename.NewPath)}";
                if (_renameValidation != null && !_renameValidation.IsValid)
                {
                    renameText += $"  [INVALID: {_renameValidation.ErrorMessage}]";
                    _chkRename.Enabled = false;
                    _lblRenameDetails.ForeColor = Color.Red;
                }
                else if (_renameValidation != null && _renameValidation.HasAssemblyReferences)
                {
                    renameText += $"  [{_renameValidation.AffectedAssemblyCount} assembly ref(s) will need updating]";
                    _lblRenameDetails.ForeColor = Color.DarkOrange;
                }

                _lblRenameDetails.Text = renameText;
                _chkRename.Checked = false; // Default to unchecked — rename is always explicit
            }

            UpdateStatus();
        }

        private void OnSelectAll(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in _gridSuggestions.Rows)
                row.Cells["Apply"].Value = true;
            UpdateStatus();
        }

        private void OnDeselectAll(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in _gridSuggestions.Rows)
                row.Cells["Apply"].Value = false;
            UpdateStatus();
        }

        private void OnApply(object sender, EventArgs e)
        {
            ApprovedSuggestions = new List<PropertySuggestion>();
            ConflictResolutions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // Collect approved suggestions
            for (int i = 0; i < _gridSuggestions.Rows.Count; i++)
            {
                var row = _gridSuggestions.Rows[i];
                bool isChecked = row.Cells["Apply"].Value is true;
                if (isChecked && i < _suggestions.Count)
                {
                    ApprovedSuggestions.Add(_suggestions[i]);
                }
            }

            // Collect conflict resolutions
            for (int i = 0; i < _gridConflicts.Rows.Count; i++)
            {
                var row = _gridConflicts.Rows[i];
                string field = row.Cells["Field"].Value?.ToString();
                string resolution = row.Cells["Resolution"].Value?.ToString() ?? "Skip";

                if (!string.IsNullOrEmpty(field) && resolution != "Skip" && i < _conflicts.Count)
                {
                    string value = resolution == "Model"
                        ? _conflicts[i].ModelValue
                        : _conflicts[i].DrawingValue;

                    ConflictResolutions[field] = value;
                }
            }

            // Rename
            RenameApproved = _chkRename.Checked && _chkRename.Enabled;

            DialogResult = DialogResult.OK;
            Close();
        }

        private void UpdateStatus()
        {
            int checkedCount = 0;
            foreach (DataGridViewRow row in _gridSuggestions.Rows)
            {
                if (row.Cells["Apply"].Value is true)
                    checkedCount++;
            }
            _lblStatus.Text = $"{checkedCount} of {_suggestions.Count} suggestions selected";
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult != DialogResult.OK && DialogResult != DialogResult.Cancel)
                DialogResult = DialogResult.Cancel;
            base.OnFormClosing(e);
        }
    }
}
