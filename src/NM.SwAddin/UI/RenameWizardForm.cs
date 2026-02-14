using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NM.Core.Rename;

namespace NM.SwAddin.UI
{
    /// <summary>
    /// WinForms wizard for reviewing AI-predicted component renames before execution.
    /// Shows a grid of current names, AI predictions, and editable final names.
    /// Includes find/replace bar for bulk name editing.
    /// </summary>
    public sealed class RenameWizardForm : Form
    {
        // Header
        private Label _lblAssembly;
        private Label _lblDrawing;
        private Label _lblStats;

        // Find/Replace bar
        private Label _lblFind;
        private TextBox _txtFind;
        private Label _lblReplace;
        private TextBox _txtReplace;
        private Button _btnFindReplace;

        // Main grid
        private DataGridView _grid;

        // Action buttons
        private Button _btnSelectAll;
        private Button _btnDeselectAll;
        private Button _btnAcceptAi;
        private Button _btnRename;
        private Button _btnCancel;

        // Status
        private Label _lblStatus;

        // Data
        private readonly List<RenameEntry> _entries;
        private readonly string _assemblyName;
        private readonly string _drawingName;

        /// <summary>
        /// Callback to highlight a component in the SolidWorks viewport when a row is selected.
        /// Set by the runner before showing the dialog.
        /// </summary>
        public Action<RenameEntry> OnRowSelected { get; set; }

        /// <summary>Approved rename entries (populated on Apply).</summary>
        public List<RenameEntry> ApprovedRenames { get; private set; }

        public RenameWizardForm(
            List<RenameEntry> entries,
            string assemblyName = null,
            string drawingName = null)
        {
            _entries = entries ?? new List<RenameEntry>();
            _assemblyName = assemblyName ?? "";
            _drawingName = drawingName ?? "";

            Text = "Rename Wizard - AI Component Matching";
            Width = 960;
            Height = 700;
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
            _lblAssembly = new Label { Left = 12, Top = 10, Width = 900, Font = new Font(Font.FontFamily, 9f, FontStyle.Bold), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            _lblDrawing = new Label { Left = 12, Top = 30, Width = 900, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            _lblStats = new Label { Left = 12, Top = 50, Width = 900, ForeColor = Color.DarkBlue, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };

            // Find/Replace bar
            int frY = 72;
            _lblFind = new Label { Left = 12, Top = frY + 3, Width = 35, Text = "Find:" };
            _txtFind = new TextBox { Left = 50, Top = frY, Width = 200 };
            _lblReplace = new Label { Left = 258, Top = frY + 3, Width = 55, Text = "Replace:" };
            _txtReplace = new TextBox { Left = 316, Top = frY, Width = 200 };
            _btnFindReplace = new Button { Left = 524, Top = frY - 1, Width = 80, Height = 24, Text = "Apply" };
            _btnFindReplace.Click += OnFindReplace;

            // Main grid
            _grid = new DataGridView
            {
                Left = 12, Top = 100, Width = 920, Height = 490,
                AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ReadOnly = false, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };

            // Columns
            var chkCol = new DataGridViewCheckBoxColumn { Name = "Apply", HeaderText = "", Width = 30 };
            _grid.Columns.Add(chkCol);
            _grid.Columns.Add("CurrentName", "Current Name");
            _grid.Columns.Add("AiPrediction", "AI Prediction");
            _grid.Columns.Add("NewName", "New Name");
            _grid.Columns.Add("Confidence", "Conf");
            _grid.Columns.Add("Reason", "Reason");

            // Read-only columns
            _grid.Columns["CurrentName"].ReadOnly = true;
            _grid.Columns["AiPrediction"].ReadOnly = true;
            _grid.Columns["Confidence"].ReadOnly = true;
            _grid.Columns["Reason"].ReadOnly = true;

            // NewName is editable
            _grid.Columns["NewName"].ReadOnly = false;

            // Column widths
            _grid.Columns["Confidence"].Width = 45;
            _grid.Columns["CurrentName"].FillWeight = 25;
            _grid.Columns["AiPrediction"].FillWeight = 25;
            _grid.Columns["NewName"].FillWeight = 25;
            _grid.Columns["Reason"].FillWeight = 25;

            _grid.SelectionChanged += OnGridSelectionChanged;
            _grid.CellEndEdit += OnCellEndEdit;

            // Buttons
            int btnY = 600;
            _btnSelectAll = new Button { Left = 12, Top = btnY, Width = 85, Height = 28, Text = "Select All", Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            _btnSelectAll.Click += OnSelectAll;

            _btnDeselectAll = new Button { Left = 105, Top = btnY, Width = 85, Height = 28, Text = "Deselect All", Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            _btnDeselectAll.Click += OnDeselectAll;

            _btnAcceptAi = new Button { Left = 198, Top = btnY, Width = 110, Height = 28, Text = "Accept All AI", Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            _btnAcceptAi.Click += OnAcceptAllAi;

            _btnRename = new Button { Left = 710, Top = btnY, Width = 110, Height = 28, Text = "Rename", Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            _btnRename.Click += OnRename;

            _btnCancel = new Button { Left = 828, Top = btnY, Width = 100, Height = 28, Text = "Cancel", Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            _btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            _lblStatus = new Label { Left = 316, Top = btnY + 5, Width = 380, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.DarkGreen, Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right };

            Controls.AddRange(new Control[]
            {
                _lblAssembly, _lblDrawing, _lblStats,
                _lblFind, _txtFind, _lblReplace, _txtReplace, _btnFindReplace,
                _grid,
                _btnSelectAll, _btnDeselectAll, _btnAcceptAi,
                _btnRename, _btnCancel, _lblStatus
            });
        }

        private void LoadData()
        {
            _lblAssembly.Text = $"Assembly: {_assemblyName}";
            _lblDrawing.Text = $"Drawing: {_drawingName}";

            int matched = _entries.Count(e => e.Confidence > 0);
            _lblStats.Text = $"Components: {_entries.Count} unique  |  AI matched: {matched}  |  Unmatched: {_entries.Count - matched}";

            foreach (var entry in _entries)
            {
                string nameNoExt = Path.GetFileNameWithoutExtension(entry.CurrentFileName) ?? entry.CurrentFileName;
                string prediction = entry.Confidence > 0 ? entry.PredictedName : "(no match)";

                int idx = _grid.Rows.Add(
                    entry.IsApproved,
                    nameNoExt,
                    prediction,
                    entry.FinalName ?? nameNoExt,
                    entry.Confidence > 0 ? (entry.Confidence * 100).ToString("F0") + "%" : "",
                    entry.MatchReason ?? ""
                );

                // Color-code by confidence
                var row = _grid.Rows[idx];
                if (entry.Confidence >= 0.80)
                    row.DefaultCellStyle.BackColor = Color.FromArgb(230, 255, 230); // Green
                else if (entry.Confidence >= 0.50)
                    row.DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 210); // Yellow
                // Low/no match stays white

                // Gray out the prediction column for no-match rows
                if (entry.Confidence <= 0)
                    row.Cells["AiPrediction"].Style.ForeColor = Color.Gray;
            }

            UpdateStatus();
        }

        private void OnFindReplace(object sender, EventArgs e)
        {
            string find = _txtFind.Text;
            string replace = _txtReplace.Text;

            if (string.IsNullOrEmpty(find)) return;

            int changed = 0;
            for (int i = 0; i < _grid.Rows.Count; i++)
            {
                var row = _grid.Rows[i];
                string currentNewName = row.Cells["NewName"].Value?.ToString() ?? "";
                if (currentNewName.IndexOf(find, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Case-insensitive replace
                    string updated = ReplaceIgnoreCase(currentNewName, find, replace ?? "");
                    row.Cells["NewName"].Value = updated;
                    changed++;
                }
            }

            _lblStatus.Text = $"Replaced in {changed} row(s)";
            _lblStatus.ForeColor = changed > 0 ? Color.DarkGreen : Color.DarkOrange;
        }

        private static string ReplaceIgnoreCase(string input, string search, string replacement)
        {
            int pos = input.IndexOf(search, StringComparison.OrdinalIgnoreCase);
            if (pos < 0) return input;
            return input.Substring(0, pos) + replacement + input.Substring(pos + search.Length);
        }

        private void OnGridSelectionChanged(object sender, EventArgs e)
        {
            if (_grid.SelectedRows.Count == 0 || OnRowSelected == null) return;

            int rowIndex = _grid.SelectedRows[0].Index;
            if (rowIndex >= 0 && rowIndex < _entries.Count)
            {
                try
                {
                    OnRowSelected(_entries[rowIndex]);
                }
                catch
                {
                    // Swallow selection highlight errors — non-critical
                }
            }
        }

        private void OnCellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == _grid.Columns["NewName"].Index && e.RowIndex >= 0 && e.RowIndex < _entries.Count)
            {
                string newVal = _grid.Rows[e.RowIndex].Cells["NewName"].Value?.ToString() ?? "";
                _entries[e.RowIndex].FinalName = newVal;
            }
            UpdateStatus();
        }

        private void OnSelectAll(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in _grid.Rows)
                row.Cells["Apply"].Value = true;
            UpdateStatus();
        }

        private void OnDeselectAll(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in _grid.Rows)
                row.Cells["Apply"].Value = false;
            UpdateStatus();
        }

        private void OnAcceptAllAi(object sender, EventArgs e)
        {
            for (int i = 0; i < _grid.Rows.Count && i < _entries.Count; i++)
            {
                if (_entries[i].Confidence >= 0.50)
                {
                    _grid.Rows[i].Cells["Apply"].Value = true;
                    _grid.Rows[i].Cells["NewName"].Value = _entries[i].PredictedName;
                    _entries[i].FinalName = _entries[i].PredictedName;
                }
            }
            UpdateStatus();
        }

        private void OnRename(object sender, EventArgs e)
        {
            // Sync grid edits back to entries
            for (int i = 0; i < _grid.Rows.Count && i < _entries.Count; i++)
            {
                var row = _grid.Rows[i];
                _entries[i].IsApproved = row.Cells["Apply"].Value is true;
                _entries[i].FinalName = row.Cells["NewName"].Value?.ToString() ?? _entries[i].CurrentFileName;
            }

            ApprovedRenames = _entries.Where(en => en.IsApproved).ToList();

            // Validate: check for empty names
            var emptyNames = ApprovedRenames.Where(r => string.IsNullOrWhiteSpace(r.FinalName)).ToList();
            if (emptyNames.Count > 0)
            {
                MessageBox.Show(
                    $"{emptyNames.Count} approved item(s) have empty names. Please fill in all names before renaming.",
                    "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Validate: check for duplicates
            var dupes = ApprovedRenames
                .GroupBy(r => r.FinalName, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key)
                .ToList();

            if (dupes.Count > 0)
            {
                MessageBox.Show(
                    $"Duplicate names detected:\n{string.Join(", ", dupes)}\n\nPlease resolve duplicates before renaming.",
                    "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private void UpdateStatus()
        {
            int checkedCount = 0;
            foreach (DataGridViewRow row in _grid.Rows)
            {
                if (row.Cells["Apply"].Value is true)
                    checkedCount++;
            }
            _lblStatus.Text = $"{checkedCount} of {_entries.Count} components selected for rename";
            _lblStatus.ForeColor = Color.DarkGreen;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult != DialogResult.OK && DialogResult != DialogResult.Cancel)
                DialogResult = DialogResult.Cancel;
            base.OnFormClosing(e);
        }
    }
}
