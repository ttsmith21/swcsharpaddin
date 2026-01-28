using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NM.Core.ProblemParts;
using NM.SwAddin.Pipeline;
using NM.SwAddin.Validation;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.UI
{
    public sealed class ProblemPartsForm : Form
    {
        private readonly ProblemPartManager _manager;
        private readonly ISldWorks _swApp;
        private readonly BindingList<ProblemPartManager.ProblemItem> _items;

        private DataGridView _grid;
        private Button _btnRetry;
        private Button _btnExport;
        private Button _btnClose;
        private Button _btnRevalidateAll;
        private Button _btnContinue;
        private Label _lblSummary;
        private Label _lblStatus;
        private Label _lblGoodCount;
        private ProgressBar _progress;

        /// <summary>
        /// Action selected by user when closing the dialog.
        /// </summary>
        public ProblemAction SelectedAction { get; private set; } = ProblemAction.Cancel;

        private int _goodCount;

        /// <summary>
        /// Creates the form using the default ProblemPartManager singleton.
        /// </summary>
        public ProblemPartsForm() : this(ProblemPartManager.Instance, null) { }

        public ProblemPartsForm(ProblemPartManager manager, ISldWorks swApp = null)
        {
            _manager = manager ?? ProblemPartManager.Instance;
            _swApp = swApp;
            _items = new BindingList<ProblemPartManager.ProblemItem>(_manager.GetProblemParts());

            Text = "Problem Parts";
            Width = 950;
            Height = 580;
            StartPosition = FormStartPosition.CenterParent;

            BuildUi();
            BindGrid();
            UpdateSummary();
        }

        /// <summary>
        /// Sets the count of good models that will be processed if user continues.
        /// </summary>
        public void ShowGoodCount(int count)
        {
            _goodCount = count;
            if (_lblGoodCount != null)
            {
                _lblGoodCount.Text = $"{count} valid parts ready to process";
                _btnContinue.Enabled = count > 0;
            }
        }

        private void BuildUi()
        {
            _grid = new DataGridView
            {
                Dock = DockStyle.Top,
                Height = 360,
                AutoGenerateColumns = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
            };

            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "FileName", HeaderText = "File Name", DataPropertyName = "DisplayName", Width = 200 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Configuration", HeaderText = "Config", DataPropertyName = "Configuration", Width = 100 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Problem", HeaderText = "Problem Description", DataPropertyName = "ProblemDescription", Width = 320 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Category", HeaderText = "Category", DataPropertyName = "Category", Width = 120 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Attempts", HeaderText = "Attempts", DataPropertyName = "RetryCount", Width = 70 });
            _grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "CanRetry", HeaderText = "Retryable", DataPropertyName = "CanRetry", Width = 70 });
            var btn = new DataGridViewButtonColumn { Name = "Action", HeaderText = "Action", Text = "Review", UseColumnTextForButtonValue = true, Width = 80 };
            _grid.Columns.Add(btn);
            _grid.CellContentClick += OnGridCellContentClick;

            _btnRetry = new Button { Text = "Retry Selected", Left = 10, Top = 370, Width = 130 };
            _btnRetry.Click += OnRetrySelected;

            _btnExport = new Button { Text = "Export CSV", Left = 150, Top = 370, Width = 110 };
            _btnExport.Click += OnExport;

            _btnRevalidateAll = new Button { Text = "Revalidate All", Left = 270, Top = 370, Width = 120 };
            _btnRevalidateAll.Click += OnRevalidateAll;

            _btnContinue = new Button { Text = "Continue with Good", Left = 400, Top = 370, Width = 140, Enabled = false };
            _btnContinue.Click += OnContinueWithGood;

            _btnClose = new Button { Text = "Cancel", Left = 550, Top = 370, Width = 80 };
            _btnClose.Click += (s, e) =>
            {
                SelectedAction = ProblemAction.Cancel;
                DialogResult = DialogResult.Cancel;
                Close();
            };

            _lblGoodCount = new Label { Left = 650, Top = 374, Width = 200, Height = 20, AutoSize = false, ForeColor = Color.DarkGreen };
            _lblSummary = new Label { Left = 10, Top = 410, Width = 900, Height = 60, AutoSize = false };
            _lblStatus = new Label { Left = 10, Top = 470, Width = 700, Height = 20, AutoSize = false, Text = "Ready" };
            _progress = new ProgressBar { Left = 720, Top = 468, Width = 200, Height = 22 };

            Controls.AddRange(new Control[] { _grid, _btnRetry, _btnExport, _btnRevalidateAll, _btnContinue, _btnClose, _lblGoodCount, _lblSummary, _lblStatus, _progress });
        }

        private void OnContinueWithGood(object sender, EventArgs e)
        {
            SelectedAction = ProblemAction.ContinueWithGood;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void BindGrid()
        {
            _grid.DataSource = _items;
        }

        private void OnGridCellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (_grid.Columns[e.ColumnIndex].Name == "Action")
            {
                var item = _items[e.RowIndex];
                using (var dlg = new ProblemReviewDialog(item, _swApp))
                {
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        item.UserReviewed = true;
                        if (dlg.Resolved)
                        {
                            _manager.RemoveResolvedPart(item);
                            _items.Remove(item);
                        }
                        _grid.Refresh();
                        UpdateSummary();
                    }
                }
            }
        }

        private void OnRetrySelected(object sender, EventArgs e)
        {
            var selected = _grid.SelectedRows
                .OfType<DataGridViewRow>()
                .Select(r => r.DataBoundItem as ProblemPartManager.ProblemItem)
                .Where(p => p != null)
                .ToList();

            if (selected.Count == 0)
            {
                MessageBox.Show(this, "Select one or more rows to retry.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _manager.MarkAsReviewed(selected);

            int retryable = selected.Count(p => p.CanRetry);
            MessageBox.Show(this, $"Marked {selected.Count} as reviewed. {retryable} retryable.", "Marked for Retry", MessageBoxButtons.OK, MessageBoxIcon.Information);
            UpdateSummary();
        }

        private void OnRevalidateAll(object sender, EventArgs e)
        {
            if (_swApp == null)
            {
                MessageBox.Show(this, "SolidWorks instance not available.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var toCheck = _manager.GetProblemParts();
            if (toCheck.Count == 0)
            {
                MessageBox.Show(this, "No problem parts to revalidate.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _progress.Maximum = toCheck.Count;
            _progress.Value = 0;

            foreach (var item in toCheck.ToList())
            {
                _lblStatus.Text = $"Revalidating: {item.DisplayName}";
                _progress.Value = Math.Min(_progress.Value + 1, _progress.Maximum);
                Application.DoEvents();

                try
                {
                    int errs = 0, warns = 0;
                    var model = _swApp.OpenDoc6(item.FilePath, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, item.Configuration, ref errs, ref warns) as IModelDoc2;
                    if (model == null || errs != 0) continue;

                    var validator = new PartValidationAdapter();
                    var swInfo = new NM.Core.Models.SwModelInfo(item.FilePath) { Configuration = item.Configuration ?? string.Empty };
                    var vr = validator.Validate(swInfo, model);
                    if (vr.Success)
                    {
                        var res = MainRunner.RunSinglePart(_swApp, model, options: null);
                        if (res.Success)
                        {
                            _manager.RemoveResolvedPart(item);
                            _items.Remove(item);
                        }
                    }
                }
                catch { }
            }
            _lblStatus.Text = "Revalidate complete";
            UpdateSummary();
        }

        private void OnExport(object sender, EventArgs e)
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                sfd.FileName = $"ProblemParts_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    ExportToCsv(sfd.FileName);
                }
            }
        }

        private void ExportToCsv(string path)
        {
            try
            {
                using (var w = new StreamWriter(path, false, Encoding.UTF8))
                {
                    w.WriteLine("File Path,Configuration,Component,Problem,Category,Attempts,First Encountered,Last Attempted");
                    foreach (var p in _manager.GetProblemParts())
                    {
                        w.WriteLine($"\"{p.FilePath}\",\"{p.Configuration}\",\"{p.ComponentName}\",\"{p.ProblemDescription}\",{p.Category},{p.RetryCount},{p.FirstEncountered:yyyy-MM-dd HH:mm:ss},{p.LastAttempted:yyyy-MM-dd HH:mm:ss}");
                    }
                }
                MessageBox.Show(this, $"Exported {_manager.GetProblemParts().Count} items to:\n{path}", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Export failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateSummary()
        {
            _lblSummary.Text = _manager.GenerateSummary();
        }
    }
}
