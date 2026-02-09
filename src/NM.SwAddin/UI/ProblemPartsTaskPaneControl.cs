using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NM.Core.ProblemParts;
using NM.SwAddin.Geometry;
using NM.SwAddin.Pipeline;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.UI
{
    /// <summary>
    /// Forwards keyboard messages to editable WinForms controls (TextBox, ListBox)
    /// when they have focus inside the task pane. Without this, SolidWorks'
    /// accelerator table intercepts keystrokes before they reach WinForms.
    /// Only intercepts for input controls — never steals keys from SolidWorks'
    /// own property manager or other native UI.
    /// </summary>
    internal sealed class TaskPaneKeyboardFilter : IMessageFilter
    {
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_CHAR = 0x0102;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetFocus();

        private readonly Control _host;

        public TaskPaneKeyboardFilter(Control host)
        {
            _host = host;
        }

        public bool PreFilterMessage(ref Message m)
        {
            if (m.Msg == WM_KEYDOWN || m.Msg == WM_KEYUP || m.Msg == WM_CHAR ||
                m.Msg == WM_SYSKEYDOWN || m.Msg == WM_SYSKEYUP)
            {
                // Use Win32 GetFocus to get the actual focused HWND — more reliable
                // than WinForms ContainsFocus in SolidWorks task pane hosting.
                IntPtr focusedHwnd = GetFocus();
                if (focusedHwnd == IntPtr.Zero) return false;

                // Check if the focused HWND belongs to an editable child of our host
                var focused = Control.FromHandle(focusedHwnd);
                if (focused == null) return false;
                if (!IsEditableControl(focused)) return false;
                if (!_host.Contains(focused) && focused != _host) return false;

                SendMessage(focusedHwnd, m.Msg, m.WParam, m.LParam);
                return true;
            }
            return false;
        }

        private static bool IsEditableControl(Control c)
        {
            return c is TextBox || c is ListBox || c is ComboBox || c is RichTextBox;
        }
    }

    /// <summary>
    /// COM-visible UserControl hosted in the SolidWorks Task Pane.
    /// Provides the same problem-part wizard functionality as ProblemWizardForm
    /// but docked in the right-hand panel alongside the Custom Property Editor.
    /// </summary>
    [ComVisible(true)]
    [ProgId(PROGID)]
    [Guid("A7E3F8B1-4C2D-4E9A-B6F5-1D8A3C7E9F02")]
    public sealed class ProblemPartsTaskPaneControl : UserControl
    {
        // WM_GETDLGCODE constants - tells Windows we want all keyboard input
        private const int WM_GETDLGCODE = 0x0087;
        private const int DLGC_WANTALLKEYS = 0x0004;
        private const int DLGC_WANTARROWS = 0x0001;
        private const int DLGC_WANTTAB = 0x0002;
        private const int DLGC_WANTCHARS = 0x0080;

        private TaskPaneKeyboardFilter _keyboardFilter;
        public const string PROGID = "NM.SwAddin.ProblemPartsTaskPane";

        private ISldWorks _swApp;
        private List<ProblemPartManager.ProblemItem> _problems;
        private readonly List<ProblemPartManager.ProblemItem> _fixedProblems = new List<ProblemPartManager.ProblemItem>();
        private readonly HashSet<ProblemPartManager.ProblemItem> _reviewedItems = new HashSet<ProblemPartManager.ProblemItem>();
        private int _currentIndex;
        private int _initialGoodCount;
        private IModelDoc2 _currentDoc;
        private TubeDiagnosticInfo _tubeDiagnostics;

        // Header
        private Label _lblHeader;
        private Label _lblProgress;

        // Navigation
        private Button _btnPrevious;
        private Button _btnNext;

        // Problem info
        private Label _lblFileName;
        private Label _lblConfig;
        private Label _lblCategory;
        private TextBox _txtError;
        private ListBox _lstSuggestions;

        // Classification
        private GroupBox _grpClassify;
        private Button _btnMarkPUR;
        private Button _btnMarkMACH;
        private Button _btnMarkCUST;

        // Processing
        private GroupBox _grpProcess;
        private Button _btnMarkSM;
        private Button _btnMarkTUBE;
        private Button _btnSplit;

        // Tube diagnostics (collapsible)
        private GroupBox _grpTubeDiag;
        private FlowLayoutPanel _pnlDiagButtons;
        private Label _lblDiagStatus;

        // Actions
        private Button _btnRetry;
        private Button _btnSkip;
        private Button _btnContinue;
        private Button _btnCancel;

        // Progress
        private ProgressBar _progressBar;
        private Label _lblFixedCount;
        private Label _lblStatus;

        // Placeholder
        private Panel _pnlPlaceholder;
        private Panel _pnlContent;

        /// <summary>Raised when the user clicks "Continue with Good" or finishes all problems.</summary>
        public event EventHandler<ProblemAction> ActionCompleted;

        public ProblemAction SelectedAction { get; private set; } = ProblemAction.Cancel;
        public List<ProblemPartManager.ProblemItem> FixedProblems => _fixedProblems;

        /// <summary>True when the panel has active problems being reviewed.</summary>
        public bool IsActive => _problems != null && _problems.Count > 0;

        public ProblemPartsTaskPaneControl()
        {
            _keyboardFilter = new TaskPaneKeyboardFilter(this);
            BuildUI();
            ShowPlaceholder();
        }

        /// <summary>
        /// Tell Windows this control wants keyboard input, but only when an
        /// editable child (TextBox, ListBox) actually has focus. Otherwise
        /// let SolidWorks handle keys normally (property manager, shortcuts, etc.).
        /// </summary>
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_GETDLGCODE)
            {
                var focused = ActiveControl;
                if (focused is TextBox || focused is ListBox || focused is ComboBox)
                {
                    m.Result = (IntPtr)(DLGC_WANTALLKEYS | DLGC_WANTARROWS | DLGC_WANTTAB | DLGC_WANTCHARS);
                    return;
                }
            }
            base.WndProc(ref m);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            // Register globally at creation — the filter itself checks ContainsFocus
            // so it only intercepts when this panel (or a child) is focused.
            // OnEnter/OnLeave on the parent UserControl is unreliable when the user
            // clicks directly into a child TextBox from outside the control.
            Application.AddMessageFilter(_keyboardFilter);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Application.RemoveMessageFilter(_keyboardFilter);
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Called by TaskPaneManager after SolidWorks hosts the control.
        /// </summary>
        public void Initialize(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        /// <summary>
        /// Loads a set of problem parts into the panel for review.
        /// Called by WorkflowDispatcher or ReviewProblems command.
        /// </summary>
        public void LoadProblems(List<ProblemPartManager.ProblemItem> problems, int initialGoodCount)
        {
            _problems = problems ?? new List<ProblemPartManager.ProblemItem>();
            _fixedProblems.Clear();
            _reviewedItems.Clear();
            _initialGoodCount = initialGoodCount;
            _currentIndex = 0;
            SelectedAction = ProblemAction.Cancel;

            if (_problems.Count == 0)
            {
                ShowPlaceholder();
                return;
            }

            ShowContent();
            LoadCurrentProblem();
            UpdateProgress();
        }

        /// <summary>
        /// Clears the panel back to its empty placeholder state.
        /// </summary>
        public void Clear()
        {
            SaveAndCloseCurrentPart();
            _problems = null;
            _fixedProblems.Clear();
            _reviewedItems.Clear();
            ShowPlaceholder();
        }

        #region UI Construction

        private void BuildUI()
        {
            AutoScroll = true;
            BackColor = SystemColors.Control;
            Font = new Font("Segoe UI", 8.25f);

            // ========== Placeholder (no problems) ==========
            _pnlPlaceholder = new Panel { Dock = DockStyle.Fill, Visible = true };
            var lblEmpty = new Label
            {
                Text = "No problem parts to review.\n\nRun the pipeline to begin processing.\nAny problem parts will appear here.",
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 9f)
            };
            _pnlPlaceholder.Controls.Add(lblEmpty);

            // ========== Content panel (has problems) ==========
            _pnlContent = new Panel { Dock = DockStyle.Fill, Visible = false, AutoScroll = true };

            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(4)
            };

            // Header
            _lblHeader = new Label
            {
                Text = "Problem Parts",
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 2),
                ForeColor = Color.DarkRed
            };
            flow.Controls.Add(_lblHeader);

            // Navigation row
            var navPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                WrapContents = false,
                Margin = new Padding(0, 2, 0, 4)
            };
            _btnPrevious = new Button { Text = "<<", Width = 36, Height = 24 };
            _btnPrevious.Click += (s, e) => NavigatePrevious();
            _lblProgress = new Label
            {
                Text = "0 / 0",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(4, 4, 4, 0),
                Font = new Font("Segoe UI", 8.25f, FontStyle.Bold)
            };
            _btnNext = new Button { Text = ">>", Width = 36, Height = 24 };
            _btnNext.Click += (s, e) => NavigateNext();
            navPanel.Controls.AddRange(new Control[] { _btnPrevious, _lblProgress, _btnNext });
            flow.Controls.Add(navPanel);

            // File info
            _lblFileName = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 8.25f, FontStyle.Bold),
                MaximumSize = new Size(280, 0),
                Margin = new Padding(0, 0, 0, 1)
            };
            flow.Controls.Add(_lblFileName);

            _lblConfig = new Label { AutoSize = true, ForeColor = Color.DimGray, MaximumSize = new Size(280, 0), Margin = new Padding(0, 0, 0, 1) };
            flow.Controls.Add(_lblConfig);

            _lblCategory = new Label { AutoSize = true, ForeColor = Color.DimGray, MaximumSize = new Size(280, 0), Margin = new Padding(0, 0, 0, 4) };
            flow.Controls.Add(_lblCategory);

            // Error box
            _txtError = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Width = 274,
                Height = 52,
                BackColor = Color.MistyRose,
                ForeColor = Color.DarkRed,
                Margin = new Padding(0, 0, 0, 4)
            };
            flow.Controls.Add(_txtError);

            // Suggestions
            _lstSuggestions = new ListBox
            {
                Width = 274,
                Height = 60,
                Margin = new Padding(0, 0, 0, 4)
            };
            flow.Controls.Add(_lstSuggestions);

            // Classification group
            _grpClassify = new GroupBox { Text = "Classify", Width = 274, Height = 56 };
            _btnMarkPUR = new Button { Left = 4, Top = 16, Width = 82, Height = 28, Text = "PUR", BackColor = Color.LightYellow };
            _btnMarkPUR.Click += (s, e) => OnClassify(ProblemPartManager.PartTypeOverride.Purchased);
            _btnMarkMACH = new Button { Left = 92, Top = 16, Width = 82, Height = 28, Text = "MACH", BackColor = Color.LightYellow };
            _btnMarkMACH.Click += (s, e) => OnClassify(ProblemPartManager.PartTypeOverride.Machined);
            _btnMarkCUST = new Button { Left = 180, Top = 16, Width = 82, Height = 28, Text = "CUST", BackColor = Color.LightYellow };
            _btnMarkCUST.Click += (s, e) => OnClassify(ProblemPartManager.PartTypeOverride.CustomerSupplied);
            _grpClassify.Controls.AddRange(new Control[] { _btnMarkPUR, _btnMarkMACH, _btnMarkCUST });
            flow.Controls.Add(_grpClassify);

            // Processing group
            _grpProcess = new GroupBox { Text = "Process As", Width = 274, Height = 56 };
            _btnMarkSM = new Button { Left = 4, Top = 16, Width = 82, Height = 28, Text = "Sheet Metal", BackColor = Color.LightSteelBlue };
            _btnMarkSM.Click += (s, e) => OnProcessAs("SheetMetal");
            _btnMarkTUBE = new Button { Left = 92, Top = 16, Width = 82, Height = 28, Text = "Tube", BackColor = Color.LightSteelBlue };
            _btnMarkTUBE.Click += (s, e) => OnProcessAs("Tube");
            _btnSplit = new Button { Left = 180, Top = 16, Width = 82, Height = 28, Text = "Split\u2192Assy", BackColor = Color.LightCoral };
            _btnSplit.Click += (s, e) => OnSplit();
            _grpProcess.Controls.AddRange(new Control[] { _btnMarkSM, _btnMarkTUBE, _btnSplit });
            flow.Controls.Add(_grpProcess);

            // Tube diagnostics (collapsed by default)
            _grpTubeDiag = new GroupBox { Text = "Tube Diagnostics", Width = 274, Height = 74 };
            _pnlDiagButtons = new FlowLayoutPanel
            {
                Left = 4, Top = 16, Width = 264, Height = 28,
                FlowDirection = FlowDirection.LeftToRight, WrapContents = false
            };
            var btnCut = new Button { Text = "Cut", Width = 42, Height = 24 };
            btnCut.Click += (s, e) => OnTubeDiag(TubeDiagnosticKind.CutLength);
            var btnHoles = new Button { Text = "Holes", Width = 42, Height = 24 };
            btnHoles.Click += (s, e) => OnTubeDiag(TubeDiagnosticKind.Holes);
            var btnBound = new Button { Text = "Bound", Width = 46, Height = 24 };
            btnBound.Click += (s, e) => OnTubeDiag(TubeDiagnosticKind.Boundary);
            var btnProf = new Button { Text = "Prof", Width = 42, Height = 24 };
            btnProf.Click += (s, e) => OnTubeDiag(TubeDiagnosticKind.Profile);
            var btnAll = new Button { Text = "All", Width = 36, Height = 24 };
            btnAll.Click += (s, e) => OnTubeDiag(TubeDiagnosticKind.All);
            var btnClr = new Button { Text = "Clr", Width = 36, Height = 24 };
            btnClr.Click += (s, e) => OnTubeDiag(TubeDiagnosticKind.Clear);
            _pnlDiagButtons.Controls.AddRange(new Control[] { btnCut, btnHoles, btnBound, btnProf, btnAll, btnClr });
            _lblDiagStatus = new Label { Left = 4, Top = 48, Width = 264, Height = 20, AutoEllipsis = true, ForeColor = Color.DimGray };
            _grpTubeDiag.Controls.AddRange(new Control[] { _pnlDiagButtons, _lblDiagStatus });
            flow.Controls.Add(_grpTubeDiag);

            // Action buttons row 1
            var actPanel1 = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                WrapContents = false,
                Margin = new Padding(0, 4, 0, 2)
            };
            _btnRetry = new Button { Text = "Retry && Validate", Width = 110, Height = 30, BackColor = Color.LightGreen };
            _btnRetry.Click += OnRetry;
            _btnSkip = new Button { Text = "Skip", Width = 60, Height = 30 };
            _btnSkip.Click += OnSkip;
            actPanel1.Controls.AddRange(new Control[] { _btnRetry, _btnSkip });
            flow.Controls.Add(actPanel1);

            // Action buttons row 2
            var actPanel2 = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 4)
            };
            _btnContinue = new Button { Text = "Continue", Width = 90, Height = 30 };
            _btnContinue.Click += OnContinue;
            _btnCancel = new Button { Text = "Cancel", Width = 70, Height = 30 };
            _btnCancel.Click += OnCancelClick;
            actPanel2.Controls.AddRange(new Control[] { _btnContinue, _btnCancel });
            flow.Controls.Add(actPanel2);

            // Progress
            _progressBar = new ProgressBar { Width = 200, Height = 16, Margin = new Padding(0, 0, 0, 2) };
            flow.Controls.Add(_progressBar);

            _lblFixedCount = new Label { AutoSize = true, Margin = new Padding(0, 0, 0, 2) };
            flow.Controls.Add(_lblFixedCount);

            _lblStatus = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(274, 40),
                ForeColor = Color.DarkBlue,
                Margin = new Padding(0, 0, 0, 4)
            };
            flow.Controls.Add(_lblStatus);

            _pnlContent.Controls.Add(flow);

            Controls.Add(_pnlContent);
            Controls.Add(_pnlPlaceholder);
        }

        private void ShowPlaceholder()
        {
            _pnlPlaceholder.Visible = true;
            _pnlContent.Visible = false;
        }

        private void ShowContent()
        {
            _pnlPlaceholder.Visible = false;
            _pnlContent.Visible = true;
        }

        #endregion

        #region Navigation

        private void LoadCurrentProblem()
        {
            SaveAndCloseCurrentPart();

            if (_problems == null || _problems.Count == 0)
            {
                ShowPlaceholder();
                return;
            }

            if (_currentIndex < 0) _currentIndex = 0;
            if (_currentIndex >= _problems.Count) _currentIndex = _problems.Count - 1;

            var item = _problems[_currentIndex];

            _lblProgress.Text = $"{_currentIndex + 1} / {_problems.Count}";
            _lblHeader.Text = $"Problem Parts ({_problems.Count - _reviewedItems.Count} remaining)";
            _lblFileName.Text = item.DisplayName ?? "(Unknown)";
            _lblConfig.Text = $"Config: {item.Configuration ?? "(default)"}";
            _lblCategory.Text = $"Category: {item.Category}";
            _txtError.Text = item.ProblemDescription ?? "Unknown error";
            _txtError.BackColor = Color.MistyRose;

            LoadSuggestions(item);
            UpdateNavigationButtons();
            UpdatePartTypeHint(item);

            _tubeDiagnostics = null;
            _lblDiagStatus.Text = "";
            _lblStatus.Text = "Fix the issue, then Retry & Validate.";

            OpenCurrentPart();
        }

        private void LoadSuggestions(ProblemPartManager.ProblemItem item)
        {
            _lstSuggestions.Items.Clear();
            var suggestions = ProblemSuggestionProvider.GetSuggestions(item.Category, item.ProblemDescription);
            foreach (var s in suggestions)
                _lstSuggestions.Items.Add(s);
        }

        private void UpdateNavigationButtons()
        {
            _btnPrevious.Enabled = _currentIndex > 0;
            _btnNext.Enabled = _currentIndex < _problems.Count - 1;
        }

        private void UpdateProgress()
        {
            int total = _problems?.Count ?? 0;
            int reviewed = _reviewedItems.Count;
            int fixedCount = _fixedProblems.Count;
            int skipped = reviewed - fixedCount;

            _progressBar.Maximum = total > 0 ? total : 1;
            _progressBar.Value = Math.Min(reviewed, _progressBar.Maximum);

            if (reviewed >= total && total > 0)
            {
                _lblFixedCount.Text = $"Done: {fixedCount} fixed, {skipped} skipped";
                _lblHeader.Text = "Problem Parts (all reviewed)";
                _lblStatus.Text = $"All {total} problems reviewed. Click Continue.";
                _lblStatus.ForeColor = Color.DarkGreen;
                _btnContinue.BackColor = Color.LightGreen;
                _btnContinue.Font = new Font(_btnContinue.Font, FontStyle.Bold);
            }
            else
            {
                _lblFixedCount.Text = $"{reviewed} / {total} reviewed ({fixedCount} fixed)";
                _lblHeader.Text = $"Problem Parts ({total - reviewed} remaining)";
            }

            int ready = _initialGoodCount + fixedCount;
            _btnContinue.Text = $"Continue ({ready})";
        }

        private void UpdatePartTypeHint(ProblemPartManager.ProblemItem item)
        {
            object hintObj;
            if (item.Metadata.TryGetValue("PurchasedHint", out hintObj) && hintObj is string hint && !string.IsNullOrEmpty(hint))
            {
                _grpClassify.Text = $"Classify (Likely PUR)";
                _grpClassify.ForeColor = Color.DarkOrange;
                _btnMarkPUR.BackColor = Color.Gold;
            }
            else
            {
                _grpClassify.Text = "Classify";
                _grpClassify.ForeColor = SystemColors.ControlText;
                _btnMarkPUR.BackColor = Color.LightYellow;
            }
            _btnMarkMACH.BackColor = Color.LightYellow;
            _btnMarkCUST.BackColor = Color.LightYellow;
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
            if (_swApp == null || _problems == null || _problems.Count == 0 || _currentIndex >= _problems.Count) return;

            var item = _problems[_currentIndex];
            _currentDoc = ProblemPartActions.OpenPart(item, _swApp);

            if (_currentDoc == null)
                _lblStatus.Text = $"Failed to open part.";
        }

        #endregion

        #region Action Handlers

        private void OnClassify(ProblemPartManager.PartTypeOverride typeOverride)
        {
            if (_problems == null || _problems.Count == 0 || _currentIndex >= _problems.Count) return;
            var item = _problems[_currentIndex];

            var result = ProblemPartActions.ClassifyPart(item, typeOverride, _currentDoc, _fixedProblems);
            _reviewedItems.Add(item);
            _lblStatus.Text = result.Message;
            _txtError.BackColor = Color.LightGoldenrodYellow;
            _txtError.Text = $"Classified as {typeOverride}";

            UpdateProgress();
            AutoAdvance();
        }

        private void OnProcessAs(string classification)
        {
            if (_currentDoc == null || _problems == null || _problems.Count == 0 || _currentIndex >= _problems.Count) return;
            var item = _problems[_currentIndex];

            _lblStatus.Text = $"Processing as {classification}...";
            Application.DoEvents();

            var result = ProblemPartActions.RunAsClassification(item, classification, _swApp, _currentDoc, _fixedProblems);

            if (result.Success)
            {
                _reviewedItems.Add(item);
                _lblStatus.Text = result.Message;
                _txtError.BackColor = Color.LightGreen;
                _txtError.Text = $"Processed as {classification}";
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

        private void OnSplit()
        {
            if (_currentDoc == null || _problems == null || _problems.Count == 0 || _currentIndex >= _problems.Count) return;
            var item = _problems[_currentIndex];

            _lblStatus.Text = "Splitting...";
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

        private void OnTubeDiag(TubeDiagnosticKind kind)
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

        private void OnRetry(object sender, EventArgs e)
        {
            if (_swApp == null || _problems == null || _problems.Count == 0) return;
            var item = _problems[_currentIndex];

            _lblStatus.Text = $"Validating...";
            Application.DoEvents();

            var result = ProblemPartActions.RetryAndValidate(item, _swApp, _currentDoc, _fixedProblems);

            if (result.Success)
            {
                _reviewedItems.Add(item);
                _lblStatus.Text = result.Message;
                _txtError.BackColor = Color.LightGreen;
                _txtError.Text = "Validation passed!";
                UpdateProgress();
                AutoAdvance();
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
            if (_problems != null && _problems.Count > 0 && _currentIndex < _problems.Count)
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

        private void OnContinue(object sender, EventArgs e)
        {
            SaveAndCloseCurrentPart();
            SelectedAction = ProblemAction.ContinueWithGood;
            ActionCompleted?.Invoke(this, ProblemAction.ContinueWithGood);
        }

        private void OnCancelClick(object sender, EventArgs e)
        {
            SaveAndCloseCurrentPart();
            SelectedAction = ProblemAction.Cancel;
            ActionCompleted?.Invoke(this, ProblemAction.Cancel);
        }

        private void AutoAdvance()
        {
            if (_currentIndex < _problems.Count - 1)
            {
                _lblStatus.Text += " - Next...";
                Application.DoEvents();
                System.Threading.Thread.Sleep(400);
                _currentIndex++;
                LoadCurrentProblem();
            }
            else
            {
                UpdateProgress();
            }
        }

        #endregion
    }
}
