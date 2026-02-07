using System;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace NM.SwAddin.UI
{
    /// <summary>
    /// Progress dialog that runs on its own background thread with an independent
    /// message loop. This means the elapsed timer keeps ticking and the window
    /// stays responsive (repaint, Cancel button) even while a long SolidWorks API
    /// call (InsertBends2, TryFlatten, etc.) blocks the main STA thread.
    ///
    /// Usage from the main thread:
    ///   using (var progress = new ProgressForm())
    ///   {
    ///       progress.SetMax(files.Count);
    ///       progress.Show();   // opens on background thread
    ///       foreach (var file in files)
    ///       {
    ///           progress.SetStep(i, file);
    ///           Application.DoEvents(); // not strictly needed, but fine
    ///           // ... long blocking work ...
    ///       }
    ///   } // Dispose closes the form
    /// </summary>
    public sealed class ProgressForm : Form
    {
        private readonly Label _lbl;
        private readonly ProgressBar _bar;
        private readonly Button _btnCancel;
        private readonly Label _summary;
        private readonly Label _elapsed;
        private readonly System.Windows.Forms.Timer _uiTimer;
        private readonly Stopwatch _stopwatch;

        // Background thread that owns this form's message loop
        private Thread _uiThread;
        private volatile bool _threadRunning;
        private readonly ManualResetEventSlim _shown = new ManualResetEventSlim(false);

        public bool IsCanceled { get; private set; }

        public ProgressForm()
        {
            Text = "Processing...";
            Width = 600;
            Height = 180;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            TopMost = true; // Stay visible above frozen SolidWorks window

            _lbl = new Label { Left = 10, Top = 10, Width = 560, Height = 18, Text = "Ready" };
            _bar = new ProgressBar { Left = 10, Top = 36, Width = 560, Height = 20, Minimum = 0, Maximum = 100, Style = ProgressBarStyle.Continuous };
            _summary = new Label { Left = 10, Top = 62, Width = 400, Height = 18, Text = string.Empty };
            _elapsed = new Label { Left = 410, Top = 62, Width = 160, Height = 18, TextAlign = ContentAlignment.TopRight, Text = "Elapsed: 0:00" };
            _btnCancel = new Button { Left = 490, Top = 88, Width = 80, Height = 28, Text = "Cancel" };
            _btnCancel.Click += (s, e) => { IsCanceled = true; _btnCancel.Enabled = false; };

            Controls.AddRange(new Control[] { _lbl, _bar, _summary, _elapsed, _btnCancel });

            // Elapsed time ticker updates every second
            _stopwatch = Stopwatch.StartNew();
            _uiTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            _uiTimer.Tick += (s, e) =>
            {
                var ts = _stopwatch.Elapsed;
                _elapsed.Text = ts.TotalHours >= 1
                    ? $"Elapsed: {(int)ts.TotalHours}:{ts.Minutes:D2}:{ts.Seconds:D2}"
                    : $"Elapsed: {(int)ts.TotalMinutes}:{ts.Seconds:D2}";
            };
        }

        /// <summary>
        /// Shows the form on a dedicated background thread with its own message loop.
        /// The elapsed timer keeps ticking even when the main thread is blocked.
        /// Call from the main STA thread. Blocks until the form is visible.
        /// </summary>
        public new void Show()
        {
            if (_threadRunning) return;
            _threadRunning = true;

            _uiThread = new Thread(() =>
            {
                // The form was constructed on the main thread, but we run the
                // message loop here. WinForms allows this as long as we're
                // consistent about which thread calls Application.Run.
                // We re-create the form on THIS thread to avoid cross-thread issues.
                RunOnOwnThread();
            });
            _uiThread.SetApartmentState(ApartmentState.STA);
            _uiThread.IsBackground = true;
            _uiThread.Name = "ProgressForm_UI";
            _uiThread.Start();

            // Wait for the form to be visible before returning to caller
            _shown.Wait(5000);
        }

        private Form _threadForm;

        private void RunOnOwnThread()
        {
            // Create a new form on THIS thread (avoids cross-thread ownership issues)
            _threadForm = new Form
            {
                Text = Text,
                Width = Width,
                Height = Height,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                TopMost = true
            };

            var lbl = new Label { Left = 10, Top = 10, Width = 560, Height = 18, Text = "Ready" };
            var bar = new ProgressBar { Left = 10, Top = 36, Width = 560, Height = 20, Minimum = 0, Maximum = _bar.Maximum, Style = ProgressBarStyle.Continuous };
            var summary = new Label { Left = 10, Top = 62, Width = 400, Height = 18, Text = string.Empty };
            var elapsed = new Label { Left = 410, Top = 62, Width = 160, Height = 18, TextAlign = ContentAlignment.TopRight, Text = "Elapsed: 0:00" };
            var btnCancel = new Button { Left = 490, Top = 88, Width = 80, Height = 28, Text = "Cancel" };
            btnCancel.Click += (s, e) => { IsCanceled = true; btnCancel.Enabled = false; };

            _threadForm.Controls.AddRange(new Control[] { lbl, bar, summary, elapsed, btnCancel });

            // Store references for cross-thread updates
            _threadLbl = lbl;
            _threadBar = bar;
            _threadSummary = summary;
            _threadElapsed = elapsed;

            // Elapsed time ticker on this thread's message loop
            var timer = new System.Windows.Forms.Timer { Interval = 1000 };
            var sw = Stopwatch.StartNew();
            timer.Tick += (s, e) =>
            {
                var ts = sw.Elapsed;
                elapsed.Text = ts.TotalHours >= 1
                    ? $"Elapsed: {(int)ts.TotalHours}:{ts.Minutes:D2}:{ts.Seconds:D2}"
                    : $"Elapsed: {(int)ts.TotalMinutes}:{ts.Seconds:D2}";
            };
            timer.Start();

            _threadForm.Shown += (s, e) => _shown.Set();
            Application.Run(_threadForm);
        }

        // Controls owned by the background thread
        private volatile Label _threadLbl;
        private volatile ProgressBar _threadBar;
        private volatile Label _threadSummary;
        private volatile Label _threadElapsed;

        public void SetMax(int max)
        {
            if (max <= 0) max = 1;
            _bar.Maximum = max; // local copy for RunOnOwnThread init

            var b = _threadBar;
            if (b != null && _threadForm != null && !_threadForm.IsDisposed)
            {
                try { _threadForm.BeginInvoke((Action)(() => { b.Maximum = max; b.Value = 0; })); }
                catch { }
            }
        }

        public void SetStep(int step, string current)
        {
            var b = _threadBar;
            var l = _threadLbl;
            if (b != null && l != null && _threadForm != null && !_threadForm.IsDisposed)
            {
                try
                {
                    _threadForm.BeginInvoke((Action)(() =>
                    {
                        if (step < 0) step = 0;
                        if (step > b.Maximum) step = b.Maximum;
                        b.Value = step;
                        l.Text = current ?? string.Empty;
                    }));
                }
                catch { }
            }
        }

        public void SetSummary(string text)
        {
            var s = _threadSummary;
            if (s != null && _threadForm != null && !_threadForm.IsDisposed)
            {
                try { _threadForm.BeginInvoke((Action)(() => { s.Text = text ?? string.Empty; })); }
                catch { }
            }
        }

        /// <summary>
        /// Routes Close to the background thread form. Callers may call Close()
        /// explicitly (e.g. in a finally block) instead of relying on Dispose.
        /// </summary>
        public new void Close()
        {
            var f = _threadForm;
            if (f != null && !f.IsDisposed)
            {
                try { f.BeginInvoke((Action)(() => f.Close())); }
                catch { }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _threadRunning = false;
                _uiTimer?.Stop();
                _uiTimer?.Dispose();
                _stopwatch?.Stop();
                _shown?.Dispose();

                // Close the background-thread form, which exits Application.Run
                var f = _threadForm;
                if (f != null && !f.IsDisposed)
                {
                    try { f.BeginInvoke((Action)(() => f.Close())); }
                    catch { }
                }

                // Give the thread a moment to exit cleanly
                if (_uiThread != null && _uiThread.IsAlive)
                {
                    _uiThread.Join(2000);
                }
            }
            base.Dispose(disposing);
        }
    }
}
