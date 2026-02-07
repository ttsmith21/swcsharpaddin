using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace NM.SwAddin.UI
{
    /// <summary>
    /// Simple progress dialog for batch operations. Shows file name, progress bar,
    /// elapsed time, and a Cancel button. Updates between parts via Application.DoEvents().
    ///
    /// During long single-part API calls (InsertBends2, etc.) this form freezes along
    /// with SolidWorks. That's OK -- the companion SwProgressBar provides a Windows
    /// taskbar progress overlay that stays visible and responsive independently.
    /// Together they give the user two signals: the taskbar shows overall progress
    /// continuously, and this form shows detailed file-level status between parts.
    /// </summary>
    public sealed class ProgressForm : Form
    {
        private readonly Label _lbl;
        private readonly ProgressBar _bar;
        private readonly Button _btnCancel;
        private readonly Label _summary;
        private readonly Label _elapsed;
        private readonly Timer _timer;
        private readonly Stopwatch _stopwatch;

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

            // Elapsed time ticker -- updates via DoEvents() between parts
            _stopwatch = Stopwatch.StartNew();
            _timer = new Timer { Interval = 1000 };
            _timer.Tick += (s, e) =>
            {
                var ts = _stopwatch.Elapsed;
                _elapsed.Text = ts.TotalHours >= 1
                    ? $"Elapsed: {(int)ts.TotalHours}:{ts.Minutes:D2}:{ts.Seconds:D2}"
                    : $"Elapsed: {(int)ts.TotalMinutes}:{ts.Seconds:D2}";
            };
            _timer.Start();
        }

        public void SetMax(int max)
        {
            if (max <= 0) max = 1;
            _bar.Maximum = max;
            _bar.Value = 0;
        }

        public void SetStep(int step, string current)
        {
            if (step < 0) step = 0;
            if (step > _bar.Maximum) step = _bar.Maximum;
            _bar.Value = step;
            _lbl.Text = current ?? string.Empty;
        }

        public void SetSummary(string text)
        {
            _summary.Text = text ?? string.Empty;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _timer?.Stop();
                _timer?.Dispose();
                _stopwatch?.Stop();
            }
            base.Dispose(disposing);
        }
    }
}
