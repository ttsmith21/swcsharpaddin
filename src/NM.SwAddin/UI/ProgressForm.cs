using System;
using System.Drawing;
using System.Windows.Forms;

namespace NM.SwAddin.UI
{
    public sealed class ProgressForm : Form
    {
        private readonly Label _lbl;
        private readonly ProgressBar _bar;
        private readonly Button _btnCancel;
        private readonly Label _summary;

        public bool IsCanceled { get; private set; }

        public ProgressForm()
        {
            Text = "Processing...";
            Width = 600;
            Height = 160;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            _lbl = new Label { Left = 10, Top = 10, Width = 560, Height = 18, Text = "Ready" };
            _bar = new ProgressBar { Left = 10, Top = 36, Width = 560, Height = 20, Minimum = 0, Maximum = 100, Style = ProgressBarStyle.Continuous };
            _summary = new Label { Left = 10, Top = 62, Width = 560, Height = 18, Text = string.Empty };
            _btnCancel = new Button { Left = 490, Top = 88, Width = 80, Height = 28, Text = "Cancel" };
            _btnCancel.Click += (s, e) => { IsCanceled = true; _btnCancel.Enabled = false; };

            Controls.AddRange(new Control[] { _lbl, _bar, _summary, _btnCancel });
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
    }
}
