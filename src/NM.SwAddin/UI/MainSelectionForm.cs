using System;
using System.Windows.Forms;
using NM.Core;

namespace NM.SwAddin.UI
{
    /// <summary>
    /// Main user dialog shown at the start of the process to choose material and options.
    /// Mirrors the legacy VBA form behavior with C# WinForms.
    /// </summary>
    public sealed class MainSelectionForm : Form
    {
        public ProcessingOptions Options { get; private set; }

        // Material radio buttons
        private RadioButton rb304L, rb316L, rb6061, rb5052, rbA36, rbALNZD, rb309, rb2205, rbC22, rbAL6XN, rb409, rbOther;

        // Bend deduction
        private RadioButton rbBendTable, rbKFactor;
        private TextBox tbKFactor;

        // Output options
        private CheckBox cbCreateDXF, cbCreateDrawing, cbReport, cbSolidWorksVisible;

        // Buttons
        private Button btnOK, btnCancel, btnCustomProps;

        public MainSelectionForm()
        {
            Options = new ProcessingOptions();
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text = "(Semi)AutoPilot";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Width = 420;
            Height = 420;

            // Group: Material
            var gbMaterial = new GroupBox
            {
                Text = "Material",
                Left = 12,
                Top = 12,
                Width = 380,
                Height = 160
            };

            rb304L = MakeRadio("304L", 20, 25, true);
            rb316L = MakeRadio("316L", 20, 50);
            rb309 = MakeRadio("309", 20, 75);
            rbAL6XN = MakeRadio("AL6XN", 20, 100);

            rbALNZD = MakeRadio("ALNZD", 140, 25);
            rbA36 = MakeRadio("A36", 140, 50);
            rb2205 = MakeRadio("2205", 140, 75);
            rb409 = MakeRadio("409", 140, 100);

            rb6061 = MakeRadio("6061", 260, 25);
            rb5052 = MakeRadio("5052", 260, 50);
            rbC22 = MakeRadio("C22", 260, 75);
            rbOther = MakeRadio("Other", 260, 100);

            gbMaterial.Controls.AddRange(new Control[] {
                rb304L, rb316L, rb309, rbAL6XN,
                rbALNZD, rbA36, rb2205, rb409,
                rb6061, rb5052, rbC22, rbOther
            });

            // Group: Bend Deduction
            var gbBend = new GroupBox
            {
                Text = "Bend Deduction",
                Left = 12,
                Top = 180,
                Width = 380,
                Height = 60
            };

            rbBendTable = new RadioButton { Text = "Bend Table", Left = 20, Top = 25, Width = 100, Checked = true };
            rbKFactor = new RadioButton { Text = "KFactor", Left = 140, Top = 25, Width = 80 };
            tbKFactor = new TextBox { Text = ".333", Left = 230, Top = 23, Width = 60, Visible = false };

            rbBendTable.CheckedChanged += (s, e) => { if (rbBendTable.Checked) tbKFactor.Visible = false; };
            rbKFactor.CheckedChanged += (s, e) => { if (rbKFactor.Checked) tbKFactor.Visible = true; };

            gbBend.Controls.AddRange(new Control[] { rbBendTable, rbKFactor, tbKFactor });

            // Output options
            cbCreateDXF = new CheckBox { Text = "Create DXF", Left = 12, Top = 250, Width = 120, Checked = true };
            cbCreateDrawing = new CheckBox { Text = "Create Drawing", Left = 12, Top = 275, Width = 140 };
            cbReport = new CheckBox { Text = "Report", Left = 12, Top = 300, Width = 100, Checked = true };
            cbSolidWorksVisible = new CheckBox { Text = "SolidWorks Visible", Left = 250, Top = 250, Width = 140 };

            // Buttons
            btnCustomProps = new Button { Text = "Custom Properties", Left = 250, Top = 275, Width = 130 };
            btnOK = new Button { Text = "OK", Left = 220, Top = 340, Width = 75, DialogResult = DialogResult.OK };
            btnCancel = new Button { Text = "Cancel", Left = 305, Top = 340, Width = 75, DialogResult = DialogResult.Cancel };

            // Events mirroring VBA enable/disable logic
            rb304L.Click += (s, e) => EnableBendTable(true);
            rb316L.Click += (s, e) => EnableBendTable(true);
            rbA36.Click += (s, e) => EnableBendTable(true);
            rbALNZD.Click += (s, e) => EnableBendTable(true);
            rb309.Click += (s, e) => EnableBendTable(true);
            rb2205.Click += (s, e) => EnableBendTable(true);
            rbC22.Click += (s, e) => EnableBendTable(true);
            rbAL6XN.Click += (s, e) => EnableBendTable(true);
            rb409.Click += (s, e) => EnableBendTable(true);

            rb6061.Click += (s, e) => { EnableBendTable(false); rbKFactor.Checked = true; };
            rb5052.Click += (s, e) => { EnableBendTable(false); rbKFactor.Checked = true; };

            btnOK.Click += BtnOK_Click;
            btnCustomProps.Click += (s, e) => MessageBox.Show("Custom Properties dialog not yet implemented.");

            Controls.AddRange(new Control[]
            {
                gbMaterial, gbBend,
                cbCreateDXF, cbCreateDrawing, cbReport, cbSolidWorksVisible,
                btnCustomProps, btnOK, btnCancel
            });
        }

        private RadioButton MakeRadio(string text, int left, int top, bool isChecked = false)
        {
            return new RadioButton { Text = text, Left = left, Top = top, Width = 100, Checked = isChecked };
        }

        private void EnableBendTable(bool enable)
        {
            rbBendTable.Enabled = enable;
            if (!enable)
            {
                rbBendTable.Checked = false;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            var o = new ProcessingOptions();

            // Material mapping (match legacy VBA intent)
            if (rb304L.Checked) { o.Material = "304L"; o.MaterialType = "AISI 304"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rb316L.Checked) { o.Material = "316L"; o.MaterialType = "AISI 316 Stainless Steel Sheet (SS)"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rb409.Checked) { o.Material = "409"; o.MaterialType = "AISI 304"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rb2205.Checked) { o.Material = "2205"; o.MaterialType = "AISI 304"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rb309.Checked) { o.Material = "309"; o.MaterialType = "AISI 304"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rbC22.Checked) { o.Material = "C22"; o.MaterialType = "AISI 304"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rbAL6XN.Checked) { o.Material = "AL6XN"; o.MaterialType = "AISI 304"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rbA36.Checked) { o.Material = "A36"; o.MaterialType = "ASTM A36 Steel"; o.MaterialCategory = MaterialCategoryKind.CarbonSteel; }
            else if (rbALNZD.Checked) { o.Material = "ALNZD"; o.MaterialType = "ASTM A36 Steel"; o.MaterialCategory = MaterialCategoryKind.CarbonSteel; }
            else if (rb6061.Checked) { o.Material = "6061"; o.MaterialType = "6061 Alloy"; o.MaterialCategory = MaterialCategoryKind.Aluminum; }
            else if (rb5052.Checked) { o.Material = "5052"; o.MaterialType = "6061 Alloy"; o.MaterialCategory = MaterialCategoryKind.Aluminum; }
            else { o.Material = "Other"; o.MaterialType = "AISI 304"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }

            // Bend table vs K-factor
            if (rbBendTable.Checked && o.MaterialCategory != MaterialCategoryKind.Aluminum)
            {
                // Prefer company primary tables on O: share
                const string root = @"O:\\Engineering Department\\Solidworks\\Sheet Metal Bend Tables";
                string ss = System.IO.Path.Combine(root, "bend deduction_SS.xls");
                string cs = System.IO.Path.Combine(root, "bend deduction_CS.xls");

                if (o.MaterialCategory == MaterialCategoryKind.StainlessSteel)
                    o.BendTable = ss;
                else if (o.MaterialCategory == MaterialCategoryKind.CarbonSteel)
                    o.BendTable = cs;
                else
                    o.BendTable = NM.Core.Configuration.FilePaths.BendTableNone;

                o.KFactor = -1.0; // sentinel for table
            }
            else
            {
                o.BendTable = NM.Core.Configuration.FilePaths.BendTableNone;
                double k;
                if (!double.TryParse(tbKFactor.Text, out k)) k = NM.Core.Configuration.Defaults.DefaultKFactor;
                o.KFactor = k;
            }

            // Output
            o.CreateDXF = cbCreateDXF.Checked;
            o.CreateDrawing = cbCreateDrawing.Checked;
            o.GenerateReport = cbReport.Checked;
            o.SolidworksVisible = cbSolidWorksVisible.Checked;

            // Log the effective options to help diagnose bend selection
            try
            {
                ErrorHandler.DebugLog($"[UI] Options: Material='{o.Material}', Category={o.MaterialCategory}, BendChoice={(rbBendTable.Checked ? "BendTable" : "KFactor")}, BendTable='{o.BendTable}', KFactor={o.KFactor}");
            }
            catch { }

            Options = o;
        }
    }
}
