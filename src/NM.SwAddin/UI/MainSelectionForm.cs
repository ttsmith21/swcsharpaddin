using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.UI
{
    /// <summary>
    /// Main settings dialog - mirrors VBA frmSolidWorksAutomatorMain.
    /// </summary>
    public sealed class MainSelectionForm : Form
    {
        public ProcessingOptions Options { get; private set; }
        private readonly ISldWorks _swApp;

        // Entry Point
        private RadioButton rbSinglePart, rbAssembly, rbFolder;
        private TextBox txtFolderPath;
        private Button btnBrowseFolder;

        // Custom Properties
        private TextBox txtCustomer, txtPrint, txtRevision, txtDescription;
        private CheckBox chkUsePartNum, chkGrainConstraint, chkCommonCut;

        // Material (18 options)
        private RadioButton rb304L, rb316L, rb309, rb310, rb321, rb330;
        private RadioButton rb409, rb430, rb2205, rb2507, rbC22, rbC276;
        private RadioButton rbAL6XN, rbALLOY31, rbA36, rbALNZD, rb5052, rb6061;

        // Costing
        private TrackBar sldComplexity, sldOptimization;
        private CheckBox chkPipeWelding;

        // Bend Deduction
        private RadioButton rbBendTable, rbKFactor;
        private TextBox txtKFactor;

        // Output Options
        private CheckBox chkCreateDXF, chkCreateDrawing, chkReport, chkErpExport, chkSolidWorksVisible;

        // Advanced Settings
        private ComboBox cboMode;
        private CheckBox chkEnableLogging, chkShowWarnings, chkPerfMonitoring, chkUseTaskPane;
        private TextBox txtLogPath;
        private Button btnBrowseLog;

        // Buttons
        private Button btnOK, btnCancel;

        public MainSelectionForm(ISldWorks swApp = null)
        {
            _swApp = swApp;
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
            Font = new Font("Segoe UI", 9F);
            ClientSize = new Size(600, 740);

            int left = 12;
            int width = 576;
            int y = 12;

            // ===== Entry Point =====
            var gbEntry = new GroupBox { Text = "Entry Point", Left = left, Top = y, Width = width, Height = 50 };
            rbSinglePart = new RadioButton { Text = "Single Part", Left = 15, Top = 20, Width = 90, Checked = true };
            rbAssembly = new RadioButton { Text = "Assembly", Left = 115, Top = 20, Width = 80 };
            rbFolder = new RadioButton { Text = "Folder", Left = 205, Top = 20, Width = 60 };
            btnBrowseFolder = new Button { Text = "Browse...", Left = 280, Top = 18, Width = 75, Height = 24, Enabled = false };
            txtFolderPath = new TextBox { Left = 360, Top = 19, Width = 200, Enabled = false };

            rbFolder.CheckedChanged += (s, e) => { btnBrowseFolder.Enabled = rbFolder.Checked; txtFolderPath.Enabled = rbFolder.Checked; };
            rbSinglePart.CheckedChanged += (s, e) => { if (rbSinglePart.Checked) { btnBrowseFolder.Enabled = false; txtFolderPath.Enabled = false; } };
            rbAssembly.CheckedChanged += (s, e) => { if (rbAssembly.Checked) { btnBrowseFolder.Enabled = false; txtFolderPath.Enabled = false; } };
            btnBrowseFolder.Click += (s, e) => { using (var fbd = new FolderBrowserDialog()) if (fbd.ShowDialog() == DialogResult.OK) txtFolderPath.Text = fbd.SelectedPath; };

            gbEntry.Controls.AddRange(new Control[] { rbSinglePart, rbAssembly, rbFolder, btnBrowseFolder, txtFolderPath });
            Controls.Add(gbEntry);
            y += 58;

            // ===== Custom Properties =====
            var gbProps = new GroupBox { Text = "Custom Properties", Left = left, Top = y, Width = width, Height = 80 };

            var lblCust = new Label { Text = "Customer:", Left = 15, Top = 22, Width = 60, TextAlign = ContentAlignment.MiddleRight };
            txtCustomer = new TextBox { Left = 80, Top = 20, Width = 120, MaxLength = 12 };
            var lblRev = new Label { Text = "Rev:", Left = 210, Top = 22, Width = 30, TextAlign = ContentAlignment.MiddleRight };
            txtRevision = new TextBox { Left = 245, Top = 20, Width = 50 };
            var lblPrint = new Label { Text = "Print:", Left = 305, Top = 22, Width = 35, TextAlign = ContentAlignment.MiddleRight };
            txtPrint = new TextBox { Left = 345, Top = 20, Width = 100 };
            var lblDesc = new Label { Text = "Desc:", Left = 455, Top = 22, Width = 35, TextAlign = ContentAlignment.MiddleRight };
            txtDescription = new TextBox { Left = 495, Top = 20, Width = 70, MaxLength = 20 };

            chkUsePartNum = new CheckBox { Text = "Use Part# as Print", Left = 15, Top = 50, Width = 130 };
            chkGrainConstraint = new CheckBox { Text = "Grain Constraint", Left = 155, Top = 50, Width = 115 };
            chkCommonCut = new CheckBox { Text = "Common Cut", Left = 280, Top = 50, Width = 100 };

            gbProps.Controls.AddRange(new Control[] { lblCust, txtCustomer, lblRev, txtRevision, lblPrint, txtPrint, lblDesc, txtDescription, chkUsePartNum, chkGrainConstraint, chkCommonCut });
            Controls.Add(gbProps);
            y += 88;

            // ===== Material =====
            var gbMaterial = new GroupBox { Text = "Material", Left = left, Top = y, Width = width, Height = 160 };

            int col1 = 15, col2 = 110, col3 = 205, col4 = 300, col5 = 395, col6 = 490;
            int row1 = 20, row2 = 45, row3 = 70, row4 = 95, row5 = 120;

            // Column 1 - 300 series
            rb304L = MakeRadio("304L", col1, row1, true); rb316L = MakeRadio("316L", col1, row2);
            rb309 = MakeRadio("309", col1, row3); rb310 = MakeRadio("310", col1, row4);
            rb321 = MakeRadio("321", col1, row5);

            // Column 2 - more stainless
            rb330 = MakeRadio("330", col2, row1); rb409 = MakeRadio("409", col2, row2);
            rb430 = MakeRadio("430", col2, row3); rb2205 = MakeRadio("2205", col2, row4);
            rb2507 = MakeRadio("2507", col2, row5);

            // Column 3 - nickel/super alloys
            rbC22 = MakeRadio("C22", col3, row1); rbC276 = MakeRadio("C276", col3, row2);
            rbAL6XN = MakeRadio("AL6XN", col3, row3); rbALLOY31 = MakeRadio("ALLOY31", col3, row4);

            // Column 4 - carbon steel
            rbA36 = MakeRadio("A36", col4, row1); rbALNZD = MakeRadio("ALNZD", col4, row2);

            // Column 5 - aluminum
            rb5052 = MakeRadio("5052", col5, row1); rb6061 = MakeRadio("6061", col5, row2);

            gbMaterial.Controls.AddRange(new Control[] {
                rb304L, rb316L, rb309, rb310, rb321,
                rb330, rb409, rb430, rb2205, rb2507,
                rbC22, rbC276, rbAL6XN, rbALLOY31,
                rbA36, rbALNZD,
                rb5052, rb6061
            });
            Controls.Add(gbMaterial);
            y += 168;

            // ===== Costing & Difficulty =====
            var gbCost = new GroupBox { Text = "Costing && Difficulty", Left = left, Top = y, Width = width, Height = 90 };

            var lblComplex = new Label { Text = "Complexity:", Left = 15, Top = 22, AutoSize = true };
            sldComplexity = new TrackBar { Left = 90, Top = 18, Width = 150, Minimum = 0, Maximum = 4, Value = 2, TickFrequency = 1 };
            var lblLow = new Label { Text = "Low", Left = 90, Top = 52, AutoSize = true, ForeColor = Color.Gray, Font = new Font(Font.FontFamily, 7.5f) };
            var lblHigh = new Label { Text = "High", Left = 215, Top = 52, AutoSize = true, ForeColor = Color.Gray, Font = new Font(Font.FontFamily, 7.5f) };

            var lblOpt = new Label { Text = "Optimization:", Left = 270, Top = 22, AutoSize = true };
            sldOptimization = new TrackBar { Left = 355, Top = 18, Width = 150, Minimum = 0, Maximum = 4, Value = 0, TickFrequency = 1 };
            var lblNone = new Label { Text = "None", Left = 355, Top = 52, AutoSize = true, ForeColor = Color.Gray, Font = new Font(Font.FontFamily, 7.5f) };
            var lblHigh2 = new Label { Text = "High", Left = 480, Top = 52, AutoSize = true, ForeColor = Color.Gray, Font = new Font(Font.FontFamily, 7.5f) };

            chkPipeWelding = new CheckBox { Text = "Pipe Welding", Left = 15, Top = 62, AutoSize = true };

            gbCost.Controls.AddRange(new Control[] { lblComplex, sldComplexity, lblLow, lblHigh, lblOpt, sldOptimization, lblNone, lblHigh2, chkPipeWelding });
            Controls.Add(gbCost);
            y += 98;

            // ===== Bend Deduction =====
            var gbBend = new GroupBox { Text = "Bend Deduction", Left = left, Top = y, Width = width, Height = 50 };

            rbBendTable = new RadioButton { Text = "Bend Table", Left = 15, Top = 20, Width = 90 };
            rbKFactor = new RadioButton { Text = "K-Factor:", Left = 120, Top = 20, Width = 80, Checked = true };
            txtKFactor = new TextBox { Text = NM.Core.Configuration.Defaults.DefaultKFactor.ToString("F3"), Left = 205, Top = 18, Width = 60 };

            rbBendTable.CheckedChanged += (s, e) => { txtKFactor.Enabled = !rbBendTable.Checked; };
            rbKFactor.CheckedChanged += (s, e) => { txtKFactor.Enabled = rbKFactor.Checked; };

            // Material-dependent bend table availability
            SetupMaterialBendLogic();

            gbBend.Controls.AddRange(new Control[] { rbBendTable, rbKFactor, txtKFactor });
            Controls.Add(gbBend);
            y += 58;

            // ===== Output Options =====
            var gbOutput = new GroupBox { Text = "Output Options", Left = left, Top = y, Width = width, Height = 50 };

            chkCreateDXF = new CheckBox { Text = "Create DXF", Left = 15, Top = 20, AutoSize = true };
            chkCreateDrawing = new CheckBox { Text = "Create Drawing", Left = 120, Top = 20, AutoSize = true };
            chkReport = new CheckBox { Text = "Report", Left = 250, Top = 20, AutoSize = true, Checked = true };
            chkErpExport = new CheckBox { Text = "ERP Export", Left = 340, Top = 20, AutoSize = true, Checked = true };
            chkSolidWorksVisible = new CheckBox { Text = "SW Visible", Left = 450, Top = 20, AutoSize = true };

            gbOutput.Controls.AddRange(new Control[] { chkCreateDXF, chkCreateDrawing, chkReport, chkErpExport, chkSolidWorksVisible });
            Controls.Add(gbOutput);
            y += 58;

            // ===== Advanced Settings =====
            var gbAdvanced = new GroupBox { Text = "Advanced Settings", Left = left, Top = y, Width = width, Height = 120 };

            var lblMode = new Label { Text = "Mode:", Left = 15, Top = 24, AutoSize = true };
            cboMode = new ComboBox { Left = 60, Top = 20, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList };
            cboMode.Items.AddRange(new[] { "Production", "Normal", "Debug" });
            cboMode.SelectedIndex = 1; // Normal

            chkEnableLogging = new CheckBox { Text = "Enable Logging", Left = 180, Top = 22, AutoSize = true, Checked = NM.Core.Configuration.Defaults.LogEnabledDefault };
            chkShowWarnings = new CheckBox { Text = "Show Warnings", Left = 310, Top = 22, AutoSize = true, Checked = NM.Core.Configuration.Defaults.ShowWarningsDefault };
            chkPerfMonitoring = new CheckBox { Text = "Perf Monitoring", Left = 440, Top = 22, AutoSize = true, Checked = NM.Core.Configuration.Defaults.EnablePerformanceMonitoringDefault };

            chkUseTaskPane = new CheckBox { Text = "Use Task Pane for Problems", Left = 15, Top = 52, AutoSize = true };

            var lblLog = new Label { Text = "Log Path:", Left = 15, Top = 75, AutoSize = true };
            txtLogPath = new TextBox { Left = 75, Top = 72, Width = 400, Text = NM.Core.Configuration.FilePaths.ErrorLogPath };
            btnBrowseLog = new Button { Text = "...", Left = 480, Top = 70, Width = 30, Height = 24 };
            btnBrowseLog.Click += (s, e) => { using (var sfd = new SaveFileDialog { Filter = "Text|*.txt", FileName = System.IO.Path.GetFileName(txtLogPath.Text) }) if (sfd.ShowDialog() == DialogResult.OK) txtLogPath.Text = sfd.FileName; };

            cboMode.SelectedIndexChanged += CboMode_SelectedIndexChanged;

            gbAdvanced.Controls.AddRange(new Control[] { lblMode, cboMode, chkEnableLogging, chkShowWarnings, chkPerfMonitoring, chkUseTaskPane, lblLog, txtLogPath, btnBrowseLog });
            Controls.Add(gbAdvanced);
            y += 128;

            // ===== Buttons =====
            btnOK = new Button { Text = "OK", Left = width - 170, Top = y, Width = 80, Height = 28, DialogResult = DialogResult.OK };
            btnCancel = new Button { Text = "Cancel", Left = width - 80, Top = y, Width = 80, Height = 28, DialogResult = DialogResult.Cancel };
            btnOK.Click += BtnOK_Click;

            Controls.AddRange(new Control[] { btnOK, btnCancel });

            AcceptButton = btnOK;
            CancelButton = btnCancel;

            // Check active document type and adjust entry point options (like VBA)
            SetEntryPointFromActiveDoc();

            // Pre-fill form from active part's material and custom properties
            PrePopulateFromActivePart();

            ApplyToolTips();
        }

        private void SetEntryPointFromActiveDoc()
        {
            if (_swApp == null) return;

            try
            {
                var swModel = (IModelDoc2)_swApp.ActiveDoc;
                if (swModel != null)
                {
                    var docType = (swDocumentTypes_e)swModel.GetType();
                    if (docType == swDocumentTypes_e.swDocASSEMBLY)
                    {
                        // Assembly open - select Assembly mode, disable Single Part
                        rbAssembly.Checked = true;
                        rbSinglePart.Enabled = false;
                    }
                    else if (docType == swDocumentTypes_e.swDocPART)
                    {
                        // Part open - select Single Part, disable Assembly
                        rbSinglePart.Checked = true;
                        rbAssembly.Enabled = false;
                    }
                    else
                    {
                        // Drawing or other - default to Folder mode
                        rbFolder.Checked = true;
                        rbSinglePart.Enabled = false;
                        rbAssembly.Enabled = false;
                    }
                }
                else
                {
                    // No document open - default to Folder mode
                    rbFolder.Checked = true;
                    rbSinglePart.Enabled = false;
                    rbAssembly.Enabled = false;
                }
            }
            catch
            {
                // If anything fails, leave defaults
            }
        }

        private RadioButton MakeRadio(string text, int left, int top, bool isChecked = false)
        {
            return new RadioButton { Text = text, Left = left, Top = top, Width = 85, Checked = isChecked };
        }

        private void SetupMaterialBendLogic()
        {
            // Aluminum disables bend table
            rb5052.Click += (s, e) => { rbBendTable.Enabled = false; rbKFactor.Checked = true; };
            rb6061.Click += (s, e) => { rbBendTable.Enabled = false; rbKFactor.Checked = true; };

            // All others enable bend table
            foreach (var rb in new[] { rb304L, rb316L, rb309, rb310, rb321, rb330, rb409, rb430, rb2205, rb2507, rbC22, rbC276, rbAL6XN, rbALLOY31, rbA36, rbALNZD })
                rb.Click += (s, e) => rbBendTable.Enabled = true;
        }

        private void CboMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cboMode.Text)
            {
                case "Production":
                    chkEnableLogging.Checked = false; chkEnableLogging.Enabled = false;
                    chkShowWarnings.Checked = false; chkShowWarnings.Enabled = false;
                    chkPerfMonitoring.Checked = false; chkPerfMonitoring.Enabled = false;
                    txtLogPath.Enabled = false; btnBrowseLog.Enabled = false;
                    break;
                case "Debug":
                    chkEnableLogging.Checked = true; chkEnableLogging.Enabled = false;
                    chkShowWarnings.Checked = true; chkShowWarnings.Enabled = false;
                    chkPerfMonitoring.Checked = true; chkPerfMonitoring.Enabled = false;
                    txtLogPath.Enabled = true; btnBrowseLog.Enabled = true;
                    break;
                default: // Normal
                    chkEnableLogging.Enabled = true;
                    chkShowWarnings.Enabled = true;
                    chkPerfMonitoring.Enabled = true;
                    txtLogPath.Enabled = chkEnableLogging.Checked;
                    btnBrowseLog.Enabled = chkEnableLogging.Checked;
                    break;
            }
        }

        private void ApplyToolTips()
        {
            var tt = new ToolTip { AutoPopDelay = 5000, InitialDelay = 400 };
            tt.SetToolTip(rbSinglePart, "Process a single SolidWorks part");
            tt.SetToolTip(rbAssembly, "Process all parts in an assembly");
            tt.SetToolTip(rbFolder, "Process all parts in a folder for quoting");
            tt.SetToolTip(sldComplexity, "Part complexity (0=Low to 4=High)");
            tt.SetToolTip(sldOptimization, "Optimization potential (0=None to 4=High)");
            tt.SetToolTip(chkPipeWelding, "Part involves pipe welding operations");
            tt.SetToolTip(rbBendTable, "Use material-specific bend table");
            tt.SetToolTip(rbKFactor, "Use K-Factor for bend calculations");
            tt.SetToolTip(cboMode, "Production=minimal logging, Normal=standard, Debug=verbose");
        }

        private void PrePopulateFromActivePart()
        {
            if (_swApp == null) return;
            try
            {
                var doc = (IModelDoc2)_swApp.ActiveDoc;
                if (doc == null) return;

                // --- Material ---
                // First try the custom property (written by previous pipeline run)
                string matProp = SwPropertyHelper.GetCustomPropertyValue(doc, "Material");
                if (string.IsNullOrEmpty(matProp))
                {
                    // Fall back to SW material database name
                    matProp = SolidWorksApiWrapper.GetMaterialName(doc);
                }
                if (!string.IsNullOrEmpty(matProp))
                    SelectMaterialRadio(matProp);

                // --- Custom Properties ---
                string cust = SwPropertyHelper.GetCustomPropertyValue(doc, "Customer");
                if (!string.IsNullOrEmpty(cust)) txtCustomer.Text = cust;

                string print = SwPropertyHelper.GetCustomPropertyValue(doc, "Print");
                if (!string.IsNullOrEmpty(print)) txtPrint.Text = print;

                string rev = SwPropertyHelper.GetCustomPropertyValue(doc, "Revision");
                if (!string.IsNullOrEmpty(rev)) txtRevision.Text = rev;

                string desc = SwPropertyHelper.GetCustomPropertyValue(doc, "Description");
                if (!string.IsNullOrEmpty(desc)) txtDescription.Text = desc;
            }
            catch
            {
                // Don't block form from opening
            }
        }

        private void SelectMaterialRadio(string materialName)
        {
            if (string.IsNullOrEmpty(materialName)) return;
            string m = materialName.Trim();

            // Map SW database names and short names to radio buttons
            var map = new Dictionary<string, RadioButton>(StringComparer.OrdinalIgnoreCase)
            {
                { "304L", rb304L }, { "AISI 304", rb304L },
                { "316L", rb316L }, { "AISI 316", rb316L }, { "AISI 316 Stainless Steel Sheet (SS)", rb316L },
                { "309", rb309 }, { "310", rb310 }, { "321", rb321 }, { "330", rb330 },
                { "409", rb409 }, { "430", rb430 },
                { "2205", rb2205 }, { "2507", rb2507 },
                { "C22", rbC22 }, { "Hastelloy C-22", rbC22 },
                { "C276", rbC276 },
                { "AL6XN", rbAL6XN }, { "ALLOY31", rbALLOY31 },
                { "A36", rbA36 }, { "ASTM A36 Steel", rbA36 },
                { "ALNZD", rbALNZD },
                { "5052", rb5052 }, { "5052-H32", rb5052 },
                { "6061", rb6061 }, { "6061 Alloy", rb6061 },
            };

            // Exact match
            if (map.TryGetValue(m, out var rb))
            {
                rb.Checked = true;
                return;
            }

            // Contains match (e.g., "AISI 304 Stainless..." contains "304")
            foreach (var kvp in map)
            {
                if (m.IndexOf(kvp.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    kvp.Value.Checked = true;
                    return;
                }
            }
            // No match — leave default 304L selected
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            var o = new ProcessingOptions();

            // Entry Point
            if (rbSinglePart.Checked) o.EntryPoint = ProcessingEntryPoint.SinglePart;
            else if (rbAssembly.Checked) o.EntryPoint = ProcessingEntryPoint.Assembly;
            else if (rbFolder.Checked) o.EntryPoint = ProcessingEntryPoint.FolderQuote;
            o.FolderPath = txtFolderPath.Text;

            // Custom Properties
            o.Customer = txtCustomer.Text;
            o.Print = txtPrint.Text;
            o.Revision = txtRevision.Text;
            o.Description = txtDescription.Text;
            o.UsePartNum = chkUsePartNum.Checked;
            o.GrainConstraint = chkGrainConstraint.Checked;
            o.CommonCut = chkCommonCut.Checked;

            // Material
            MapMaterial(o);

            // Costing
            o.Complexity = sldComplexity.Value;
            o.OptimizationPotential = sldOptimization.Value;
            o.PipeWelding = chkPipeWelding.Checked;

            // Bend Deduction
            if (rbBendTable.Checked && o.MaterialCategory != MaterialCategoryKind.Aluminum)
            {
                const string root = @"O:\Engineering Department\Solidworks\Sheet Metal Bend Tables";
                o.BendTable = o.MaterialCategory == MaterialCategoryKind.StainlessSteel
                    ? System.IO.Path.Combine(root, "bend deduction_SS.xls")
                    : System.IO.Path.Combine(root, "bend deduction_CS.xls");
                o.KFactor = -1.0;
            }
            else
            {
                o.BendTable = NM.Core.Configuration.FilePaths.BendTableNone;
                o.KFactor = double.TryParse(txtKFactor.Text, out double k) ? k : NM.Core.Configuration.Defaults.DefaultKFactor;
            }

            // Output
            o.CreateDXF = chkCreateDXF.Checked;
            o.CreateDrawing = chkCreateDrawing.Checked;
            o.GenerateReport = chkReport.Checked;
            o.GenerateErpExport = chkErpExport.Checked;
            o.SolidworksVisible = chkSolidWorksVisible.Checked;

            // Advanced Settings
            switch (cboMode.Text)
            {
                case "Production":
                    o.DebugLevel = DebuggingLevel.Production;
                    o.EnableLogging = false;
                    o.ShowWarnings = false;
                    o.EnablePerformanceMonitoring = false;
                    break;
                case "Debug":
                    o.DebugLevel = DebuggingLevel.Debug;
                    o.EnableLogging = true;
                    o.ShowWarnings = true;
                    o.EnablePerformanceMonitoring = true;
                    break;
                default:
                    o.DebugLevel = DebuggingLevel.Normal;
                    o.EnableLogging = chkEnableLogging.Checked;
                    o.ShowWarnings = chkShowWarnings.Checked;
                    o.EnablePerformanceMonitoring = chkPerfMonitoring.Checked;
                    break;
            }
            o.LogFilePath = txtLogPath.Text;

            // UI
            o.UseTaskPane = chkUseTaskPane.Checked;

            Options = o;
        }

        private void MapMaterial(ProcessingOptions o)
        {
            if (rb304L.Checked) { o.Material = "304L"; o.MaterialType = "AISI 304"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rb316L.Checked) { o.Material = "316L"; o.MaterialType = "AISI 316 Stainless Steel Sheet (SS)"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rb309.Checked) { o.Material = "309"; o.MaterialType = "309"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rb310.Checked) { o.Material = "310"; o.MaterialType = "310"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rb321.Checked) { o.Material = "321"; o.MaterialType = "321"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rb330.Checked) { o.Material = "330"; o.MaterialType = "330"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rb409.Checked) { o.Material = "409"; o.MaterialType = "409"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rb430.Checked) { o.Material = "430"; o.MaterialType = "430"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rb2205.Checked) { o.Material = "2205"; o.MaterialType = "2205"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rb2507.Checked) { o.Material = "2507"; o.MaterialType = "2507"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rbC22.Checked) { o.Material = "C22"; o.MaterialType = "Hastelloy C-22"; o.MaterialCategory = MaterialCategoryKind.Other; }
            else if (rbC276.Checked) { o.Material = "C276"; o.MaterialType = "C276"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rbAL6XN.Checked) { o.Material = "AL6XN"; o.MaterialType = "AL6XN"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rbALLOY31.Checked) { o.Material = "ALLOY31"; o.MaterialType = "ALLOY31"; o.MaterialCategory = MaterialCategoryKind.StainlessSteel; }
            else if (rbA36.Checked) { o.Material = "A36"; o.MaterialType = "ASTM A36 Steel"; o.MaterialCategory = MaterialCategoryKind.CarbonSteel; }
            else if (rbALNZD.Checked) { o.Material = "ALNZD"; o.MaterialType = "ASTM A36 Steel"; o.MaterialCategory = MaterialCategoryKind.CarbonSteel; }
            else if (rb5052.Checked) { o.Material = "5052"; o.MaterialType = "5052-H32"; o.MaterialCategory = MaterialCategoryKind.Aluminum; }
            else if (rb6061.Checked) { o.Material = "6061"; o.MaterialType = "6061 Alloy"; o.MaterialCategory = MaterialCategoryKind.Aluminum; }
        }
    }
}
