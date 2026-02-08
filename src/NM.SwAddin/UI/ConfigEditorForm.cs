using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using NM.Core.Config;
using NM.Core.Config.Sections;

namespace NM.SwAddin.UI
{
    /// <summary>
    /// Configuration editor form with tabbed interface for editing nm-config.json values.
    /// Tabs: Rates, Material, Manufacturing, Paths &amp; Logging.
    /// </summary>
    public sealed class ConfigEditorForm : Form
    {
        private TabControl _tabs;
        private Label _lblConfigPath;
        private Button _btnSave, _btnCancel, _btnReset;

        // Tab 1 - Rates
        private TextBox _txtF115, _txtF140, _txtF145, _txtF155, _txtF210, _txtF220;
        private TextBox _txtF300, _txtF325, _txtF385, _txtF400, _txtF500, _txtF525, _txtEng;

        // Tab 2 - Material pricing
        private TextBox _txtSs304, _txtSs316, _txtCarbon, _txtAl6061, _txtAl5052, _txtGalv, _txtDefaultCost;
        // Material densities
        private TextBox _txtDensSs, _txtDensCarbon, _txtDensAl, _txtDensSteel;

        // Tab 3 - Manufacturing
        private TextBox _txtBrakeSetupFixed, _txtBrakeSetupPerFt;
        private TextBox _txtLaserSetupPerSheet, _txtLaserSetupFixed, _txtLaserMinSetup;
        private TextBox _txtWaterjetSetupFixed, _txtWaterjetSetupPerLoad;
        private TextBox _txtSheetWidth, _txtSheetLength;
        private TextBox _txtDeburRate;

        // Tab 4 - Paths & Logging
        private TextBox _txtMaterialDataPath, _txtErrorLogPath;
        private TextBox _txtBendTableSs, _txtBendTableCs;
        private CheckBox _chkLogEnabled, _chkShowWarnings, _chkDebugMode, _chkPerfMon, _chkProductionMode;

        public ConfigEditorForm()
        {
            InitializeComponent();
            LoadCurrentValues();
        }

        private void InitializeComponent()
        {
            Text = "NM AutoPilot - Settings";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 9F);
            ClientSize = new Size(640, 580);

            _tabs = new TabControl { Left = 8, Top = 8, Width = 624, Height = 490 };

            BuildRatesTab();
            BuildMaterialTab();
            BuildManufacturingTab();
            BuildPathsTab();

            Controls.Add(_tabs);

            // Bottom bar
            string configPath = NmConfigProvider.LoadedConfigPath ?? "(using compiled defaults)";
            _lblConfigPath = new Label
            {
                Text = configPath,
                Left = 12, Top = 508, Width = 380, Height = 20,
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleLeft
            };
            Controls.Add(_lblConfigPath);

            _btnReset = new Button { Text = "Reset Defaults", Left = 340, Top = 540, Width = 100, Height = 30 };
            _btnReset.Click += OnResetDefaults;
            Controls.Add(_btnReset);

            _btnSave = new Button { Text = "Save", Left = 448, Top = 540, Width = 85, Height = 30 };
            _btnSave.Click += OnSave;
            Controls.Add(_btnSave);

            _btnCancel = new Button { Text = "Cancel", Left = 541, Top = 540, Width = 85, Height = 30 };
            _btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            Controls.Add(_btnCancel);

            AcceptButton = _btnSave;
            CancelButton = _btnCancel;
        }

        #region Tab Builders

        private void BuildRatesTab()
        {
            var page = new TabPage("Rates");
            int y = 16;

            var header = new Label { Text = "Work Center Hourly Rates ($/hr)", Left = 12, Top = y, Width = 300, Font = new Font(Font, FontStyle.Bold) };
            page.Controls.Add(header);
            y += 28;

            _txtF115 = AddLabeledField(page, "F115 Laser Cutting", ref y);
            _txtF140 = AddLabeledField(page, "F140 Press Brake", ref y);
            _txtF145 = AddLabeledField(page, "F145 CNC Bending", ref y);
            _txtF155 = AddLabeledField(page, "F155 Waterjet", ref y);
            _txtF210 = AddLabeledField(page, "F210 Deburring", ref y);
            _txtF220 = AddLabeledField(page, "F220 Tapping", ref y);
            _txtF300 = AddLabeledField(page, "F300 Material Handling", ref y);
            _txtF325 = AddLabeledField(page, "F325 Roll Forming", ref y);
            _txtF385 = AddLabeledField(page, "F385 Assembly", ref y);
            _txtF400 = AddLabeledField(page, "F400 Welding", ref y);
            _txtF500 = AddLabeledField(page, "F500 Finishing", ref y);
            _txtF525 = AddLabeledField(page, "F525 Packaging", ref y);
            _txtEng = AddLabeledField(page, "ENG Engineering", ref y);

            page.AutoScroll = true;
            _tabs.TabPages.Add(page);
        }

        private void BuildMaterialTab()
        {
            var page = new TabPage("Material");
            int y = 16;

            var header1 = new Label { Text = "Material Pricing ($/lb)", Left = 12, Top = y, Width = 300, Font = new Font(Font, FontStyle.Bold) };
            page.Controls.Add(header1);
            y += 28;

            _txtSs304 = AddLabeledField(page, "Stainless 304", ref y, "$/lb");
            _txtSs316 = AddLabeledField(page, "Stainless 316", ref y, "$/lb");
            _txtCarbon = AddLabeledField(page, "Carbon Steel", ref y, "$/lb");
            _txtAl6061 = AddLabeledField(page, "Aluminum 6061", ref y, "$/lb");
            _txtAl5052 = AddLabeledField(page, "Aluminum 5052", ref y, "$/lb");
            _txtGalv = AddLabeledField(page, "Galvanized", ref y, "$/lb");
            _txtDefaultCost = AddLabeledField(page, "Default Cost", ref y, "$/lb");

            y += 12;
            var header2 = new Label { Text = "Material Densities (lb/in\u00B3)", Left = 12, Top = y, Width = 300, Font = new Font(Font, FontStyle.Bold) };
            page.Controls.Add(header2);
            y += 28;

            _txtDensSs = AddLabeledField(page, "Stainless 304/316", ref y, "lb/in\u00B3");
            _txtDensCarbon = AddLabeledField(page, "Carbon Steel A36", ref y, "lb/in\u00B3");
            _txtDensAl = AddLabeledField(page, "Aluminum", ref y, "lb/in\u00B3");
            _txtDensSteel = AddLabeledField(page, "Steel (General)", ref y, "lb/in\u00B3");

            page.AutoScroll = true;
            _tabs.TabPages.Add(page);
        }

        private void BuildManufacturingTab()
        {
            var page = new TabPage("Manufacturing");
            int y = 16;

            // Press Brake
            var header1 = new Label { Text = "Press Brake", Left = 12, Top = y, Width = 200, Font = new Font(Font, FontStyle.Bold) };
            page.Controls.Add(header1);
            y += 24;

            _txtBrakeSetupFixed = AddLabeledField(page, "Setup Fixed", ref y, "min");
            _txtBrakeSetupPerFt = AddLabeledField(page, "Setup Per Foot", ref y, "min/ft");

            y += 8;
            var header2 = new Label { Text = "Laser", Left = 12, Top = y, Width = 200, Font = new Font(Font, FontStyle.Bold) };
            page.Controls.Add(header2);
            y += 24;

            _txtLaserSetupPerSheet = AddLabeledField(page, "Setup Per Sheet", ref y, "min");
            _txtLaserSetupFixed = AddLabeledField(page, "Setup Fixed", ref y, "min");
            _txtLaserMinSetup = AddLabeledField(page, "Min Setup Hours", ref y, "hrs");

            y += 8;
            var header3 = new Label { Text = "Waterjet", Left = 12, Top = y, Width = 200, Font = new Font(Font, FontStyle.Bold) };
            page.Controls.Add(header3);
            y += 24;

            _txtWaterjetSetupFixed = AddLabeledField(page, "Setup Fixed", ref y, "min");
            _txtWaterjetSetupPerLoad = AddLabeledField(page, "Setup Per Load", ref y, "min");

            y += 8;
            var header4 = new Label { Text = "Standard Sheet", Left = 12, Top = y, Width = 200, Font = new Font(Font, FontStyle.Bold) };
            page.Controls.Add(header4);
            y += 24;

            _txtSheetWidth = AddLabeledField(page, "Width", ref y, "in");
            _txtSheetLength = AddLabeledField(page, "Length", ref y, "in");

            y += 8;
            var header5 = new Label { Text = "Deburring", Left = 12, Top = y, Width = 200, Font = new Font(Font, FontStyle.Bold) };
            page.Controls.Add(header5);
            y += 24;

            _txtDeburRate = AddLabeledField(page, "Rate", ref y, "in/min");

            page.AutoScroll = true;
            _tabs.TabPages.Add(page);
        }

        private void BuildPathsTab()
        {
            var page = new TabPage("Paths & Logging");
            int y = 16;

            var header1 = new Label { Text = "File Paths", Left = 12, Top = y, Width = 200, Font = new Font(Font, FontStyle.Bold) };
            page.Controls.Add(header1);
            y += 24;

            _txtMaterialDataPath = AddPathField(page, "Material Data", ref y);
            _txtBendTableSs = AddPathField(page, "Bend Table (SS)", ref y);
            _txtBendTableCs = AddPathField(page, "Bend Table (CS)", ref y);
            _txtErrorLogPath = AddPathField(page, "Error Log", ref y);

            y += 16;
            var header2 = new Label { Text = "Logging", Left = 12, Top = y, Width = 200, Font = new Font(Font, FontStyle.Bold) };
            page.Controls.Add(header2);
            y += 28;

            _chkLogEnabled = new CheckBox { Text = "Enable Logging", Left = 30, Top = y, Width = 200 };
            page.Controls.Add(_chkLogEnabled);
            y += 26;

            _chkShowWarnings = new CheckBox { Text = "Show Warnings", Left = 30, Top = y, Width = 200 };
            page.Controls.Add(_chkShowWarnings);
            y += 26;

            _chkDebugMode = new CheckBox { Text = "Debug Mode", Left = 30, Top = y, Width = 200 };
            page.Controls.Add(_chkDebugMode);
            y += 26;

            _chkPerfMon = new CheckBox { Text = "Performance Monitoring", Left = 30, Top = y, Width = 200 };
            page.Controls.Add(_chkPerfMon);
            y += 26;

            _chkProductionMode = new CheckBox { Text = "Production Mode", Left = 30, Top = y, Width = 200 };
            page.Controls.Add(_chkProductionMode);

            page.AutoScroll = true;
            _tabs.TabPages.Add(page);
        }

        #endregion

        #region Helpers

        private TextBox AddLabeledField(TabPage page, string labelText, ref int y, string suffix = "$/hr")
        {
            var lbl = new Label { Text = labelText, Left = 30, Top = y + 2, Width = 180, TextAlign = ContentAlignment.MiddleRight };
            var txt = new TextBox { Left = 220, Top = y, Width = 100, TextAlign = HorizontalAlignment.Right };
            var suf = new Label { Text = suffix, Left = 326, Top = y + 2, Width = 60, ForeColor = Color.Gray };
            page.Controls.Add(lbl);
            page.Controls.Add(txt);
            page.Controls.Add(suf);
            y += 28;
            return txt;
        }

        private TextBox AddPathField(TabPage page, string labelText, ref int y)
        {
            var lbl = new Label { Text = labelText, Left = 12, Top = y + 2, Width = 120, TextAlign = ContentAlignment.MiddleRight };
            var txt = new TextBox { Left = 140, Top = y, Width = 400 };
            var btn = new Button { Text = "...", Left = 546, Top = y - 1, Width = 40, Height = 24 };
            btn.Click += (s, e) =>
            {
                using (var ofd = new OpenFileDialog { FileName = txt.Text })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        txt.Text = ofd.FileName;
                }
            };
            page.Controls.Add(lbl);
            page.Controls.Add(txt);
            page.Controls.Add(btn);
            y += 32;
            return txt;
        }

        #endregion

        #region Load / Validate / Apply

        private void LoadCurrentValues()
        {
            var cfg = NmConfigProvider.Current;

            // Rates
            _txtF115.Text = cfg.WorkCenters.F115_LaserCutting.ToString("F2");
            _txtF140.Text = cfg.WorkCenters.F140_PressBrake.ToString("F2");
            _txtF145.Text = cfg.WorkCenters.F145_CncBending.ToString("F2");
            _txtF155.Text = cfg.WorkCenters.F155_Waterjet.ToString("F2");
            _txtF210.Text = cfg.WorkCenters.F210_Deburring.ToString("F2");
            _txtF220.Text = cfg.WorkCenters.F220_Tapping.ToString("F2");
            _txtF300.Text = cfg.WorkCenters.F300_MaterialHandling.ToString("F2");
            _txtF325.Text = cfg.WorkCenters.F325_RollForming.ToString("F2");
            _txtF385.Text = cfg.WorkCenters.F385_Assembly.ToString("F2");
            _txtF400.Text = cfg.WorkCenters.F400_Welding.ToString("F2");
            _txtF500.Text = cfg.WorkCenters.F500_Finishing.ToString("F2");
            _txtF525.Text = cfg.WorkCenters.F525_Packaging.ToString("F2");
            _txtEng.Text = cfg.WorkCenters.ENG_Engineering.ToString("F2");

            // Material pricing
            _txtSs304.Text = cfg.MaterialPricing.Stainless304_PerLb.ToString("F2");
            _txtSs316.Text = cfg.MaterialPricing.Stainless316_PerLb.ToString("F2");
            _txtCarbon.Text = cfg.MaterialPricing.CarbonSteel_PerLb.ToString("F2");
            _txtAl6061.Text = cfg.MaterialPricing.Aluminum6061_PerLb.ToString("F2");
            _txtAl5052.Text = cfg.MaterialPricing.Aluminum5052_PerLb.ToString("F2");
            _txtGalv.Text = cfg.MaterialPricing.Galvanized_PerLb.ToString("F2");
            _txtDefaultCost.Text = cfg.MaterialPricing.DefaultCostPerLb.ToString("F2");

            // Densities
            _txtDensSs.Text = cfg.MaterialDensities.Stainless304_316.ToString("F4");
            _txtDensCarbon.Text = cfg.MaterialDensities.CarbonSteelA36.ToString("F4");
            _txtDensAl.Text = cfg.MaterialDensities.Aluminum.ToString("F4");
            _txtDensSteel.Text = cfg.MaterialDensities.Steel_General.ToString("F4");

            // Manufacturing
            _txtBrakeSetupFixed.Text = cfg.Manufacturing.PressBrake.SetupFixedMinutes.ToString("F1");
            _txtBrakeSetupPerFt.Text = cfg.Manufacturing.PressBrake.SetupMinutesPerFoot.ToString("F2");
            _txtLaserSetupPerSheet.Text = cfg.Manufacturing.Laser.SetupMinutesPerSheet.ToString("F1");
            _txtLaserSetupFixed.Text = cfg.Manufacturing.Laser.SetupFixedMinutes.ToString("F1");
            _txtLaserMinSetup.Text = cfg.Manufacturing.Laser.MinSetupHours.ToString("F3");
            _txtWaterjetSetupFixed.Text = cfg.Manufacturing.Waterjet.SetupFixedMinutes.ToString("F1");
            _txtWaterjetSetupPerLoad.Text = cfg.Manufacturing.Waterjet.SetupMinutesPerLoad.ToString("F1");
            _txtSheetWidth.Text = cfg.Manufacturing.StandardSheet.WidthIn.ToString("F1");
            _txtSheetLength.Text = cfg.Manufacturing.StandardSheet.LengthIn.ToString("F1");
            _txtDeburRate.Text = cfg.Manufacturing.Deburring.RateInchesPerMinute.ToString("F1");

            // Paths
            var paths = cfg.Paths;
            _txtMaterialDataPath.Text = paths.MaterialDataPaths != null && paths.MaterialDataPaths.Length > 0 ? paths.MaterialDataPaths[0] : "";
            _txtBendTableSs.Text = paths.BendTables.StainlessSteel != null && paths.BendTables.StainlessSteel.Length > 0 ? paths.BendTables.StainlessSteel[0] : "";
            _txtBendTableCs.Text = paths.BendTables.CarbonSteel != null && paths.BendTables.CarbonSteel.Length > 0 ? paths.BendTables.CarbonSteel[0] : "";
            _txtErrorLogPath.Text = paths.ErrorLogPath ?? "";

            // Logging
            _chkLogEnabled.Checked = cfg.Logging.LogEnabled;
            _chkShowWarnings.Checked = cfg.Logging.ShowWarnings;
            _chkDebugMode.Checked = cfg.Logging.DebugMode;
            _chkPerfMon.Checked = cfg.Logging.PerformanceMonitoring;
            _chkProductionMode.Checked = cfg.Logging.ProductionMode;
        }

        private bool ValidateInputs(out List<string> errors)
        {
            errors = new List<string>();
            var fieldMap = new Dictionary<string, TextBox>
            {
                { "F115 Laser Cutting", _txtF115 }, { "F140 Press Brake", _txtF140 },
                { "F145 CNC Bending", _txtF145 }, { "F155 Waterjet", _txtF155 },
                { "F210 Deburring", _txtF210 }, { "F220 Tapping", _txtF220 },
                { "F300 Material Handling", _txtF300 }, { "F325 Roll Forming", _txtF325 },
                { "F385 Assembly", _txtF385 }, { "F400 Welding", _txtF400 },
                { "F500 Finishing", _txtF500 }, { "F525 Packaging", _txtF525 },
                { "ENG Engineering", _txtEng },
                { "SS 304 Price", _txtSs304 }, { "SS 316 Price", _txtSs316 },
                { "Carbon Steel Price", _txtCarbon }, { "AL 6061 Price", _txtAl6061 },
                { "AL 5052 Price", _txtAl5052 }, { "Galvanized Price", _txtGalv },
                { "Default Cost/lb", _txtDefaultCost },
                { "Density SS 304/316", _txtDensSs }, { "Density Carbon", _txtDensCarbon },
                { "Density Aluminum", _txtDensAl }, { "Density Steel", _txtDensSteel },
                { "Brake Setup Fixed", _txtBrakeSetupFixed }, { "Brake Setup/ft", _txtBrakeSetupPerFt },
                { "Laser Setup/Sheet", _txtLaserSetupPerSheet }, { "Laser Setup Fixed", _txtLaserSetupFixed },
                { "Laser Min Setup", _txtLaserMinSetup },
                { "Waterjet Setup Fixed", _txtWaterjetSetupFixed }, { "Waterjet Setup/Load", _txtWaterjetSetupPerLoad },
                { "Sheet Width", _txtSheetWidth }, { "Sheet Length", _txtSheetLength },
                { "Deburring Rate", _txtDeburRate }
            };

            foreach (var kvp in fieldMap)
            {
                if (!double.TryParse(kvp.Value.Text, out double val) || val < 0)
                {
                    errors.Add(kvp.Key);
                    kvp.Value.BackColor = Color.MistyRose;
                }
                else
                {
                    kvp.Value.BackColor = SystemColors.Window;
                }
            }

            return errors.Count == 0;
        }

        private void ApplyToConfig()
        {
            var cfg = NmConfigProvider.Current;

            // Rates
            cfg.WorkCenters.F115_LaserCutting = double.Parse(_txtF115.Text);
            cfg.WorkCenters.F140_PressBrake = double.Parse(_txtF140.Text);
            cfg.WorkCenters.F145_CncBending = double.Parse(_txtF145.Text);
            cfg.WorkCenters.F155_Waterjet = double.Parse(_txtF155.Text);
            cfg.WorkCenters.F210_Deburring = double.Parse(_txtF210.Text);
            cfg.WorkCenters.F220_Tapping = double.Parse(_txtF220.Text);
            cfg.WorkCenters.F300_MaterialHandling = double.Parse(_txtF300.Text);
            cfg.WorkCenters.F325_RollForming = double.Parse(_txtF325.Text);
            cfg.WorkCenters.F385_Assembly = double.Parse(_txtF385.Text);
            cfg.WorkCenters.F400_Welding = double.Parse(_txtF400.Text);
            cfg.WorkCenters.F500_Finishing = double.Parse(_txtF500.Text);
            cfg.WorkCenters.F525_Packaging = double.Parse(_txtF525.Text);
            cfg.WorkCenters.ENG_Engineering = double.Parse(_txtEng.Text);

            // Material pricing
            cfg.MaterialPricing.Stainless304_PerLb = double.Parse(_txtSs304.Text);
            cfg.MaterialPricing.Stainless316_PerLb = double.Parse(_txtSs316.Text);
            cfg.MaterialPricing.CarbonSteel_PerLb = double.Parse(_txtCarbon.Text);
            cfg.MaterialPricing.Aluminum6061_PerLb = double.Parse(_txtAl6061.Text);
            cfg.MaterialPricing.Aluminum5052_PerLb = double.Parse(_txtAl5052.Text);
            cfg.MaterialPricing.Galvanized_PerLb = double.Parse(_txtGalv.Text);
            cfg.MaterialPricing.DefaultCostPerLb = double.Parse(_txtDefaultCost.Text);

            // Densities
            cfg.MaterialDensities.Stainless304_316 = double.Parse(_txtDensSs.Text);
            cfg.MaterialDensities.CarbonSteelA36 = double.Parse(_txtDensCarbon.Text);
            cfg.MaterialDensities.Aluminum = double.Parse(_txtDensAl.Text);
            cfg.MaterialDensities.Steel_General = double.Parse(_txtDensSteel.Text);

            // Manufacturing
            cfg.Manufacturing.PressBrake.SetupFixedMinutes = double.Parse(_txtBrakeSetupFixed.Text);
            cfg.Manufacturing.PressBrake.SetupMinutesPerFoot = double.Parse(_txtBrakeSetupPerFt.Text);
            cfg.Manufacturing.Laser.SetupMinutesPerSheet = double.Parse(_txtLaserSetupPerSheet.Text);
            cfg.Manufacturing.Laser.SetupFixedMinutes = double.Parse(_txtLaserSetupFixed.Text);
            cfg.Manufacturing.Laser.MinSetupHours = double.Parse(_txtLaserMinSetup.Text);
            cfg.Manufacturing.Waterjet.SetupFixedMinutes = double.Parse(_txtWaterjetSetupFixed.Text);
            cfg.Manufacturing.Waterjet.SetupMinutesPerLoad = double.Parse(_txtWaterjetSetupPerLoad.Text);
            cfg.Manufacturing.StandardSheet.WidthIn = double.Parse(_txtSheetWidth.Text);
            cfg.Manufacturing.StandardSheet.LengthIn = double.Parse(_txtSheetLength.Text);
            cfg.Manufacturing.Deburring.RateInchesPerMinute = double.Parse(_txtDeburRate.Text);

            // Paths
            if (!string.IsNullOrWhiteSpace(_txtMaterialDataPath.Text))
                cfg.Paths.MaterialDataPaths = new[] { _txtMaterialDataPath.Text };
            if (!string.IsNullOrWhiteSpace(_txtBendTableSs.Text))
                cfg.Paths.BendTables.StainlessSteel = new[] { _txtBendTableSs.Text };
            if (!string.IsNullOrWhiteSpace(_txtBendTableCs.Text))
                cfg.Paths.BendTables.CarbonSteel = new[] { _txtBendTableCs.Text };
            cfg.Paths.ErrorLogPath = _txtErrorLogPath.Text;

            // Logging
            cfg.Logging.LogEnabled = _chkLogEnabled.Checked;
            cfg.Logging.ShowWarnings = _chkShowWarnings.Checked;
            cfg.Logging.DebugMode = _chkDebugMode.Checked;
            cfg.Logging.PerformanceMonitoring = _chkPerfMon.Checked;
            cfg.Logging.ProductionMode = _chkProductionMode.Checked;
        }

        #endregion

        #region Event Handlers

        private void OnSave(object sender, EventArgs e)
        {
            if (!ValidateInputs(out var errors))
            {
                MessageBox.Show(
                    "Invalid values in: " + string.Join(", ", errors) + "\n\nPlease enter valid non-negative numbers.",
                    "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                ApplyToConfig();
                NmConfigProvider.Save();
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save configuration:\n" + ex.Message,
                    "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnResetDefaults(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "Reset all settings to compiled defaults?\n\nThis will discard any custom values.",
                "Confirm Reset", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                NmConfigProvider.ResetToDefaults();
                LoadCurrentValues();
            }
        }

        #endregion
    }
}
