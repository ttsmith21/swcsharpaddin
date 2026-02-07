using System.Linq;
using NM.Core.Config;
using NM.Core.Config.Tables;
using Xunit;

namespace NM.Core.Tests
{
    /// <summary>
    /// Smoke tests for the NmConfigProvider configuration system.
    /// </summary>
    public class ConfigTests
    {
        public ConfigTests()
        {
            // Reset to compiled defaults before each test to avoid cross-test pollution
            NmConfigProvider.ResetToDefaults();
        }

        [Fact]
        public void DefaultConfig_PassesValidation()
        {
            var config = new NmConfig();
            var messages = ConfigValidator.Validate(config);
            var errors = messages.Where(m => m.IsError).ToList();
            Assert.Empty(errors);
        }

        [Fact]
        public void DefaultConfig_HasExpectedWorkCenterRates()
        {
            var config = new NmConfig();
            Assert.Equal(120.0, config.WorkCenters.F115_LaserCutting);
            Assert.Equal(80.0, config.WorkCenters.F140_PressBrake);
            Assert.Equal(175.0, config.WorkCenters.F145_CncBending);
            Assert.Equal(42.0, config.WorkCenters.F210_Deburring);
            Assert.Equal(65.0, config.WorkCenters.F220_Tapping);
            Assert.Equal(44.0, config.WorkCenters.F300_MaterialHandling);
        }

        [Fact]
        public void Initialize_WithJsonFiles_LoadsLaserSpeedData()
        {
            // Look for config dir relative to test assembly (copied by csproj Content items)
            string asmDir = System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location);
            string configDir = System.IO.Path.Combine(asmDir, "config");

            if (!System.IO.Directory.Exists(configDir))
            {
                // Skip gracefully if JSON files aren't deployed
                return;
            }

            NmConfigProvider.Reload(configDir);

            Assert.NotNull(NmConfigProvider.LoadedConfigPath);
            Assert.NotNull(NmConfigProvider.LoadedTablesPath);

            // Verify laser speed data loaded
            var tables = NmConfigProvider.Tables;
            Assert.NotEmpty(tables.LaserSpeeds.StainlessSteel);
            Assert.NotEmpty(tables.LaserSpeeds.CarbonSteel);
            Assert.NotEmpty(tables.LaserSpeeds.Aluminum);
            Assert.Equal(28, tables.LaserSpeeds.StainlessSteel.Count);
            Assert.Equal(25, tables.LaserSpeeds.CarbonSteel.Count);
            Assert.Equal(24, tables.LaserSpeeds.Aluminum.Count);
        }

        [Fact]
        public void Initialize_WithJsonFiles_LoadsTubeCuttingData()
        {
            string asmDir = System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location);
            string configDir = System.IO.Path.Combine(asmDir, "config");

            if (!System.IO.Directory.Exists(configDir))
                return;

            NmConfigProvider.Reload(configDir);

            var tc = NmConfigProvider.Tables.TubeCutting;
            Assert.NotEmpty(tc.CarbonSteel.FeedRates);
            Assert.NotEmpty(tc.StainlessSteel.FeedRates);
            Assert.Equal(0.85, tc.CarbonSteel.FeedMultiplier);
            Assert.Equal(0.02, tc.CarbonSteel.KerfIn);
            Assert.Equal(0.03, tc.Aluminum.KerfIn);
        }

        [Fact]
        public void NmTablesProvider_GetLaserSpeed_ReturnsCorrectValues()
        {
            string asmDir = System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location);
            string configDir = System.IO.Path.Combine(asmDir, "config");

            if (!System.IO.Directory.Exists(configDir))
                return;

            NmConfigProvider.Reload(configDir);

            // Stainless steel, 0.048" gauge 18: feed=1600, pierce=0.1
            var (feed, pierce) = NmTablesProvider.GetLaserSpeed(
                NmConfigProvider.Tables, 0.048, "304L");
            Assert.Equal(1600, feed);
            Assert.Equal(0.1, pierce);

            // Carbon steel, 0.060" gauge 16: feed=1800, pierce=0.01
            var (feed2, pierce2) = NmTablesProvider.GetLaserSpeed(
                NmConfigProvider.Tables, 0.060, "A36");
            Assert.Equal(1800, feed2);
            Assert.Equal(0.01, pierce2);
        }

        [Fact]
        public void NmTablesProvider_GetTubeCuttingParams_ReturnsCorrectValues()
        {
            string asmDir = System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location);
            string configDir = System.IO.Path.Combine(asmDir, "config");

            if (!System.IO.Directory.Exists(configDir))
                return;

            NmConfigProvider.Reload(configDir);

            // Carbon steel, 0.045" wall: feedRate=295, multiplier=0.85, so cutSpeed=250.75
            var (cutSpeed, pierceSec, kerfIn) = NmTablesProvider.GetTubeCuttingParams(
                NmConfigProvider.Tables, "CarbonSteel", 0.045);
            Assert.Equal(295 * 0.85, cutSpeed);
            Assert.Equal(0.05, pierceSec);
            Assert.Equal(0.02, kerfIn);
        }

        [Fact]
        public void ResetToDefaults_ClearsLoadedPaths()
        {
            NmConfigProvider.ResetToDefaults();
            Assert.Null(NmConfigProvider.LoadedConfigPath);
            Assert.Null(NmConfigProvider.LoadedTablesPath);
            Assert.NotNull(NmConfigProvider.Current);
            Assert.NotNull(NmConfigProvider.Tables);
        }
    }
}
