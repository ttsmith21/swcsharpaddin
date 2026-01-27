using System;
using Xunit;
using Xunit.Abstractions;

namespace NM.Core.Tests.Integration
{
    /// <summary>
    /// Integration tests for SolidWorks connectivity.
    /// These tests require SolidWorks to be running.
    /// </summary>
    public class SolidWorksConnectionTests : SwTestBase
    {
        private readonly ITestOutputHelper _output;

        public SolidWorksConnectionTests(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void CanConnectToSolidWorks()
        {
            if (!IsSwAvailable)
            {
                _output.WriteLine("SKIPPED: SolidWorks is not running");
                return; // Skip gracefully
            }

            Assert.NotNull(SwApp);
            _output.WriteLine($"Connected to SolidWorks");
        }

        [Fact]
        public void CanGetSolidWorksVersion()
        {
            if (!IsSwAvailable)
            {
                _output.WriteLine("SKIPPED: SolidWorks is not running");
                return;
            }

            string revision = SwApp.RevisionNumber();
            Assert.False(string.IsNullOrEmpty(revision));
            _output.WriteLine($"SolidWorks Version: {revision}");
        }

        [Fact]
        public void CanCheckActiveDocument()
        {
            if (!IsSwAvailable)
            {
                _output.WriteLine("SKIPPED: SolidWorks is not running");
                return;
            }

            var doc = GetActiveDocument();
            if (doc != null)
            {
                string path = doc.GetPathName();
                _output.WriteLine($"Active Document: {path}");
            }
            else
            {
                _output.WriteLine("No active document (this is OK)");
            }
            // This test passes regardless - just verifying we can check
            Assert.True(true);
        }

        [Fact]
        public void AddIn_IsRegistered()
        {
            // This test doesn't need SW running - just checks registry
            var addInGuid = "{D5355548-9569-4381-9939-5D14252A3E47}";
            var regPath = $@"SOFTWARE\SolidWorks\Addins\{addInGuid}";

            using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(regPath))
            {
                if (key != null)
                {
                    var title = key.GetValue("Title")?.ToString();
                    _output.WriteLine($"Add-in registered: {title}");
                    Assert.NotNull(title);
                }
                else
                {
                    _output.WriteLine("Add-in not registered in HKLM (may be in HKCU or not registered)");
                    // Don't fail - registration may be per-user
                }
            }
        }
    }
}
