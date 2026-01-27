using System;
using System.Runtime.InteropServices;
using Xunit;
using Xunit.Abstractions;

namespace NM.Core.Tests.Integration
{
    /// <summary>
    /// Base class for integration tests that require SolidWorks.
    /// Tests are skipped if SW is not running.
    /// </summary>
    public abstract class SwTestBase : IDisposable
    {
        protected dynamic SwApp { get; private set; }
        protected bool IsSwAvailable { get; private set; }

        protected SwTestBase()
        {
            try
            {
                // Try to connect to running SolidWorks instance
                SwApp = Marshal.GetActiveObject("SldWorks.Application");
                IsSwAvailable = SwApp != null;
            }
            catch (COMException)
            {
                // SolidWorks not running - tests will be skipped
                IsSwAvailable = false;
                SwApp = null;
            }
        }

        /// <summary>
        /// Check if SolidWorks is available for integration tests.
        /// Tests should check IsSwAvailable and return early if false.
        /// </summary>
        protected bool ShouldSkipSwTest(Xunit.Abstractions.ITestOutputHelper output)
        {
            if (!IsSwAvailable)
            {
                output?.WriteLine("SKIPPED: SolidWorks is not running");
                return true;
            }
            return false;
        }

        /// <summary>
        /// Gets the active document, or null if none open.
        /// </summary>
        protected dynamic GetActiveDocument()
        {
            if (!IsSwAvailable) return null;
            return SwApp?.ActiveDoc;
        }

        public void Dispose()
        {
            // Don't release SwApp - we're connecting to existing instance
            // Let SW manage its own lifetime
            SwApp = null;
        }
    }
}
