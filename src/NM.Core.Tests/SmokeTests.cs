using Xunit;

namespace NM.Core.Tests
{
    /// <summary>
    /// Smoke tests to verify the test infrastructure works.
    /// </summary>
    public class SmokeTests
    {
        [Fact]
        public void TestInfrastructure_Works()
        {
            // Simple test to verify xUnit runs
            Assert.True(true);
        }

        [Fact]
        public void CanReference_MainProject()
        {
            // Verify we can access types from the main project
            // ErrorHandler exists in NM.Core namespace
            NM.Core.ErrorHandler.PushCallStack("TestMethod");
            NM.Core.ErrorHandler.PopCallStack();
            // If we got here without exception, the reference works
            Assert.True(true);
        }

        [Theory]
        [InlineData(1, 2, 3)]
        [InlineData(0, 0, 0)]
        [InlineData(-1, 1, 0)]
        public void Math_Addition_Works(int a, int b, int expected)
        {
            Assert.Equal(expected, a + b);
        }
    }
}
