using System;
using Xunit;
using Xunit.Abstractions;
using NM.SwAddin.UI;

namespace NM.Core.Tests
{
    /// <summary>
    /// Unit tests for ProblemPartColorizer color save/restore logic.
    /// Validates the HasValidLighting check that prevents the "black on restore" bug.
    /// </summary>
    public class ProblemPartColorizerTests
    {
        private readonly ITestOutputHelper _output;

        public ProblemPartColorizerTests(ITestOutputHelper output)
        {
            _output = output;
        }

        #region HasValidLighting Tests

        [Fact]
        public void HasValidLighting_NullArray_ReturnsFalse()
        {
            Assert.False(ProblemPartColorizer.HasValidLighting(null));
        }

        [Fact]
        public void HasValidLighting_EmptyArray_ReturnsFalse()
        {
            Assert.False(ProblemPartColorizer.HasValidLighting(new double[0]));
        }

        [Fact]
        public void HasValidLighting_TooShortArray_ReturnsFalse()
        {
            // Only 5 elements, need at least 9
            Assert.False(ProblemPartColorizer.HasValidLighting(new double[] { 1, 0, 0, 0.5, 1.0 }));
        }

        [Fact]
        public void HasValidLighting_AllZeros_ReturnsFalse()
        {
            // This is the "black on restore" case: all lighting values are zero
            var allZeros = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            bool result = ProblemPartColorizer.HasValidLighting(allZeros);
            Assert.False(result);
            _output.WriteLine("All-zero lighting correctly rejected (would render black)");
        }

        [Fact]
        public void HasValidLighting_ZeroRgbButValidLighting_ReturnsTrue()
        {
            // Black color (RGB=0,0,0) but with valid lighting — this is a legitimate black appearance
            var blackWithLighting = new double[] { 0, 0, 0, 0.5, 1.0, 0.5, 0.3, 0, 0 };
            Assert.True(ProblemPartColorizer.HasValidLighting(blackWithLighting));
        }

        [Fact]
        public void HasValidLighting_ValidRedHighlight_ReturnsTrue()
        {
            // The red highlight color used by the colorizer
            var red = new double[] { 1.0, 0.0, 0.0, 0.5, 1.0, 0.5, 0.3, 0.0, 0.0 };
            Assert.True(ProblemPartColorizer.HasValidLighting(red));
        }

        [Fact]
        public void HasValidLighting_TypicalSwDefaultColor_ReturnsTrue()
        {
            // Typical SolidWorks default part color (gray, moderate lighting)
            var swDefault = new double[] { 0.78, 0.78, 0.78, 0.4, 0.8, 0.5, 0.4, 0.0, 0.0 };
            Assert.True(ProblemPartColorizer.HasValidLighting(swDefault));
            _output.WriteLine("Typical SW default color accepted");
        }

        [Fact]
        public void HasValidLighting_OnlyAmbientNonZero_ReturnsTrue()
        {
            var ambientOnly = new double[] { 0.5, 0.5, 0.5, 0.3, 0, 0, 0, 0, 0 };
            Assert.True(ProblemPartColorizer.HasValidLighting(ambientOnly));
        }

        [Fact]
        public void HasValidLighting_OnlyDiffuseNonZero_ReturnsTrue()
        {
            var diffuseOnly = new double[] { 0.5, 0.5, 0.5, 0, 0.8, 0, 0, 0, 0 };
            Assert.True(ProblemPartColorizer.HasValidLighting(diffuseOnly));
        }

        [Fact]
        public void HasValidLighting_OnlySpecularNonZero_ReturnsTrue()
        {
            var specularOnly = new double[] { 0.5, 0.5, 0.5, 0, 0, 0.5, 0, 0, 0 };
            Assert.True(ProblemPartColorizer.HasValidLighting(specularOnly));
        }

        [Fact]
        public void HasValidLighting_NearZeroLighting_ReturnsFalse()
        {
            // Values just barely above zero but below threshold
            var nearZero = new double[] { 0.5, 0.5, 0.5, 0.0005, 0.0005, 0.0005, 0, 0, 0 };
            Assert.False(ProblemPartColorizer.HasValidLighting(nearZero));
            _output.WriteLine("Near-zero lighting correctly rejected");
        }

        #endregion

        #region Color Array Format Tests

        [Fact]
        public void RedHighlightColor_HasCorrectFormat()
        {
            // Verify the red highlight color array matches expected format
            // [R, G, B, Ambient, Diffuse, Specular, Shininess, Transparency, Emission]
            var red = new double[] { 1.0, 0.0, 0.0, 0.5, 1.0, 0.5, 0.3, 0.0, 0.0 };
            Assert.Equal(9, red.Length);
            Assert.Equal(1.0, red[0]); // R
            Assert.Equal(0.0, red[1]); // G
            Assert.Equal(0.0, red[2]); // B
            Assert.True(red[3] > 0, "Ambient must be > 0 to avoid black");
            Assert.True(red[4] > 0, "Diffuse must be > 0 to avoid black");
        }

        [Fact]
        public void ColorArrayClone_DoesNotShareReference()
        {
            var original = new double[] { 1.0, 0.0, 0.0, 0.5, 1.0, 0.5, 0.3, 0.0, 0.0 };
            var cloned = (double[])original.Clone();

            // Modify clone
            cloned[0] = 0.0;
            cloned[1] = 1.0;

            // Original should be unchanged
            Assert.Equal(1.0, original[0]);
            Assert.Equal(0.0, original[1]);
        }

        #endregion

        #region Restore Behavior Validation

        [Fact]
        public void RestoreScenario_NullSaved_ShouldUseRemoveMaterialProperty()
        {
            // Document the expected behavior:
            // When saved appearance is null, the component had no explicit color override.
            // On restore, we must call RemoveMaterialProperty2() instead of SetMaterialPropertyValues2().
            // Setting null or all-zeros would cause black appearance.
            _output.WriteLine("=== Restore Scenario: Component with no explicit color ===");
            _output.WriteLine("Original: GetMaterialPropertyValues2() returned null (inheriting from part)");
            _output.WriteLine("Action:   Applied red highlight via SetMaterialPropertyValues2()");
            _output.WriteLine("Restore:  Call RemoveMaterialProperty2() to strip component-level override");
            _output.WriteLine("Expected: Component returns to inherited part color (NOT black)");

            // Verify our logic: null saved → use RemoveMaterialProperty path
            double[] saved = null;
            bool shouldRemoveOverride = saved == null || !ProblemPartColorizer.HasValidLighting(saved);
            Assert.True(shouldRemoveOverride, "Null saved appearance should trigger RemoveMaterialProperty2");
        }

        [Fact]
        public void RestoreScenario_ValidSaved_ShouldRestoreValues()
        {
            // When saved appearance has valid lighting, restore it directly.
            _output.WriteLine("=== Restore Scenario: Component with explicit color ===");
            _output.WriteLine("Original: GetMaterialPropertyValues2() returned valid color array");
            _output.WriteLine("Action:   Applied red highlight");
            _output.WriteLine("Restore:  Call SetMaterialPropertyValues2(savedValues)");
            _output.WriteLine("Expected: Component returns to its original explicit color");

            var saved = new double[] { 0.2, 0.5, 0.8, 0.4, 0.8, 0.5, 0.4, 0.0, 0.0 };
            bool shouldRemoveOverride = saved == null || !ProblemPartColorizer.HasValidLighting(saved);
            Assert.False(shouldRemoveOverride, "Valid saved appearance should restore via SetMaterialPropertyValues2");
        }

        [Fact]
        public void RestoreScenario_AllZerosSaved_ShouldUseRemoveMaterialProperty()
        {
            // Some SW API versions return all-zeros for "default" state.
            // Restoring all-zeros would render black. Use RemoveMaterialProperty2 instead.
            _output.WriteLine("=== Restore Scenario: Component returned all-zeros ===");
            _output.WriteLine("Original: GetMaterialPropertyValues2() returned {0,0,0,0,0,0,0,0,0}");
            _output.WriteLine("Restore:  Call RemoveMaterialProperty2() (avoids black rendering)");

            var saved = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            bool shouldRemoveOverride = saved == null || !ProblemPartColorizer.HasValidLighting(saved);
            Assert.True(shouldRemoveOverride, "All-zeros saved should trigger RemoveMaterialProperty2 to avoid black");
        }

        #endregion
    }
}
