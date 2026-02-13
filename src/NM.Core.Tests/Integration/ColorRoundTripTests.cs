using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Xunit;
using Xunit.Abstractions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.Core.Tests.Integration
{
    /// <summary>
    /// Integration tests that validate the color save/apply/restore round-trip
    /// against a live SolidWorks instance. Skips gracefully if SW not running.
    ///
    /// These tests open a sample part, record its original colors, apply a highlight,
    /// verify the highlight, restore original colors, and verify the restoration.
    /// All color values are logged for visual inspection.
    /// </summary>
    public class ColorRoundTripTests : SwTestBase
    {
        private readonly ITestOutputHelper _output;

        // Test highlight color: bright red
        private static readonly double[] HighlightRed = { 1.0, 0.0, 0.0, 0.5, 1.0, 0.5, 0.3, 0.0, 0.0 };

        // Test inputs directory (relative to test execution dir)
        private const string TestInputDir = @"..\..\..\..\tests\GoldStandard_Inputs";

        public ColorRoundTripTests(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void PartBodyColor_SaveApplyRestore_RoundTrip()
        {
            if (ShouldSkipSwTest(_output)) return;

            var swApp = (ISldWorks)SwApp;
            var model = swApp.ActiveDoc as IModelDoc2;
            if (model == null)
            {
                _output.WriteLine("SKIPPED: No active document — open a part in SolidWorks first");
                return;
            }

            if (model.GetType() != (int)swDocumentTypes_e.swDocPART)
            {
                _output.WriteLine("SKIPPED: Active document is not a part — open a part file first");
                return;
            }

            var part = (IPartDoc)model;
            var bodiesRaw = part.GetBodies2((int)swBodyType_e.swSolidBody, true);
            if (bodiesRaw == null || ((object[])bodiesRaw).Length == 0)
            {
                _output.WriteLine("SKIPPED: Part has no solid bodies");
                return;
            }

            var body = (IBody2)((object[])bodiesRaw)[0];
            _output.WriteLine($"Testing on: {model.GetTitle()}");
            _output.WriteLine($"Body: {body.Name}");
            _output.WriteLine("---");

            // Step 1: Read and log original body color
            var originalBodyColor = (double[])body.MaterialPropertyValues2;
            bool hadExplicitBodyColor = originalBodyColor != null;

            _output.WriteLine("STEP 1: Original body color");
            if (hadExplicitBodyColor)
            {
                LogColorArray("  Body color", originalBodyColor);
                _output.WriteLine("  (Component has EXPLICIT body-level color)");
            }
            else
            {
                _output.WriteLine("  Body color: null (inheriting from part)");
                // Read model-level color as reference
                var modelColor = (double[])model.MaterialPropertyValues;
                if (modelColor != null)
                    LogColorArray("  Model-level color (inherited)", modelColor);
            }

            // Step 2: Apply highlight color
            _output.WriteLine("");
            _output.WriteLine("STEP 2: Apply red highlight");
            body.MaterialPropertyValues2 = (double[])HighlightRed.Clone();
            model.GraphicsRedraw2();

            var afterApply = (double[])body.MaterialPropertyValues2;
            Assert.NotNull(afterApply);
            LogColorArray("  Body color after highlight", afterApply);
            Assert.Equal(1.0, afterApply[0], 2); // Red = 1.0
            Assert.Equal(0.0, afterApply[1], 2); // Green = 0.0
            Assert.Equal(0.0, afterApply[2], 2); // Blue = 0.0
            _output.WriteLine("  VERIFIED: Red highlight applied correctly");

            // Step 3: Restore original color
            _output.WriteLine("");
            _output.WriteLine("STEP 3: Restore original color");
            if (hadExplicitBodyColor)
            {
                body.MaterialPropertyValues2 = (double[])originalBodyColor.Clone();
                _output.WriteLine("  Method: SetMaterialPropertyValues2 (restoring explicit color)");
            }
            else
            {
                body.RemoveMaterialProperty(
                    (int)swInConfigurationOpts_e.swThisConfiguration, null);
                _output.WriteLine("  Method: RemoveMaterialProperty (removing override, reverting to inheritance)");
            }
            model.GraphicsRedraw2();

            // Step 4: Verify restoration
            _output.WriteLine("");
            _output.WriteLine("STEP 4: Verify restoration");
            var afterRestore = (double[])body.MaterialPropertyValues2;

            if (hadExplicitBodyColor)
            {
                Assert.NotNull(afterRestore);
                LogColorArray("  Restored body color", afterRestore);
                // Compare RGB values
                Assert.Equal(originalBodyColor[0], afterRestore[0], 3);
                Assert.Equal(originalBodyColor[1], afterRestore[1], 3);
                Assert.Equal(originalBodyColor[2], afterRestore[2], 3);
                _output.WriteLine("  VERIFIED: Original explicit color restored");
            }
            else
            {
                // After RemoveMaterialProperty, body should return null (no override)
                if (afterRestore == null)
                {
                    _output.WriteLine("  Body color: null (correctly inheriting from part again)");
                    _output.WriteLine("  VERIFIED: Override removed, inheritance restored");
                }
                else
                {
                    // Some SW versions may still return the part color
                    LogColorArray("  Body color after restore", afterRestore);
                    // Verify it's NOT still red
                    bool isStillRed = afterRestore[0] > 0.9 && afterRestore[1] < 0.1 && afterRestore[2] < 0.1;
                    Assert.False(isStillRed, "Body should NOT still be red after restore");
                    // Verify lighting values are reasonable (not all zeros = not black)
                    bool isBlack = afterRestore[3] < 0.01 && afterRestore[4] < 0.01 && afterRestore[5] < 0.01;
                    Assert.False(isBlack, "Body should NOT have all-zero lighting (black) after restore");
                    _output.WriteLine("  VERIFIED: Color is not red and not black");
                }
            }

            _output.WriteLine("");
            _output.WriteLine("=== Color round-trip test PASSED ===");
        }

        [Fact]
        public void ComponentColor_SaveApplyRestore_RoundTrip()
        {
            if (ShouldSkipSwTest(_output)) return;

            var swApp = (ISldWorks)SwApp;
            var model = swApp.ActiveDoc as IModelDoc2;
            if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                _output.WriteLine("SKIPPED: No assembly open — open an assembly in SolidWorks first");
                return;
            }

            var config = model.ConfigurationManager?.ActiveConfiguration;
            if (config == null)
            {
                _output.WriteLine("SKIPPED: No active configuration");
                return;
            }

            var root = config.GetRootComponent3(true) as IComponent2;
            if (root == null)
            {
                _output.WriteLine("SKIPPED: Could not get root component");
                return;
            }

            var childrenRaw = root.GetChildren() as object[];
            if (childrenRaw == null || childrenRaw.Length == 0)
            {
                _output.WriteLine("SKIPPED: Assembly has no children");
                return;
            }

            // Test on first non-suppressed component
            IComponent2 testComp = null;
            foreach (IComponent2 child in childrenRaw)
            {
                if (child.GetSuppression2() != (int)swComponentSuppressionState_e.swComponentSuppressed)
                {
                    testComp = child;
                    break;
                }
            }
            if (testComp == null)
            {
                _output.WriteLine("SKIPPED: All components are suppressed");
                return;
            }

            _output.WriteLine($"Assembly: {model.GetTitle()}");
            _output.WriteLine($"Test component: {testComp.Name2}");
            _output.WriteLine($"Component path: {testComp.GetPathName()}");
            _output.WriteLine("---");

            // Step 1: Read original component color
            var originalColor = (double[])testComp.GetMaterialPropertyValues2(
                (int)swInConfigurationOpts_e.swThisConfiguration, null);
            bool hadExplicitColor = originalColor != null;

            _output.WriteLine("STEP 1: Original component color");
            if (hadExplicitColor)
            {
                LogColorArray("  Component color", originalColor);
                bool hasValidLighting = originalColor.Length >= 9 &&
                    (originalColor[3] > 0.001 || originalColor[4] > 0.001 || originalColor[5] > 0.001);
                _output.WriteLine($"  Has valid lighting: {hasValidLighting}");
            }
            else
            {
                _output.WriteLine("  Component color: null (inheriting from part document)");
            }

            // Step 2: Apply red highlight
            _output.WriteLine("");
            _output.WriteLine("STEP 2: Apply red highlight");
            testComp.SetMaterialPropertyValues2(HighlightRed,
                (int)swInConfigurationOpts_e.swThisConfiguration, null);
            model.GraphicsRedraw2();

            var afterApply = (double[])testComp.GetMaterialPropertyValues2(
                (int)swInConfigurationOpts_e.swThisConfiguration, null);
            Assert.NotNull(afterApply);
            LogColorArray("  Component color after highlight", afterApply);
            Assert.Equal(1.0, afterApply[0], 2); // Red
            _output.WriteLine("  VERIFIED: Red highlight applied");

            // Step 3: Restore using the correct method
            _output.WriteLine("");
            _output.WriteLine("STEP 3: Restore original color");
            bool shouldRemove = !hadExplicitColor ||
                (originalColor != null && originalColor.Length >= 9 &&
                 originalColor[3] < 0.001 && originalColor[4] < 0.001 && originalColor[5] < 0.001);

            if (!shouldRemove && originalColor != null)
            {
                testComp.SetMaterialPropertyValues2(originalColor,
                    (int)swInConfigurationOpts_e.swThisConfiguration, null);
                _output.WriteLine("  Method: SetMaterialPropertyValues2 (restoring explicit color)");
            }
            else
            {
                testComp.RemoveMaterialProperty2(
                    (int)swInConfigurationOpts_e.swThisConfiguration, null);
                _output.WriteLine("  Method: RemoveMaterialProperty2 (removing override)");
            }
            model.GraphicsRedraw2();

            // Step 4: Verify restoration
            _output.WriteLine("");
            _output.WriteLine("STEP 4: Verify restoration");
            var afterRestore = (double[])testComp.GetMaterialPropertyValues2(
                (int)swInConfigurationOpts_e.swThisConfiguration, null);

            if (hadExplicitColor && !shouldRemove)
            {
                Assert.NotNull(afterRestore);
                LogColorArray("  Restored color", afterRestore);
                Assert.Equal(originalColor[0], afterRestore[0], 3);
                Assert.Equal(originalColor[1], afterRestore[1], 3);
                Assert.Equal(originalColor[2], afterRestore[2], 3);
                _output.WriteLine("  VERIFIED: Original color restored");
            }
            else
            {
                if (afterRestore == null)
                {
                    _output.WriteLine("  Component color: null (correctly inheriting again)");
                    _output.WriteLine("  VERIFIED: Override removed successfully");
                }
                else
                {
                    LogColorArray("  Component color after restore", afterRestore);
                    bool isStillRed = afterRestore[0] > 0.9 && afterRestore[1] < 0.1 && afterRestore[2] < 0.1;
                    Assert.False(isStillRed, "Component should NOT still be red after restore");
                    _output.WriteLine("  VERIFIED: Not red (restore worked)");
                }
            }

            _output.WriteLine("");
            _output.WriteLine("=== Component color round-trip test PASSED ===");
        }

        [Fact]
        public void BugRepro_NullSavedColor_ShouldNotTurnBlack()
        {
            // This test specifically reproduces the reported bug:
            // 1. Component has no explicit color (GetMaterialPropertyValues2 returns null)
            // 2. Apply red highlight
            // 3. Restore by calling RemoveMaterialProperty2 (not SetMaterialPropertyValues2(null))
            // 4. Verify component is NOT black

            if (ShouldSkipSwTest(_output)) return;

            var swApp = (ISldWorks)SwApp;
            var model = swApp.ActiveDoc as IModelDoc2;
            if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                _output.WriteLine("SKIPPED: Need an assembly open in SolidWorks");
                return;
            }

            var config = model.ConfigurationManager?.ActiveConfiguration;
            var root = config?.GetRootComponent3(true) as IComponent2;
            var childrenRaw = root?.GetChildren() as object[];
            if (childrenRaw == null || childrenRaw.Length == 0)
            {
                _output.WriteLine("SKIPPED: No components found");
                return;
            }

            // Find a component with no explicit color (null from GetMaterialPropertyValues2)
            IComponent2 testComp = null;
            foreach (IComponent2 child in childrenRaw)
            {
                if (child.GetSuppression2() == (int)swComponentSuppressionState_e.swComponentSuppressed)
                    continue;

                var color = (double[])child.GetMaterialPropertyValues2(
                    (int)swInConfigurationOpts_e.swThisConfiguration, null);
                if (color == null)
                {
                    testComp = child;
                    break;
                }
            }

            if (testComp == null)
            {
                _output.WriteLine("SKIPPED: All components have explicit colors — cannot reproduce null-save scenario");
                _output.WriteLine("(To test this, ensure at least one component inherits its color from its part file)");
                return;
            }

            _output.WriteLine($"=== Bug Repro: Null saved color should NOT turn black ===");
            _output.WriteLine($"Component: {testComp.Name2}");
            _output.WriteLine($"Original color: null (inheriting from part)");

            // Apply red
            testComp.SetMaterialPropertyValues2(HighlightRed,
                (int)swInConfigurationOpts_e.swThisConfiguration, null);
            model.GraphicsRedraw2();
            _output.WriteLine("Applied red highlight");

            // THE FIX: Use RemoveMaterialProperty2 instead of doing nothing or setting null
            testComp.RemoveMaterialProperty2(
                (int)swInConfigurationOpts_e.swThisConfiguration, null);
            model.GraphicsRedraw2();
            _output.WriteLine("Called RemoveMaterialProperty2 to remove override");

            // Verify NOT red and NOT black
            var afterRestore = (double[])testComp.GetMaterialPropertyValues2(
                (int)swInConfigurationOpts_e.swThisConfiguration, null);

            if (afterRestore == null)
            {
                _output.WriteLine("RESULT: null (correctly inheriting from part again)");
                _output.WriteLine("PASS: Component is back to its natural appearance");
            }
            else
            {
                LogColorArray("RESULT", afterRestore);
                bool isRed = afterRestore[0] > 0.9 && afterRestore[1] < 0.1 && afterRestore[2] < 0.1;
                bool isBlack = afterRestore.Length >= 6 &&
                    afterRestore[0] < 0.01 && afterRestore[1] < 0.01 && afterRestore[2] < 0.01 &&
                    afterRestore[3] < 0.01 && afterRestore[4] < 0.01 && afterRestore[5] < 0.01;

                Assert.False(isRed, "BUG: Component is still red after RemoveMaterialProperty2");
                Assert.False(isBlack, "BUG: Component turned black (all-zero lighting values)");
                _output.WriteLine("PASS: Component color is not red and not black");
            }

            _output.WriteLine("=== Bug repro test PASSED ===");
        }

        #region Helpers

        private void LogColorArray(string label, double[] color)
        {
            if (color == null)
            {
                _output.WriteLine($"{label}: null");
                return;
            }

            if (color.Length >= 9)
            {
                _output.WriteLine($"{label}: RGB({color[0]:F3}, {color[1]:F3}, {color[2]:F3}) " +
                    $"Ambient={color[3]:F3} Diffuse={color[4]:F3} Specular={color[5]:F3} " +
                    $"Shininess={color[6]:F3} Transparency={color[7]:F3} Emission={color[8]:F3}");
            }
            else
            {
                _output.WriteLine($"{label}: [{string.Join(", ", color.Select(v => v.ToString("F3")))}] (length={color.Length})");
            }
        }

        #endregion
    }
}
