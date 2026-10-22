using System.Globalization;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 3 coverage for the Shape Size Panel rules (QA9-BtnSizePanel-001): the point
    /// conversion, typed dimension parsing with suffixes and culture fallbacks, the typed-input
    /// validity, the step-adjust floor, the per-shape change detection and the uniform display
    /// rule, used by ShapeSizePanel.
    /// </summary>
    [TestClass]
    public sealed class ShapeSizePolicyTests
    {
        [TestMethod]
        public void ConversionUsesThePanelsPointsPerCentimetreConstant()
        {
            Assert.AreEqual(28.3464567, ShapeSizePolicy.CmToPoints(1d), 0.0001);
            Assert.AreEqual(1d, ShapeSizePolicy.PointsToCm(28.3464567d), 0.0001);
            Assert.AreEqual(10d * ShapeSizePolicy.PointsPerCentimetre, ShapeSizePolicy.CmToPoints(10d), 0.0001);
        }

        [TestMethod]
        public void TypedParsingStripsSuffixesAndConvertsPoints()
        {
            Assert.IsTrue(ShapeSizePolicy.TryParseDisplay("5", out double bare));
            Assert.AreEqual(5d, bare, 0.0001);

            Assert.IsTrue(ShapeSizePolicy.TryParseDisplay("5 cm", out double withCm));
            Assert.AreEqual(5d, withCm, 0.0001);

            Assert.IsTrue(ShapeSizePolicy.TryParseDisplay("141.73pt", out double withPt));
            Assert.AreEqual(5d, withPt, 0.01);
        }

        [TestMethod]
        public void TypedParsingFallsBackAcrossCultures()
        {
            var original = Thread.CurrentThread.CurrentCulture;
            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
                Assert.IsTrue(ShapeSizePolicy.TryParseDisplay("12,5", out double comma));
                Assert.AreEqual(12.5d, comma, 0.0001);

                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                Assert.IsTrue(ShapeSizePolicy.TryParseDisplay("12.5", out double dot));
                Assert.AreEqual(12.5d, dot, 0.0001);

                Assert.IsFalse(ShapeSizePolicy.TryParseDisplay("abc", out _), "Non-numeric input is rejected.");
                Assert.IsFalse(ShapeSizePolicy.TryParseDisplay("", out _), "Empty input is rejected.");
                Assert.IsFalse(ShapeSizePolicy.TryParseDisplay(null, out _), "Null input is rejected.");
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = original;
            }
        }

        [TestMethod]
        public void TypedValidityRequiresPositiveValues()
        {
            Assert.IsTrue(ShapeSizePolicy.IsValidTypedCm(0.5d));
            Assert.IsFalse(ShapeSizePolicy.IsValidTypedCm(0d), "Zero is rejected.");
            Assert.IsFalse(ShapeSizePolicy.IsValidTypedCm(-2d), "Negative dimensions are rejected.");
            Assert.IsFalse(ShapeSizePolicy.IsValidTypedCm(double.NaN), "NaN is rejected.");
        }

        [TestMethod]
        public void StepAdjustFloorsAtTheMinimum()
        {
            double min = ShapeSizePolicy.MinimumAdjustPoints;
            Assert.AreEqual(ShapeSizePolicy.CmToPoints(0.01d), min, 0.0001, "The floor is 0.01 cm.");

            Assert.AreEqual(min, ShapeSizePolicy.TargetAfterAdjust(10d, -5d), 0.0001, "Stepping below the floor clamps to it.");
            Assert.AreEqual(10d + ShapeSizePolicy.CmToPoints(0.1d), ShapeSizePolicy.TargetAfterAdjust(10d, 0.1d), 0.0001);
        }

        [TestMethod]
        public void ChangeDetectionSkipsInvalidAndNearEqualTargets()
        {
            Assert.IsFalse(ShapeSizePolicy.NeedsDimensionChange(10f, 0f), "Zero targets are skipped.");
            Assert.IsFalse(ShapeSizePolicy.NeedsDimensionChange(10f, float.NaN), "NaN targets are skipped.");
            Assert.IsFalse(ShapeSizePolicy.NeedsDimensionChange(10f, 10f), "Equal targets are skipped.");
            Assert.IsFalse(ShapeSizePolicy.NeedsDimensionChange(10f, 10.0005f), "Differences below the epsilon are skipped.");
            Assert.IsTrue(ShapeSizePolicy.NeedsDimensionChange(10f, 10.01f), "Meaningful differences are applied.");
        }

        [TestMethod]
        public void UniformDisplayUsesThePointTolerance()
        {
            Assert.IsTrue(ShapeSizePolicy.AreUniform(0f, 0.01f), "A difference of exactly the tolerance constant is uniform.");
            Assert.IsTrue(ShapeSizePolicy.AreUniform(100f, 100f), "Equal sizes are uniform.");
            Assert.IsFalse(ShapeSizePolicy.AreUniform(100f, 100.01f),
                "Float reality: 100.01f rounds above the 0.01f tolerance, so production shows 100 vs 100.01 as mixed.");
            Assert.IsFalse(ShapeSizePolicy.AreUniform(100f, 100.02f), "A difference above the tolerance is mixed.");
        }
    }
}
