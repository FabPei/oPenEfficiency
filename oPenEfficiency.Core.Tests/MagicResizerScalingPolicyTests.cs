using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;
using static oPenEfficiency.Core.Services.MagicResizerScalingPolicy;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 align-and-size-wave coverage for Scale Selected Shapes
    /// (QA9-BtnMagicResizer-002): the scale-factor derivations, the nine anchor fractions
    /// and the anchor-preserving per-shape bounds used by MagicResizerWindow.
    /// </summary>
    [TestClass]
    public sealed class MagicResizerScalingPolicyTests
    {
        // Reference rectangle: left 200, top 100, 80x60.
        private const float L = 200f, T = 100f, W = 80f, H = 60f;

        [TestMethod]
        public void PercentageModeScalesBothAxesByTheSameFactor()
        {
            Assert.IsTrue(TryPercentageScaleFactors(150.0, out double scaleX, out double scaleY));
            Assert.AreEqual(1.5, scaleX, 1e-12);
            Assert.AreEqual(1.5, scaleY, 1e-12);
        }

        [TestMethod]
        public void PercentageModeRejectsNonPositiveValues()
        {
            Assert.IsFalse(TryPercentageScaleFactors(0.0, out _, out _));
            Assert.IsFalse(TryPercentageScaleFactors(-25.0, out _, out _));
        }

        [TestMethod]
        public void AbsoluteModeDerivesTheFactorsFromTheReferenceShapeInPoints()
        {
            // 5.08 cm = 144 pt; a 72 pt wide reference therefore doubles on X.
            // 2.54 cm = 72 pt; a 60 pt high reference scales by 1.2 on Y.
            Assert.IsTrue(TryAbsoluteScaleFactors(5.08, 2.54, 72f, 60f, out double scaleX, out double scaleY));
            Assert.AreEqual(2.0, scaleX, 1e-9);
            Assert.AreEqual(1.2, scaleY, 1e-9);
        }

        [TestMethod]
        public void AbsoluteModeRejectsNonPositiveTargetsAndReferenceDimensions()
        {
            Assert.IsFalse(TryAbsoluteScaleFactors(0.0, 2.54, 72f, 60f, out _, out _));
            Assert.IsFalse(TryAbsoluteScaleFactors(5.08, -1.0, 72f, 60f, out _, out _));
            Assert.IsFalse(TryAbsoluteScaleFactors(5.08, 2.54, 0f, 60f, out _, out _));
            Assert.IsFalse(TryAbsoluteScaleFactors(5.08, 2.54, 72f, -0.5f, out _, out _));
        }

        [TestMethod]
        public void TheCentimetrePointConversionsAreMutuallyInverse()
        {
            Assert.AreEqual(2.54, PointsToCentimetres(72f), 1e-12);
            Assert.AreEqual(72.0, CentimetresToPoints(2.54), 1e-12);
            Assert.AreEqual(100.0, CentimetresToPoints(PointsToCentimetres(100f)), 1e-9);
        }

        [TestMethod]
        public void AllNineAnchorFractionsMatchTheThreeByThreeGrid()
        {
            void AssertFractions(ScaleAnchor anchor, double x, double y)
            {
                GetAnchorFractions(anchor, out double anchorX, out double anchorY);
                Assert.AreEqual(x, anchorX, 1e-12, anchor.ToString());
                Assert.AreEqual(y, anchorY, 1e-12, anchor.ToString());
            }

            AssertFractions(ScaleAnchor.TopLeft, 0.0, 0.0);
            AssertFractions(ScaleAnchor.Top, 0.5, 0.0);
            AssertFractions(ScaleAnchor.TopRight, 1.0, 0.0);
            AssertFractions(ScaleAnchor.Left, 0.0, 0.5);
            AssertFractions(ScaleAnchor.Center, 0.5, 0.5);
            AssertFractions(ScaleAnchor.Right, 1.0, 0.5);
            AssertFractions(ScaleAnchor.BottomLeft, 0.0, 1.0);
            AssertFractions(ScaleAnchor.Bottom, 0.5, 1.0);
            AssertFractions(ScaleAnchor.BottomRight, 1.0, 1.0);
        }

        [TestMethod]
        public void CenterAnchorKeepsTheCenterFixed()
        {
            ComputeScaledBounds(L, T, W, H, 2.0, 2.0, ScaleAnchor.Center,
                out float newLeft, out float newTop, out float newWidth, out float newHeight);

            Assert.AreEqual(160f, newLeft, 0.001);   // 240 - 160 * 0.5
            Assert.AreEqual(70f, newTop, 0.001);     // 130 - 120 * 0.5
            Assert.AreEqual(160f, newWidth, 0.001);
            Assert.AreEqual(120f, newHeight, 0.001);
        }

        [TestMethod]
        public void TopLeftAnchorKeepsTheTopLeftCornerFixed()
        {
            ComputeScaledBounds(L, T, W, H, 0.5, 0.5, ScaleAnchor.TopLeft,
                out float newLeft, out float newTop, out float newWidth, out float newHeight);

            Assert.AreEqual(L, newLeft, 0.001);
            Assert.AreEqual(T, newTop, 0.001);
            Assert.AreEqual(40f, newWidth, 0.001);
            Assert.AreEqual(30f, newHeight, 0.001);
        }

        [TestMethod]
        public void TheAnchorPointStaysFixedForEveryAnchor()
        {
            foreach (ScaleAnchor anchor in System.Enum.GetValues(typeof(ScaleAnchor)))
            {
                GetAnchorFractions(anchor, out double anchorX, out double anchorY);
                double fixedX = L + (W * anchorX);
                double fixedY = T + (H * anchorY);

                ComputeScaledBounds(L, T, W, H, 1.5, 0.75, anchor,
                    out float newLeft, out float newTop, out float newWidth, out float newHeight);

                Assert.AreEqual(fixedX, newLeft + (newWidth * anchorX), 0.001, anchor.ToString());
                Assert.AreEqual(fixedY, newTop + (newHeight * anchorY), 0.001, anchor.ToString());
            }
        }

        [TestMethod]
        public void NoOpDetectionRequiresBothFactorsAtExactlyOne()
        {
            Assert.IsTrue(IsNoOp(1.0, 1.0));
            Assert.IsFalse(IsNoOp(1.0, 1.0000001));
            Assert.IsFalse(IsNoOp(0.5, 1.0));
        }

        [TestMethod]
        public void TheScalarFactorIsTheAverageOfBothAxes()
        {
            Assert.AreEqual(1.5, ScalarFactor(2.0, 1.0), 1e-12);
            Assert.AreEqual(2.0, ScaleScalar(4f, ScalarFactor(0.5, 0.5)), 1e-6);
        }

        [TestMethod]
        public void TheScaledFontSizeIsFlooredAtOnePoint()
        {
            Assert.AreEqual(30f, ScaledFontSize(20f, 1.5), 1e-6);
            Assert.AreEqual(1f, ScaledFontSize(10f, 0.05), 1e-6);
        }
    }
}
