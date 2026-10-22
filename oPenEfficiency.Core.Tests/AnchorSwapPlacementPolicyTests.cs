using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 arrange-wave coverage for Swap Positions (QA9-BtnSwapPositions-002): the nine
    /// anchor points and the anchor-based position exchange, used by SwapPositionsFeature.
    /// </summary>
    [TestClass]
    public sealed class AnchorSwapPlacementPolicyTests
    {
        private const float L1 = 100f, T1 = 50f, W1 = 60f, H1 = 40f;   // shape 1
        private const float L2 = 400f, T2 = 300f, W2 = 80f, H2 = 60f;  // shape 2

        [TestMethod]
        public void CenterSwapExchangesTheAnchorPositions()
        {
            AnchorSwapPlacementPolicy.SwapPlacement(AnchorSwapPlacementPolicy.SwapAnchor.Center,
                L1, T1, W1, H1, L2, T2, W2, H2,
                out float newLeft1, out float newTop1, out float newLeft2, out float newTop2);
            // Shape 1's center (130,70) moves to shape 2's center (440,330):
            // offset was (30,20), so new top-left = (410,310).
            Assert.AreEqual(410f, newLeft1, 0.001);
            Assert.AreEqual(310f, newTop1, 0.001);
            // Shape 2's center (440,330) moves to (130,70): offset (40,30) -> (90,40).
            Assert.AreEqual(90f, newLeft2, 0.001);
            Assert.AreEqual(40f, newTop2, 0.001);
        }

        [TestMethod]
        public void EqualSizesSwappingCentersExchangeExactTopLefts()
        {
            AnchorSwapPlacementPolicy.SwapPlacement(AnchorSwapPlacementPolicy.SwapAnchor.Center,
                L1, T1, W1, H1, 400f, 300f, W1, H1,
                out float newLeft1, out float newTop1, out float newLeft2, out float newTop2);
            Assert.AreEqual(400f, newLeft1, 0.001);
            Assert.AreEqual(300f, newTop1, 0.001);
            Assert.AreEqual(100f, newLeft2, 0.001);
            Assert.AreEqual(50f, newTop2, 0.001);
        }

        [TestMethod]
        public void TopLeftAnchorAlignsTheTopLeftCorners()
        {
            AnchorSwapPlacementPolicy.SwapPlacement(AnchorSwapPlacementPolicy.SwapAnchor.TopLeft,
                L1, T1, W1, H1, L2, T2, W2, H2,
                out float newLeft1, out float newTop1, out _, out _);
            // Shape 1's top-left lands exactly on shape 2's top-left.
            Assert.AreEqual(400f, newLeft1, 0.001);
            Assert.AreEqual(300f, newTop1, 0.001);
        }

        [TestMethod]
        public void EveryAnchorPointSitsOnItsRectangle()
        {
            var anchors = new[]
            {
                AnchorSwapPlacementPolicy.SwapAnchor.TopLeft, AnchorSwapPlacementPolicy.SwapAnchor.Top,
                AnchorSwapPlacementPolicy.SwapAnchor.TopRight, AnchorSwapPlacementPolicy.SwapAnchor.Right,
                AnchorSwapPlacementPolicy.SwapAnchor.BottomRight, AnchorSwapPlacementPolicy.SwapAnchor.Bottom,
                AnchorSwapPlacementPolicy.SwapAnchor.BottomLeft, AnchorSwapPlacementPolicy.SwapAnchor.Left,
                AnchorSwapPlacementPolicy.SwapAnchor.Center
            };
            float x, y;
            foreach (var anchor in anchors)
            {
                AnchorSwapPlacementPolicy.GetAnchorPoint(anchor, L1, T1, W1, H1, out x, out y);
                Assert.IsTrue(x >= L1 - 0.001 && x <= L1 + W1 + 0.001, $"{anchor}: x {x} outside rectangle");
                Assert.IsTrue(y >= T1 - 0.001 && y <= T1 + H1 + 0.001, $"{anchor}: y {y} outside rectangle");
            }

            // Spot-check three corners.
            AnchorSwapPlacementPolicy.GetAnchorPoint(AnchorSwapPlacementPolicy.SwapAnchor.TopRight, L1, T1, W1, H1, out x, out y);
            Assert.AreEqual(160f, x, 0.001);
            Assert.AreEqual(50f, y, 0.001);
            AnchorSwapPlacementPolicy.GetAnchorPoint(AnchorSwapPlacementPolicy.SwapAnchor.Bottom, L1, T1, W1, H1, out x, out y);
            Assert.AreEqual(130f, x, 0.001);
            Assert.AreEqual(90f, y, 0.001);
            AnchorSwapPlacementPolicy.GetAnchorPoint(AnchorSwapPlacementPolicy.SwapAnchor.Left, L1, T1, W1, H1, out x, out y);
            Assert.AreEqual(100f, x, 0.001);
            Assert.AreEqual(70f, y, 0.001);
        }
    }
}
