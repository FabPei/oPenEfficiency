using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 arrange-wave coverage for Snap Shape to Object (QA9-BtnSnapShape-002): the
    /// nearest-side direction decision and the flush placement centered on the perpendicular
    /// axis, used by SnapShapeFeature.
    /// </summary>
    [TestClass]
    public sealed class ShapeSnapPlacementPolicyTests
    {
        private const float TL = 200f, TT = 100f, TW = 100f, TH = 50f; // target rectangle

        [TestMethod]
        public void NearestDirectionFollowsTheLargerCenterDelta()
        {
            // Moving left of the target: horizontal delta dominates -> Left.
            Assert.AreEqual(ShapeSnapPlacementPolicy.SnapDirection.Left,
                ShapeSnapPlacementPolicy.GetNearestDirection(0f, 110f, 50f, 30f, TL, TT, TW, TH));
            // Moving right: -> Right.
            Assert.AreEqual(ShapeSnapPlacementPolicy.SnapDirection.Right,
                ShapeSnapPlacementPolicy.GetNearestDirection(450f, 110f, 50f, 30f, TL, TT, TW, TH));
            // Above: vertical delta dominates -> Above.
            Assert.AreEqual(ShapeSnapPlacementPolicy.SnapDirection.Above,
                ShapeSnapPlacementPolicy.GetNearestDirection(215f, 0f, 50f, 30f, TL, TT, TW, TH));
            // Below: -> Below.
            Assert.AreEqual(ShapeSnapPlacementPolicy.SnapDirection.Below,
                ShapeSnapPlacementPolicy.GetNearestDirection(215f, 200f, 50f, 30f, TL, TT, TW, TH));
        }

        [TestMethod]
        public void HorizontalWinsTies()
        {
            // Equal center deltas on both axes resolve horizontally (>= comparison).
            Assert.AreEqual(ShapeSnapPlacementPolicy.SnapDirection.Left,
                ShapeSnapPlacementPolicy.GetNearestDirection(0f, 0f, 50f, 30f, TL, TT, TW, TH));
            // Identical centers: both deltas zero, tie -> Right (non-negative delta).
            Assert.AreEqual(ShapeSnapPlacementPolicy.SnapDirection.Right,
                ShapeSnapPlacementPolicy.GetNearestDirection(225f, 110f, 50f, 30f, TL, TT, TW, TH));
        }

        [TestMethod]
        public void PlacementSitsFlushAndCentersThePerpendicularAxis()
        {
            ShapeSnapPlacementPolicy.Place(ShapeSnapPlacementPolicy.SnapDirection.Left,
                TL, TT, TW, TH, 60f, 30f, out float leftL, out float topL);
            Assert.AreEqual(140f, leftL, 0.001); // target left minus moving width
            Assert.AreEqual(110f, topL, 0.001);  // centered on the target's height

            ShapeSnapPlacementPolicy.Place(ShapeSnapPlacementPolicy.SnapDirection.Right,
                TL, TT, TW, TH, 60f, 30f, out float leftR, out float topR);
            Assert.AreEqual(300f, leftR, 0.001); // target right edge
            Assert.AreEqual(110f, topR, 0.001);

            ShapeSnapPlacementPolicy.Place(ShapeSnapPlacementPolicy.SnapDirection.Above,
                TL, TT, TW, TH, 60f, 30f, out float leftA, out float topA);
            Assert.AreEqual(220f, leftA, 0.001); // centered on the target's width
            Assert.AreEqual(70f, topA, 0.001);   // target top minus moving height

            ShapeSnapPlacementPolicy.Place(ShapeSnapPlacementPolicy.SnapDirection.Below,
                TL, TT, TW, TH, 60f, 30f, out float leftB, out float topB);
            Assert.AreEqual(220f, leftB, 0.001);
            Assert.AreEqual(150f, topB, 0.001); // target bottom edge
        }

        [TestMethod]
        public void OversizedMovingShapeGetsNegativePerpendicularOffset()
        {
            // A moving shape taller than the target centers with a negative offset.
            ShapeSnapPlacementPolicy.Place(ShapeSnapPlacementPolicy.SnapDirection.Right,
                TL, TT, TW, TH, 40f, 90f, out float left, out float top);
            Assert.AreEqual(300f, left, 0.001);
            Assert.AreEqual(80f, top, 0.001); // 100 + (50 - 90) / 2
        }
    }
}
