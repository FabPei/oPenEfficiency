using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave-5 coverage for Repeat Shape geometry (QA9-BtnRepeatShape-002): the inferred
    /// gap between the two selected rectangles and the duplicate placement in each repeat
    /// direction, used by RepeatShapeFeature.
    /// </summary>
    [TestClass]
    public sealed class RepeatShapeGeometryPlannerTests
    {
        [TestMethod]
        public void GapInferenceUsesEdgeToEdgeDistanceWhenDisjoint()
        {
            Assert.AreEqual(30f, RepeatShapeGeometryPlanner.InferGapX(0f, 100f, 130f, 50f), 0.001);
            Assert.AreEqual(20f, RepeatShapeGeometryPlanner.InferGapY(10f, 40f, 70f, 40f), 0.001);
        }

        [TestMethod]
        public void GapInferenceFallsBackToTheAbsoluteOffsetWhenOverlapping()
        {
            Assert.AreEqual(15f, RepeatShapeGeometryPlanner.InferGapX(0f, 100f, 15f, 50f), 0.001);
            Assert.AreEqual(5f, RepeatShapeGeometryPlanner.InferGapY(10f, 40f, 15f, 40f), 0.001);
        }

        [TestMethod]
        public void HorizontalPlacementAnchorsOnTheSecondSelectedShape()
        {
            RepeatShapeGeometryPlanner.NextPlacement(
                RepeatShapeGeometryPlanner.RepeatPlacementDirection.Right,
                130f, 70f, 50f, 40f, 50f, 40f, 30f, 20f,
                out float rightLeft, out float rightTop);
            Assert.AreEqual(210f, rightLeft, 0.001);
            Assert.AreEqual(70f, rightTop, 0.001);

            RepeatShapeGeometryPlanner.NextPlacement(
                RepeatShapeGeometryPlanner.RepeatPlacementDirection.Left,
                130f, 70f, 50f, 40f, 60f, 40f, 30f, 20f,
                out float leftLeft, out float leftTop);
            Assert.AreEqual(40f, leftLeft, 0.001); // uses the duplicate's own width
            Assert.AreEqual(70f, leftTop, 0.001);
        }

        [TestMethod]
        public void VerticalPlacementAnchorsOnTheSecondSelectedShape()
        {
            RepeatShapeGeometryPlanner.NextPlacement(
                RepeatShapeGeometryPlanner.RepeatPlacementDirection.Down,
                130f, 70f, 50f, 40f, 50f, 40f, 30f, 20f,
                out float downLeft, out float downTop);
            Assert.AreEqual(130f, downLeft, 0.001);
            Assert.AreEqual(130f, downTop, 0.001);

            RepeatShapeGeometryPlanner.NextPlacement(
                RepeatShapeGeometryPlanner.RepeatPlacementDirection.Up,
                130f, 70f, 50f, 40f, 50f, 55f, 30f, 20f,
                out float upLeft, out float upTop);
            Assert.AreEqual(130f, upLeft, 0.001);
            Assert.AreEqual(-5f, upTop, 0.001); // uses the duplicate's own height
        }
    }
}
