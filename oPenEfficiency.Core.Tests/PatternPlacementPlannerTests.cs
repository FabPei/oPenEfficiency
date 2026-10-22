using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// Pins the host-neutral pattern-placement math of the geometric arrangement commands
    /// (BtnArrangeInShape / BtnArrangeCircle / ArrangePro presets). These are the named behaviors
    /// for QA9-BtnArrangeInShape-001 and QA9-BtnArrangeCircle-001.
    /// </summary>
    [TestClass]
    public sealed class PatternPlacementPlannerTests
    {
        private const float Tolerance = 0.001f;

        private static void AssertPoint(float expectedX, float expectedY, Tuple<float, float> actual)
        {
            Assert.IsTrue(Math.Abs(expectedX - actual.Item1) <= Tolerance, $"x: expected {expectedX}, actual {actual.Item1}");
            Assert.IsTrue(Math.Abs(expectedY - actual.Item2) <= Tolerance, $"y: expected {expectedY}, actual {actual.Item2}");
        }

        [TestMethod]
        public void CirclePatternPlacesCountPointsOnTheRadiusClockwiseFromTop()
        {
            var targets = PatternPlacementPlanner.BuildTargets("circle", 4, 160f, 200f, 147.5f, 0f, false, true, 360f, false);

            Assert.AreEqual(4, targets.Count);
            AssertPoint(200f, -12.5f, targets[0]);
            AssertPoint(360f, 147.5f, targets[1]);
            AssertPoint(200f, 307.5f, targets[2]);
            AssertPoint(40f, 147.5f, targets[3]);
        }

        [TestMethod]
        public void TriangleOutlinePlacesItemsOnTheTrianglePerimeter()
        {
            var targets = PatternPlacementPlanner.BuildTargets("triangle", 4, 160f, 200f, 147.5f, 0f, false, true, 270f, false);

            AssertPoint(200f, -12.5f, targets[0]);
            AssertPoint(303.9231f, 167.5f, targets[1]);
            AssertPoint(200f, 227.5f, targets[2]);
            AssertPoint(96.0769f, 167.5f, targets[3]);
        }

        [TestMethod]
        public void CounterclockwiseTraversesThePathBackwards()
        {
            var targets = PatternPlacementPlanner.BuildTargets("circle", 4, 160f, 200f, 147.5f, 0f, false, false, 360f, false);

            AssertPoint(200f, -12.5f, targets[0]);
            AssertPoint(40f, 147.5f, targets[1]);
            AssertPoint(200f, 307.5f, targets[2]);
            AssertPoint(360f, 147.5f, targets[3]);
        }

        [TestMethod]
        public void StartAngleRotatesAllTargetsAroundTheCenter()
        {
            var rotated = PatternPlacementPlanner.BuildTargets("circle", 4, 160f, 200f, 147.5f, 90f, false, true, 360f, false);

            AssertPoint(360f, 147.5f, rotated[0]);
            AssertPoint(200f, 307.5f, rotated[1]);
            AssertPoint(40f, 147.5f, rotated[2]);
            AssertPoint(200f, -12.5f, rotated[3]);
        }

        [TestMethod]
        public void ArcDistributesAcrossTheSpanIncludingTheEndpoints()
        {
            var targets = PatternPlacementPlanner.BuildTargets("arc", 3, 100f, 50f, 50f, 0f, false, true, 180f, false);

            AssertPoint(50f, -50f, targets[0]);
            AssertPoint(150f, 50f, targets[1]);
            AssertPoint(50f, 150f, targets[2]);
        }

        [TestMethod]
        public void FullCircleArcSplitsTheCircleWithoutDuplicatedEndpoints()
        {
            var targets = PatternPlacementPlanner.BuildTargets("arc", 4, 100f, 0f, 0f, 0f, false, true, 360f, false);

            AssertPoint(0f, -100f, targets[0]);
            AssertPoint(100f, 0f, targets[1]);
            AssertPoint(0f, 100f, targets[2]);
            AssertPoint(-100f, 0f, targets[3]);
        }

        [TestMethod]
        public void FilledSquarePlacesItemsOnAGrid()
        {
            var targets = PatternPlacementPlanner.BuildTargets("square", 4, 100f, 0f, 0f, 0f, true, true, 360f, false);

            AssertPoint(-100f, -100f, targets[0]);
            AssertPoint(100f, -100f, targets[1]);
            AssertPoint(-100f, 100f, targets[2]);
            AssertPoint(100f, 100f, targets[3]);
        }

        [TestMethod]
        public void FilledTriangleBuildsCompleteRowsBottomHeavy()
        {
            var targets = PatternPlacementPlanner.BuildTargets("triangle", 6, 100f, 0f, 0f, 0f, true, true, 360f, false);

            Assert.AreEqual(6, targets.Count);
            var rowTops = targets.Select(t => Math.Round(t.Item2, 3)).Distinct().OrderBy(v => v).ToList();
            Assert.AreEqual(3, rowTops.Count);
        }

        [TestMethod]
        public void StarVerticesAlternateBetweenOuterAndInnerRadius()
        {
            float x, y;
            PatternPlacementPlanner.GetOutlinePoint("star", 0f, 0f, 100f, 0, 1, 10, out x, out y);

            double distanceFromCenter = Math.Sqrt(x * x + y * y);
            Assert.IsTrue(Math.Abs(distanceFromCenter - 45f) <= 0.5f,
                $"inner star vertices use 0.45r; expected 45, actual {distanceFromCenter}");
        }

        [TestMethod]
        public void ReservedCenterMovesOnlyTheCenteredTarget()
        {
            var targets = PatternPlacementPlanner.BuildTargets("circle", 1, 100f, 0f, 0f, 0f, true, true, 360f, true);

            double distanceFromCenter = Math.Sqrt(targets[0].Item1 * targets[0].Item1 + targets[0].Item2 * targets[0].Item2);
            Assert.IsTrue(distanceFromCenter > 1f, "the reserved center must not stay on the center");
            Assert.IsTrue(distanceFromCenter <= 100f, "the displaced target must stay within the radius");
        }
    }
}
