using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;
using static oPenEfficiency.Core.Services.DiagonalPlacementPolicy;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 align-and-size-wave coverage for Align Diagonally (QA9-BtnDiagonalAlign-002):
    /// the projection, ordering, spacing and anchor arithmetic used by DiagonalAlignFeature.
    /// </summary>
    [TestClass]
    public sealed class DiagonalPlacementPolicyTests
    {
        // Three shapes, horizontal fixture: centers (100,50) 20x10, (200,80) 40x20, (300,10) 20x10.
        private static readonly float[] Cx = { 100f, 200f, 300f };
        private static readonly float[] Cy = { 50f, 80f, 10f };
        private static readonly float[] W = { 20f, 40f, 20f };
        private static readonly float[] H = { 10f, 20f, 10f };

        [TestMethod]
        public void FirstAnchorPlacesTheLineOnTheFirstProjectedShape()
        {
            ComputePlacements(Cx, Cy, W, H, 0f, DiagonalAnchor.First, out float[] left, out float[] top);

            CollectionAssert.AreEqual(new[] { 90f, 180f, 290f }, left);
            CollectionAssert.AreEqual(new[] { 45f, 40f, 45f }, top);
        }

        [TestMethod]
        public void CenterAnchorKeepsTheSelectionCenterInTheMiddle()
        {
            ComputePlacements(Cx, Cy, W, H, 0f, DiagonalAnchor.Center, out float[] left, out float[] top);

            CollectionAssert.AreEqual(new[] { 90f, 180f, 290f }, left);
            Assert.AreEqual(41.6667f, top[0], 0.01);
            Assert.AreEqual(36.6667f, top[1], 0.01);
            Assert.AreEqual(41.6667f, top[2], 0.01);
        }

        [TestMethod]
        public void LastAnchorPlacesTheLineOnTheLastProjectedShape()
        {
            ComputePlacements(Cx, Cy, W, H, 0f, DiagonalAnchor.Last, out float[] left, out float[] top);

            CollectionAssert.AreEqual(new[] { 90f, 180f, 290f }, left);
            CollectionAssert.AreEqual(new[] { 5f, 0f, 5f }, top);
        }

        [TestMethod]
        public void AStaggeredSelectionIsOrderedByProjectionNotBySelectionOrder()
        {
            // Same three shapes, but selected in the order C, A, B (sizes follow their shapes).
            ComputePlacements(
                new[] { 300f, 100f, 200f }, new[] { 10f, 50f, 80f },
                new[] { 20f, 20f, 40f }, new[] { 10f, 10f, 20f },
                0f, DiagonalAnchor.First, out float[] left, out float[] top);

            CollectionAssert.AreEqual(new[] { 290f, 90f, 180f }, left);   // C, A, B
            CollectionAssert.AreEqual(new[] { 45f, 45f, 40f }, top);
        }

        [TestMethod]
        public void StackedShapesFallBackToTheAverageExtentAlongTheDirection()
        {
            // All three centers share the same projection, so the span-derived spacing collapses.
            ComputePlacements(
                new[] { 150f, 150f, 150f }, new[] { 50f, 80f, 10f }, W, H, 0f, DiagonalAnchor.First,
                out float[] left, out float[] top);

            Assert.AreEqual(140f, left[0], 0.01);
            Assert.AreEqual(156.6667f, left[1], 0.01);
            Assert.AreEqual(193.3333f, left[2], 0.01);
            CollectionAssert.AreEqual(new[] { 45f, 40f, 45f }, top);
        }

        [TestMethod]
        public void FortyFiveDegreesWalksTheLineDownRight()
        {
            ComputePlacements(
                new[] { 0f, 100f }, new[] { 0f, 0f }, new[] { 10f, 10f }, new[] { 10f, 10f },
                45f, DiagonalAnchor.First, out float[] left, out float[] top);

            Assert.AreEqual(-5f, left[0], 0.01);
            Assert.AreEqual(-5f, top[0], 0.01);
            Assert.AreEqual(45f, left[1], 0.01);
            Assert.AreEqual(45f, top[1], 0.01);
        }

        [TestMethod]
        public void TheToolbarPathPlacesTheLineWithTheUsersSpacing()
        {
            ComputePlacementsWithSpacing(Cx, Cy, W, H, 0f, 50f, DiagonalAnchor.First,
                out float[] left, out float[] top);

            CollectionAssert.AreEqual(new[] { 90f, 130f, 190f }, left);
            CollectionAssert.AreEqual(new[] { 45f, 40f, 45f }, top);
        }

        [TestMethod]
        public void TheDirectionVectorFollowsTheAngle()
        {
            GetDirection(0f, out float x0, out float y0);
            Assert.AreEqual(1f, x0, 1e-6);
            Assert.AreEqual(0f, y0, 1e-6);

            GetDirection(90f, out float x90, out float y90);
            Assert.AreEqual(0f, x90, 1e-6);
            Assert.AreEqual(1f, y90, 1e-6);

            GetDirection(45f, out float x45, out float y45);
            Assert.AreEqual(0.7071068f, x45, 1e-6);
            Assert.AreEqual(0.7071068f, y45, 1e-6);
        }

        [TestMethod]
        public void TheAnchorStringMapsExactlyLikeTheHistoricalComparison()
        {
            Assert.AreEqual(DiagonalAnchor.First, ResolveAnchor("First"));
            Assert.AreEqual(DiagonalAnchor.Last, ResolveAnchor("Last"));
            Assert.AreEqual(DiagonalAnchor.Center, ResolveAnchor("Center"));
            Assert.AreEqual(DiagonalAnchor.Center, ResolveAnchor("first"));   // case-sensitive, as before
            Assert.AreEqual(DiagonalAnchor.Center, ResolveAnchor(null));
        }
    }
}
