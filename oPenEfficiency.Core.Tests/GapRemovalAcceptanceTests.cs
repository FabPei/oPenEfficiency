using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 1 acceptance coverage for the gap-removal features (BtnRemoveHorizontalGaps,
    /// BtnRemoveVerticalGaps and their "centered" context commands), which decide the packed coordinates with
    /// <c>ShapeGeometryPlanner.RemoveGaps</c> - the host-neutral production code driven here. The COM leg
    /// (applying the coordinates, grouping the Undo entry) stays office-automated.
    /// </summary>
    [TestClass]
    public sealed class GapRemovalAcceptanceTests
    {
        private const float Precision = 1e-5f;

        [TestMethod]
        public void RemoveHorizontalGapsPacksFromTheAnchorAndLeavesTheOtherAxisAlone()
        {
            var source = Shapes(
                Shape(1, 0, 5, 10, 5),
                Shape(2, 40, 20, 20, 10),
                Shape(3, 100, 50, 30, 20));

            var packed = ShapeGeometryPlanner.RemoveGaps(source, horizontal: true, packToCenter: false, anchorId: 2);

            // The anchor keeps its own coordinate; everything left of it is packed leftwards and everything
            // right of it rightwards, so the shapes touch without a gap.
            Assert.AreEqual(40f, ById(packed, 2).Left, Precision, "The anchor moved.");
            Assert.AreEqual(30f, ById(packed, 1).Left, Precision);
            Assert.AreEqual(60f, ById(packed, 3).Left, Precision);
            AssertNoGap(packed, horizontal: true);

            // Packing one axis must not move the other one, and the plan must not mutate its input.
            Assert.AreEqual(5f, ById(packed, 1).Top, Precision);
            Assert.AreEqual(20f, ById(packed, 2).Top, Precision);
            Assert.AreEqual(50f, ById(packed, 3).Top, Precision);
            Assert.AreEqual(0f, source[0].Left, Precision, "Planning mutated the source geometry.");

            // Removing the gaps again is a no-op.
            var again = ShapeGeometryPlanner.RemoveGaps(packed, horizontal: true, packToCenter: false, anchorId: 2);
            foreach (var shape in packed)
            {
                Assert.AreEqual(shape.Left, ById(again, shape.Id).Left, Precision, "Gap removal is not idempotent.");
            }

            // Invalid scope: an anchor that is not part of the selection is refused.
            Assert.ThrowsException<ArgumentException>(() =>
                ShapeGeometryPlanner.RemoveGaps(source, horizontal: true, packToCenter: false, anchorId: 99));
        }

        [TestMethod]
        public void RemoveVerticalGapsPacksVerticallyAndPreservesTheBoundingBoxCenterWhenCentered()
        {
            var source = Shapes(
                Shape(1, 0, 0, 10, 10),
                Shape(2, 50, 30, 20, 20),
                Shape(3, 100, 90, 30, 30));

            var packed = ShapeGeometryPlanner.RemoveGaps(source, horizontal: false, packToCenter: false, anchorId: 1);

            Assert.AreEqual(0f, ById(packed, 1).Top, Precision, "The anchor moved.");
            Assert.AreEqual(10f, ById(packed, 2).Top, Precision);
            Assert.AreEqual(30f, ById(packed, 3).Top, Precision);
            AssertNoGap(packed, horizontal: false);
            Assert.AreEqual(0f, ById(packed, 1).Left, Precision);
            Assert.AreEqual(50f, ById(packed, 2).Left, Precision);
            Assert.AreEqual(100f, ById(packed, 3).Left, Precision);

            // The centered context command packs around the middle of the original bounding box.
            float originalCenter = (source.Min(shape => shape.Top) +
                source.Max(shape => shape.Top + shape.Height)) / 2f;
            var centered = ShapeGeometryPlanner.RemoveGaps(source, horizontal: false, packToCenter: true, anchorId: null);
            float centeredCenter = (centered.Min(shape => shape.Top) +
                centered.Max(shape => shape.Top + shape.Height)) / 2f;
            Assert.AreEqual(originalCenter, centeredCenter, Precision, "Centered packing moved the selection centre.");
            AssertNoGap(centered, horizontal: false);

            var again = ShapeGeometryPlanner.RemoveGaps(centered, horizontal: false, packToCenter: true, anchorId: null);
            foreach (var shape in centered)
            {
                Assert.AreEqual(shape.Top, ById(again, shape.Id).Top, Precision, "Centered gap removal is not idempotent.");
            }
        }

        private static void AssertNoGap(IEnumerable<ShapeGeometry> shapes, bool horizontal)
        {
            var ordered = shapes
                .OrderBy(shape => horizontal ? shape.Left : shape.Top)
                .ThenBy(shape => shape.Id)
                .ToList();
            for (int index = 1; index < ordered.Count; index++)
            {
                float edge = horizontal ? ordered[index - 1].Left + ordered[index - 1].Width : ordered[index - 1].Top + ordered[index - 1].Height;
                float next = horizontal ? ordered[index].Left : ordered[index].Top;
                Assert.AreEqual(edge, next, Precision, "A gap is left between two shapes.");
            }
        }

        private static List<ShapeGeometry> Shapes(params ShapeGeometry[] shapes) => shapes.ToList();

        private static ShapeGeometry Shape(int id, float left, float top, float width, float height)
            => new ShapeGeometry { Id = id, Left = left, Top = top, Width = width, Height = height };

        private static ShapeGeometry ById(IEnumerable<ShapeGeometry> shapes, int id)
            => shapes.Single(shape => shape.Id == id);
    }
}
