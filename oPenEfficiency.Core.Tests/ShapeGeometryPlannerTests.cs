using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class ShapeGeometryPlannerTests
    {
        [TestMethod]
        public void SpacingAdjustmentUsesSpatialOrderAndCanPreserveCenter()
        {
            var source = Shapes(
                Shape(1, 100, 0, 10, 10),
                Shape(2, 0, 0, 10, 10),
                Shape(3, 50, 0, 10, 10));

            var normal = ShapeGeometryPlanner.AdjustSpacing(source, 10, true, false);
            var centered = ShapeGeometryPlanner.AdjustSpacing(source, 10, true, true);

            Assert.AreEqual(0, ById(normal, 2).Left);
            Assert.AreEqual(60, ById(normal, 3).Left);
            Assert.AreEqual(120, ById(normal, 1).Left);
            Assert.AreEqual(-10, ById(centered, 2).Left);
            Assert.AreEqual(50, ById(centered, 3).Left);
            Assert.AreEqual(110, ById(centered, 1).Left);
            Assert.AreEqual(0, source[1].Left, "Planning mutated the source geometry.");
        }

        [TestMethod]
        public void RemoveGapsKeepsAnchorFixedAndPacksBothDirections()
        {
            var result = ShapeGeometryPlanner.RemoveGaps(
                Shapes(
                    Shape(1, 0, 5, 10, 5),
                    Shape(2, 40, 20, 20, 10),
                    Shape(3, 100, 50, 30, 20)),
                true,
                false,
                2);

            Assert.AreEqual(30, ById(result, 1).Left);
            Assert.AreEqual(40, ById(result, 2).Left);
            Assert.AreEqual(60, ById(result, 3).Left);
            Assert.AreEqual(5, ById(result, 1).Top);
        }

        [TestMethod]
        public void CenterPackingPreservesSelectionBoundingBoxCenter()
        {
            var source = Shapes(
                Shape(1, 10, 0, 20, 10),
                Shape(2, 100, 0, 30, 10),
                Shape(3, 200, 0, 10, 10));
            float originalCenter = (source.Min(shape => shape.Left) + source.Max(shape => shape.Left + shape.Width)) / 2;

            var result = ShapeGeometryPlanner.RemoveGaps(source, true, true, null);
            float resultCenter = (result.Min(shape => shape.Left) + result.Max(shape => shape.Left + shape.Width)) / 2;

            Assert.AreEqual(originalCenter, resultCenter, 1e-5);
            var ordered = result.OrderBy(shape => shape.Left).ToList();
            Assert.AreEqual(ordered[0].Left + ordered[0].Width, ordered[1].Left, 1e-5);
            Assert.AreEqual(ordered[1].Left + ordered[1].Width, ordered[2].Left, 1e-5);
        }

        [TestMethod]
        public void GridUsesLargestCellDimensionsAndSelectionOrder()
        {
            var result = ShapeGeometryPlanner.ArrangeGrid(
                Shapes(
                    Shape(1, 20, 30, 10, 5),
                    Shape(2, 50, 60, 30, 15),
                    Shape(3, 90, 90, 20, 25)),
                2,
                2,
                4);

            AssertGeometry(result[0], 20, 30, 10, 5);
            AssertGeometry(result[1], 44, 30, 30, 15);
            AssertGeometry(result[2], 20, 49, 20, 25);
        }

        [TestMethod]
        public void SplitGridFillsBoundsWithExactSpacing()
        {
            var result = ShapeGeometryPlanner.SplitGrid(Shape(1, 10, 20, 100, 50), 2, 3, 5);

            Assert.AreEqual(6, result.Count);
            Assert.AreEqual(30, result[0].Width, 1e-5);
            Assert.AreEqual(22.5, result[0].Height, 1e-5);
            Assert.AreEqual(80, result[2].Left, 1e-5);
            Assert.AreEqual(47.5, result[3].Top, 1e-5);
            Assert.AreEqual(110, result.Max(shape => shape.Left + shape.Width), 1e-5);
            Assert.AreEqual(70, result.Max(shape => shape.Top + shape.Height), 1e-5);
        }

        [TestMethod]
        public void GeometryPlannerRejectsInvalidInputs()
        {
            Assert.ThrowsException<ArgumentException>(() =>
                ShapeGeometryPlanner.ArrangeGrid(Shapes(Shape(1, 0, 0, 10, 10), Shape(2, 0, 0, 10, 10)), 1, 1, 0));
            Assert.ThrowsException<ArgumentException>(() =>
                ShapeGeometryPlanner.RemoveGaps(Shapes(Shape(1, 0, 0, 10, 10)), true, false, 99));
            Assert.ThrowsException<ArgumentException>(() =>
                ShapeGeometryPlanner.SplitGrid(Shape(1, 0, 0, 10, 10), 1, 3, 6));
            Assert.ThrowsException<ArgumentException>(() =>
                ShapeGeometryPlanner.AdjustSpacing(Shapes(Shape(1, 0, 0, 0, 10)), 1, true, false));
        }

        private static List<ShapeGeometry> Shapes(params ShapeGeometry[] shapes) => shapes.ToList();

        private static ShapeGeometry Shape(int id, float left, float top, float width, float height)
        {
            return new ShapeGeometry { Id = id, Left = left, Top = top, Width = width, Height = height };
        }

        private static ShapeGeometry ById(IEnumerable<ShapeGeometry> shapes, int id)
            => shapes.Single(shape => shape.Id == id);

        private static void AssertGeometry(ShapeGeometry shape, float left, float top, float width, float height)
        {
            Assert.AreEqual(left, shape.Left, 1e-5);
            Assert.AreEqual(top, shape.Top, 1e-5);
            Assert.AreEqual(width, shape.Width, 1e-5);
            Assert.AreEqual(height, shape.Height, 1e-5);
        }
    }
}
