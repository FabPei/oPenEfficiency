using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 1 acceptance coverage for the grid geometry features: BtnArrangeGrid
    /// (<c>ShapeGeometryPlanner.ArrangeGrid</c>, driven by ArrangeGridFeature after the layout dialog) and
    /// BtnSplitShape (<c>ShapeGeometryPlanner.SplitGrid</c>, the feature's documented 2x2 / no-spacing
    /// default). Both plans are the production decision the features apply to the COM shapes; the COM leg
    /// (writing coordinates, grouping the split result, the Undo entry) stays office-automated.
    /// </summary>
    [TestClass]
    public sealed class GridArrangeAcceptanceTests
    {
        private const float Precision = 1e-5f;

        [TestMethod]
        public void ArrangeGridPlanIsIdempotentAndAnchoredAtTheSelectionBoundingBox()
        {
            var source = Shapes(
                Shape(1, 20, 30, 10, 5),
                Shape(2, 50, 60, 30, 15),
                Shape(3, 90, 90, 20, 25));

            var plan = ShapeGeometryPlanner.ArrangeGrid(source, rows: 2, columns: 2, spacing: 4f);

            // Cells are filled in selection order and sized by the widest/tallest shape of their row/column.
            AssertGeometry(plan[0], 20, 30, 10, 5);
            AssertGeometry(plan[1], 44, 30, 30, 15);
            AssertGeometry(plan[2], 20, 49, 20, 25);

            // The selection's bounding box corner stays where it was and the requested spacing is exact: the
            // gap is measured from the widest shape of column 0 (plan[2]) and the tallest shape of row 0 (plan[1]).
            Assert.AreEqual(source.Min(shape => shape.Left), plan.Min(shape => shape.Left), Precision);
            Assert.AreEqual(source.Min(shape => shape.Top), plan.Min(shape => shape.Top), Precision);
            Assert.AreEqual(4f, plan[1].Left - plan[0].Left - plan[2].Width, Precision);
            Assert.AreEqual(4f, plan[2].Top - plan[0].Top - plan[1].Height, Precision);

            // Planning must not mutate its input, and planning the plan again must change nothing.
            Assert.AreEqual(50f, source[1].Left, Precision, "Planning mutated the source geometry.");
            var second = ShapeGeometryPlanner.ArrangeGrid(plan, rows: 2, columns: 2, spacing: 4f);
            for (int index = 0; index < plan.Count; index++)
            {
                AssertGeometry(second[index], plan[index].Left, plan[index].Top, plan[index].Width, plan[index].Height);
            }

            // Invalid scope: a grid that cannot hold the selection is refused instead of leaving shapes stacked.
            Assert.ThrowsException<ArgumentException>(() =>
                ShapeGeometryPlanner.ArrangeGrid(source, rows: 1, columns: 2, spacing: 4f));
        }

        [TestMethod]
        public void SplitShapePlanTilesTheBoundsWithoutOverlap()
        {
            var bounds = Shape(0, 10, 20, 100, 50);

            var cells = ShapeGeometryPlanner.SplitGrid(bounds, rows: 2, columns: 3, spacing: 5f);

            Assert.AreEqual(6, cells.Count, "A 2x3 split must produce one shape per cell.");
            Assert.AreEqual(30f, cells[0].Width, Precision);
            Assert.AreEqual(22.5f, cells[0].Height, Precision);

            // The tiling covers the bounds exactly: the last column/row ends on the original edge.
            Assert.AreEqual(110f, cells.Max(cell => cell.Left + cell.Width), Precision);
            Assert.AreEqual(70f, cells.Max(cell => cell.Top + cell.Height), Precision);
            Assert.AreEqual(10f, cells.Min(cell => cell.Left), Precision);
            Assert.AreEqual(20f, cells.Min(cell => cell.Top), Precision);

            // No overlap and exactly the requested spacing between neighbours.
            foreach (var row in cells.GroupBy(cell => cell.Top))
            {
                var ordered = row.OrderBy(cell => cell.Left).ToList();
                for (int index = 1; index < ordered.Count; index++)
                {
                    Assert.AreEqual(5f, ordered[index].Left - ordered[index - 1].Left - ordered[index - 1].Width, Precision);
                }
            }
            foreach (var column in cells.GroupBy(cell => cell.Left))
            {
                var ordered = column.OrderBy(cell => cell.Top).ToList();
                for (int index = 1; index < ordered.Count; index++)
                {
                    Assert.AreEqual(5f, ordered[index].Top - ordered[index - 1].Top - ordered[index - 1].Height, Precision);
                }
            }

            // The documented default of BtnSplitShape: 2x2 with no spacing quarter-tiles the selected shape.
            var defaultCells = ShapeGeometryPlanner.SplitGrid(bounds, rows: 2, columns: 2, spacing: 0f);
            Assert.AreEqual(4, defaultCells.Count);
            Assert.AreEqual(50f, defaultCells[0].Width, Precision);
            Assert.AreEqual(25f, defaultCells[0].Height, Precision);
            Assert.AreEqual(110f, defaultCells.Max(cell => cell.Left + cell.Width), Precision);

            // Invalid scope: spacing that leaves no positive cell area is refused.
            Assert.ThrowsException<ArgumentException>(() =>
                ShapeGeometryPlanner.SplitGrid(bounds, rows: 1, columns: 3, spacing: 50f));
        }

        private static List<ShapeGeometry> Shapes(params ShapeGeometry[] shapes) => shapes.ToList();

        private static ShapeGeometry Shape(int id, float left, float top, float width, float height)
            => new ShapeGeometry { Id = id, Left = left, Top = top, Width = width, Height = height };

        private static void AssertGeometry(ShapeGeometry shape, float left, float top, float width, float height)
        {
            Assert.AreEqual(left, shape.Left, Precision);
            Assert.AreEqual(top, shape.Top, Precision);
            Assert.AreEqual(width, shape.Width, Precision);
            Assert.AreEqual(height, shape.Height, Precision);
        }
    }
}
