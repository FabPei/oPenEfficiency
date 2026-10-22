using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.UI;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class GridPreviewPlannerTests
    {
        private const string MultiplySign = "\u00D7";
        private const string InvalidLabel = "\u2013 \u00D7 \u2013";

        [TestMethod]
        public void StepButtonColumnChangeUpdatesPreviewState()
        {
            var before = GridPreviewPlanner.Create("3", "3");
            var after = GridPreviewPlanner.Create("4", "3");

            Assert.AreEqual("3 " + MultiplySign + " 3", before.DimensionLabel);
            Assert.AreEqual("4 " + MultiplySign + " 3", after.DimensionLabel);
            Assert.AreEqual(3, before.VisibleColumns);
            Assert.AreEqual(4, after.VisibleColumns);
            Assert.AreEqual(9, before.CellCount);
            Assert.AreEqual(12, after.CellCount);
            Assert.AreEqual("Grid preview, 4 columns by 3 rows", after.AutomationName);
        }

        [TestMethod]
        public void StepButtonRowChangeUpdatesPreviewState()
        {
            var before = GridPreviewPlanner.Create("3", "3");
            var after = GridPreviewPlanner.Create("3", "2");

            Assert.AreEqual("3 " + MultiplySign + " 2", after.DimensionLabel);
            Assert.AreEqual(2, after.VisibleRows);
            Assert.AreEqual(6, after.CellCount);
            Assert.AreEqual("Grid preview, 3 columns by 2 rows", after.AutomationName);
        }

        [TestMethod]
        public void DirectTextChangeUpdatesPreviewState()
        {
            var before = GridPreviewPlanner.Create("3", "3");
            var after = GridPreviewPlanner.Create("12", "3");

            Assert.IsTrue(after.IsValid);
            Assert.AreEqual("12 " + MultiplySign + " 3", after.DimensionLabel);
            Assert.AreEqual(6, after.VisibleColumns, "Columns beyond the visible cap must be condensed.");
            Assert.AreEqual(3, after.VisibleRows);
            Assert.IsTrue(after.IsCondensed);
        }

        [TestMethod]
        public void LargeValuesCondenseCellsButKeepExactDimensionLabel()
        {
            var state = GridPreviewPlanner.Create("50", "60");

            Assert.IsTrue(state.IsValid);
            Assert.AreEqual(GridPreviewState.MaxVisibleColumns, state.VisibleColumns);
            Assert.AreEqual(GridPreviewState.MaxVisibleRows, state.VisibleRows);
            Assert.AreEqual(GridPreviewState.MaxVisibleColumns * GridPreviewState.MaxVisibleRows, state.CellCount);
            Assert.IsTrue(state.IsCondensed);
            Assert.AreEqual("50 " + MultiplySign + " 60", state.DimensionLabel);
            Assert.AreEqual("Grid preview, 50 columns by 60 rows", state.AutomationName);
        }

        [TestMethod]
        public void BoundaryValuesAreDrawnExactlyAndNotCondensed()
        {
            var atCap = GridPreviewPlanner.Create(6, 4);
            Assert.IsFalse(atCap.IsCondensed);
            Assert.AreEqual(6, atCap.VisibleColumns);
            Assert.AreEqual(4, atCap.VisibleRows);
            Assert.AreEqual(24, atCap.CellCount);

            var beyondCap = GridPreviewPlanner.Create(7, 4);
            Assert.IsTrue(beyondCap.IsCondensed);
            Assert.AreEqual(6, beyondCap.VisibleColumns);
        }

        [TestMethod]
        public void InvalidOrEmptyInputShowsNeutralState()
        {
            var cases = new[]
            {
                GridPreviewPlanner.Create("", "3"),
                GridPreviewPlanner.Create("3", ""),
                GridPreviewPlanner.Create("0", "3"),
                GridPreviewPlanner.Create("3", "-2"),
                GridPreviewPlanner.Create("abc", "3"),
                GridPreviewPlanner.Create(null, "3")
            };

            foreach (var state in cases)
            {
                Assert.IsFalse(state.IsValid);
                Assert.AreEqual(0, state.CellCount);
                Assert.AreEqual(InvalidLabel, state.DimensionLabel);
                Assert.AreEqual("Grid preview, invalid dimensions", state.AutomationName);
            }
        }

        [TestMethod]
        public void FewerShapesThanCellsMarksOnlyArrangedCellsFilled()
        {
            var state = GridPreviewPlanner.Create("3", "6", arrangedShapeCount: 4);

            Assert.IsTrue(state.IsSelectionAware);
            Assert.AreEqual(4, state.ArrangedCellCount);
            Assert.AreEqual("Grid preview, 3 columns by 6 rows, 4 of 18 cells filled", state.AutomationName);
        }

        [TestMethod]
        public void MoreOrEqualShapesThanCellsFillEveryCell()
        {
            var over = GridPreviewPlanner.Create("3", "3", arrangedShapeCount: 20);
            var exact = GridPreviewPlanner.Create("3", "3", arrangedShapeCount: 9);

            Assert.AreEqual(9, over.ArrangedCellCount, "Capacity caps the arranged count.");
            Assert.AreEqual("Grid preview, 3 columns by 3 rows, 9 of 9 cells filled", over.AutomationName);
            Assert.AreEqual(9, exact.ArrangedCellCount);
            Assert.AreEqual("Grid preview, 3 columns by 3 rows, 9 of 9 cells filled", exact.AutomationName);
        }

        [TestMethod]
        public void UnknownShapeCountKeepsAllCellsFilledWithoutSelectionSuffix()
        {
            var state = GridPreviewPlanner.Create("3", "3");

            Assert.IsFalse(state.IsSelectionAware);
            Assert.AreEqual(9, state.ArrangedCellCount);
            Assert.AreEqual("Grid preview, 3 columns by 3 rows", state.AutomationName);
        }

        [TestMethod]
        public void ZeroSelectedShapesMarksAllCellsUnfilled()
        {
            var state = GridPreviewPlanner.Create("3", "3", arrangedShapeCount: 0);

            Assert.IsTrue(state.IsSelectionAware);
            Assert.AreEqual(0, state.ArrangedCellCount);
            Assert.AreEqual("Grid preview, 3 columns by 3 rows, 0 of 9 cells filled", state.AutomationName);
        }

        [TestMethod]
        public void CondensedLargeGridKeepsExactArrangedCountAgainstFullCapacity()
        {
            var state = GridPreviewPlanner.Create("50", "60", arrangedShapeCount: 7);

            Assert.AreEqual(24, state.CellCount, "Visible cells stay capped.");
            Assert.AreEqual(7, state.ArrangedCellCount, "Arranged count reflects the exact selection, not the cap.");
            Assert.AreEqual("Grid preview, 50 columns by 60 rows, 7 of 3000 cells filled", state.AutomationName);
        }
    }
}
