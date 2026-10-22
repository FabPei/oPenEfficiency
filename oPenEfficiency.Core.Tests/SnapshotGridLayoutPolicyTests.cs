using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 external-data-wave coverage for the Insert Slide Snapshots grid layout
    /// (QA9-BtnSlidePaste-002): the iterative shrink loop, the safety minimum and the final
    /// column count, used by SlidePasteWindow.
    /// </summary>
    [TestClass]
    public sealed class SnapshotGridLayoutPolicyTests
    {
        private const float Gap = 0.3f * SnapshotGridLayoutPolicy.CmToPt;
        private const float Width = 7.11f * SnapshotGridLayoutPolicy.CmToPt;
        private const float Height = 4.0f * SnapshotGridLayoutPolicy.CmToPt;

        [TestMethod]
        public void SmallSelectionKeepsTheDefaultSize()
        {
            // A single 7.11x4cm snapshot fits comfortably on a 33.87x19.05cm slide.
            SnapshotGridLayoutPolicy.FitGrid(1, 33.87f * SnapshotGridLayoutPolicy.CmToPt, 19.05f * SnapshotGridLayoutPolicy.CmToPt,
                Gap, Width, Height, Width / Height,
                out float w, out float h, out int columns);
            Assert.AreEqual(Width, w, 0.01);
            Assert.AreEqual(Height, h, 0.01);
            Assert.AreEqual(4, columns); // floor((960.2+8.5)/(201.6+8.5)) on a 960pt slide
        }

        [TestMethod]
        public void ManyRowsShrinkTheHeightUntilTheGridFits()
        {
            float availH = 19.05f * SnapshotGridLayoutPolicy.CmToPt;
            SnapshotGridLayoutPolicy.FitGrid(40, 33.87f * SnapshotGridLayoutPolicy.CmToPt, availH,
                Gap, Width, Height, Width / Height,
                out float w, out float h, out int columns);
            // The default size cannot fit 40 snapshots; the loop must have shrunk the height.
            Assert.IsTrue(h < Height, "height was never shrunk");
            Assert.IsTrue(w < Width, "width must follow the height via the ratio");
            Assert.IsTrue(w / h > (Width / Height) - 0.001, "the aspect ratio is preserved");
            // With the final size the rows must fit the available height.
            int rows = (40 + columns - 1) / columns;
            float totalHeight = rows * h + (rows - 1) * Gap;
            Assert.IsTrue(totalHeight <= availH + 0.01, "the grid still overflows after fitting");
        }

        [TestMethod]
        public void SafetyMinimumStopsTheShrinkBelowTheFloor()
        {
            // An absurdly small area forces the shrink to run into the ~0.5cm floor. The
            // historical loop breaks AFTER the shrink that crosses the floor, so the final
            // height sits just below it (within one 5% step), not exactly on it.
            SnapshotGridLayoutPolicy.FitGrid(500, 100f, 30f, 2f,
                Width, Height, Width / Height,
                out _, out float h, out _);
            Assert.IsTrue(h < SnapshotGridLayoutPolicy.MinimumHeightPt, "the shrink never crossed the floor");
            Assert.IsTrue(h >= SnapshotGridLayoutPolicy.MinimumHeightPt * 0.95f, "the shrink ran past one step below the floor");
        }

        [TestMethod]
        public void ColumnsNeverDropBelowOne()
        {
            Assert.AreEqual(1, SnapshotGridLayoutPolicy.ColumnsThatFit(10f, 500f, 2f));
            Assert.AreEqual(1, SnapshotGridLayoutPolicy.ColumnsThatFit(1001f, 500f, 2f)); // (1001+2)/502 floors to 1
            Assert.AreEqual(2, SnapshotGridLayoutPolicy.ColumnsThatFit(1004f, 500f, 2f)); // (1004+2)/502 floors to 2
        }
    }
}
