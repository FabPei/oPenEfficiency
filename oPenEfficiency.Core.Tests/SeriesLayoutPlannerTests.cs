using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave-5 coverage for Generate Shape Series layout (QA9-BtnSeriesGenerator-002): the
    /// required series extent per direction, the shrink-only auto-fit scale and the per-item
    /// placement, used by SeriesGeneratorFeature.
    /// </summary>
    [TestClass]
    public sealed class SeriesLayoutPlannerTests
    {
        [TestMethod]
        public void RequiredSizeFollowsTheDirectionLayout()
        {
            SeriesLayoutPlanner.RequiredSize("Horizontal", 4, 3, 50f, 20f, 10f, 10f,
                out float horizontalWidth, out float horizontalHeight);
            Assert.AreEqual(230f, horizontalWidth, 0.001); // 4 items + 3 gaps
            Assert.AreEqual(20f, horizontalHeight, 0.001);

            SeriesLayoutPlanner.RequiredSize("Vertical", 4, 3, 50f, 20f, 10f, 10f,
                out float verticalWidth, out float verticalHeight);
            Assert.AreEqual(50f, verticalWidth, 0.001);
            Assert.AreEqual(110f, verticalHeight, 0.001); // 4 items + 3 gaps
        }

        [TestMethod]
        public void RequiredSizeForMatrixCapsColumnsAndCeilsRows()
        {
            SeriesLayoutPlanner.RequiredSize("Matrix", 7, 3, 30f, 10f, 5f, 2f,
                out float width, out float height);
            Assert.AreEqual(100f, width, 0.001);  // 3 columns + 2 gaps
            Assert.AreEqual(34f, height, 0.001);  // 3 ceiled rows + 2 gaps

            SeriesLayoutPlanner.RequiredSize("Matrix", 2, 3, 30f, 10f, 5f, 2f,
                out float shortRowWidth, out float shortRowHeight);
            Assert.AreEqual(65f, shortRowWidth, 0.001);  // columns capped at 2
            Assert.AreEqual(10f, shortRowHeight, 0.001); // single row
        }

        [TestMethod]
        public void RequiredSizeReportsZeroExtentForAnUnknownDirection()
        {
            SeriesLayoutPlanner.RequiredSize("Diagonal", 4, 3, 50f, 20f, 10f, 10f,
                out float width, out float height);
            Assert.AreEqual(0f, width, 0.001);
            Assert.AreEqual(0f, height, 0.001);
        }

        [TestMethod]
        public void AutoFitScaleShrinksOnTheBindingAxisOnly()
        {
            Assert.AreEqual(0.5f, SeriesLayoutPlanner.AutoFitScale(400f, 100f, 200f, 150f), 0.001);
            Assert.AreEqual(0.25f, SeriesLayoutPlanner.AutoFitScale(400f, 200f, 200f, 50f), 0.001);
        }

        [TestMethod]
        public void AutoFitScaleNeverEnlargesAndGuardsZeroRequirements()
        {
            Assert.AreEqual(1f, SeriesLayoutPlanner.AutoFitScale(100f, 100f, 500f, 500f), 0.001);
            Assert.AreEqual(1f, SeriesLayoutPlanner.AutoFitScale(0f, 0f, 100f, 100f), 0.001);
            // A zero requirement never binds; the other axis still shrinks.
            Assert.AreEqual(0.8f, SeriesLayoutPlanner.AutoFitScale(0f, 50f, 100f, 40f), 0.001);
        }

        [TestMethod]
        public void PlaceItemWalksHorizontalAndVerticalRuns()
        {
            SeriesLayoutPlanner.PlaceItem("Horizontal", 2, 3, 100f, 70f, 50f, 40f, 10f, 10f,
                out float runLeft, out float runTop);
            Assert.AreEqual(220f, runLeft, 0.001); // start + 2 * (width + spacing)
            Assert.AreEqual(70f, runTop, 0.001);

            SeriesLayoutPlanner.PlaceItem("Vertical", 3, 3, 100f, 70f, 50f, 40f, 10f, 10f,
                out float columnLeft, out float columnTop);
            Assert.AreEqual(100f, columnLeft, 0.001);
            Assert.AreEqual(220f, columnTop, 0.001); // start + 3 * (height + spacing)
        }

        [TestMethod]
        public void PlaceItemWrapsMatrixRowsAndColumns()
        {
            SeriesLayoutPlanner.PlaceItem("Matrix", 0, 3, 100f, 70f, 50f, 40f, 10f, 20f,
                out float firstLeft, out float firstTop);
            Assert.AreEqual(100f, firstLeft, 0.001);
            Assert.AreEqual(70f, firstTop, 0.001);

            SeriesLayoutPlanner.PlaceItem("Matrix", 4, 3, 100f, 70f, 50f, 40f, 10f, 20f,
                out float wrappedLeft, out float wrappedTop);
            Assert.AreEqual(160f, wrappedLeft, 0.001);  // column 1
            Assert.AreEqual(130f, wrappedTop, 0.001);   // row 1

            // Historical degenerate cases: a matrix without columns and an unknown
            // direction anchor every item at the start.
            SeriesLayoutPlanner.PlaceItem("Matrix", 2, 0, 100f, 70f, 50f, 40f, 10f, 20f,
                out float noColumnsLeft, out float noColumnsTop);
            Assert.AreEqual(100f, noColumnsLeft, 0.001);
            Assert.AreEqual(70f, noColumnsTop, 0.001);

            SeriesLayoutPlanner.PlaceItem("Diagonal", 2, 3, 100f, 70f, 50f, 40f, 10f, 20f,
                out float unknownLeft, out float unknownTop);
            Assert.AreEqual(100f, unknownLeft, 0.001);
            Assert.AreEqual(70f, unknownTop, 0.001);
        }
    }
}
