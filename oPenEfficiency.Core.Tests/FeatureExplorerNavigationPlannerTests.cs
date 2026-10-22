using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.UI;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class FeatureExplorerNavigationPlannerTests
    {
        [TestMethod]
        public void OverflowUsesTheNoChromeHostCapacityInBothDirections()
        {
            var beforeThreshold = FeatureExplorerNavigationPlanner.Create(1000, 1000, 936, 0);
            var afterThreshold = FeatureExplorerNavigationPlanner.Create(1000.51, 1000, 936, 0);
            var collapsedAgain = FeatureExplorerNavigationPlanner.Create(1000, 1000, 936, 180);

            Assert.IsFalse(beforeThreshold.RequiresNavigation,
                "A 64-DIP chevron reduction must not create overflow when the full category row fits.");
            Assert.IsTrue(afterThreshold.RequiresNavigation);
            Assert.IsFalse(collapsedAgain.RequiresNavigation,
                "Removing overflow must collapse navigation even while the rendered viewport still reflects chevrons.");
            Assert.AreEqual(0d, collapsedAgain.HorizontalOffset, FeatureExplorerNavigationPlanner.Epsilon);
        }

        [TestMethod]
        public void OverflowThresholdIsStableAtTheToleranceBoundary()
        {
            Assert.IsFalse(FeatureExplorerNavigationPlanner.Create(1000.5, 1000, 936, 0).RequiresNavigation);
            Assert.IsTrue(FeatureExplorerNavigationPlanner.Create(1000.51, 1000, 936, 0).RequiresNavigation);
        }

        [TestMethod]
        public void NavigationButtonsReflectStartMiddleAndEndOffsets()
        {
            var start = FeatureExplorerNavigationPlanner.Create(1200, 1000, 936, 0);
            var middle = FeatureExplorerNavigationPlanner.Create(1200, 1000, 936, 100);
            var end = FeatureExplorerNavigationPlanner.Create(1200, 1000, 936, 9999);

            Assert.IsFalse(start.CanScrollLeft);
            Assert.IsTrue(start.CanScrollRight);
            Assert.IsTrue(middle.CanScrollLeft);
            Assert.IsTrue(middle.CanScrollRight);
            Assert.AreEqual(264d, end.HorizontalOffset, FeatureExplorerNavigationPlanner.Epsilon);
            Assert.IsTrue(end.CanScrollLeft);
            Assert.IsFalse(end.CanScrollRight);
        }

        [TestMethod]
        public void RevealOffsetMovesOnlyItemsOutsideTheViewport()
        {
            Assert.AreEqual(100d, FeatureExplorerNavigationPlanner.CalculateRevealOffset(1000, 300, 200, 100, 50));
            Assert.AreEqual(200d, FeatureExplorerNavigationPlanner.CalculateRevealOffset(1000, 300, 200, 350, 50));
            Assert.AreEqual(400d, FeatureExplorerNavigationPlanner.CalculateRevealOffset(1000, 300, 200, 600, 100));
        }

        [TestMethod]
        public void RevealOffsetClampsAndDoesNotScrollWhenContentFits()
        {
            Assert.AreEqual(0d, FeatureExplorerNavigationPlanner.CalculateRevealOffset(1000, 300, 200, -50, 20));
            Assert.AreEqual(700d, FeatureExplorerNavigationPlanner.CalculateRevealOffset(1000, 300, 200, 950, 200));
            Assert.AreEqual(0d, FeatureExplorerNavigationPlanner.CalculateRevealOffset(300, 300, 80, 250, 40));
        }
    }
}
