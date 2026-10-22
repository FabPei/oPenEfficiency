using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 6 acceptance coverage for the z-order commands (BtnBringForward, BtnSendBackward,
    /// BtnBringToFront, BtnSendToBack). Each command decides its target stacking order in
    /// <see cref="ZOrderRankPlanner"/>, the host-neutral production code driven here; the COM leg
    /// (applying the planned order via Shape.ZOrder and the PowerPoint Undo entry) is the
    /// office-automated leg of the same feature and is not covered by this test. All fixture orders are
    /// bottom-to-top (the last element ends up topmost).
    /// </summary>
    [TestClass]
    public sealed class ZOrderRankPlannerTests
    {
        [TestMethod]
        public void StepPlanPutsTheFirstSelectedShapeOnTopOfTheSelection()
        {
            // Bring Forward (ranked variant): one-step commands fill the recaptured selection slots
            // with the ranked selection; first-selected-on-top ranks it last (topmost).
            List<int> plan = ZOrderRankPlanner.PlanRanked(
                new List<int> { 11, 1, 12, 2 },
                new List<int> { 12, 11 },
                firstSelectedOnTop: true,
                bringToFront: false,
                sendToBack: false);
            CollectionAssert.AreEqual(new List<int> { 11, 1, 12, 2 }, plan);

            List<int> inverted = ZOrderRankPlanner.PlanRanked(
                new List<int> { 11, 1, 12, 2 },
                new List<int> { 12, 11 },
                firstSelectedOnTop: false,
                bringToFront: false,
                sendToBack: false);
            CollectionAssert.AreEqual(new List<int> { 12, 1, 11, 2 }, inverted);
        }

        [TestMethod]
        public void StepPlanPutsTheFirstSelectedShapeAtTheBottomOfTheSelection()
        {
            // Send Backward (ranked variant): the same slot fill with the opposite ranking keeps the
            // first selected shape at the bottom of the selection and never disturbs unselected shapes.
            List<int> plan = ZOrderRankPlanner.PlanRanked(
                new List<int> { 11, 1, 12, 2 },
                new List<int> { 12, 11 },
                firstSelectedOnTop: false,
                bringToFront: false,
                sendToBack: false);
            CollectionAssert.AreEqual(new List<int> { 12, 1, 11, 2 }, plan);
            Assert.AreEqual(1, plan.IndexOf(1), "The unselected shape 1 must keep its slot.");
            Assert.AreEqual(3, plan.IndexOf(2), "The unselected shape 2 must keep its slot.");
        }

        [TestMethod]
        public void RankedPlanBringsTheSelectionToFrontWithTheFirstSelectedOnTop()
        {
            List<int> plan = ZOrderRankPlanner.PlanRanked(
                new List<int> { 1, 11, 2, 12 },
                new List<int> { 12, 11 },
                firstSelectedOnTop: true,
                bringToFront: true,
                sendToBack: false);
            CollectionAssert.AreEqual(new List<int> { 1, 2, 11, 12 }, plan);
            Assert.AreEqual(12, plan[plan.Count - 1], "The first selected shape must end up topmost.");
        }

        [TestMethod]
        public void RankedPlanSendsTheSelectionToBackWithTheFirstSelectedAtBottom()
        {
            List<int> plan = ZOrderRankPlanner.PlanRanked(
                new List<int> { 1, 11, 2, 12 },
                new List<int> { 12, 11 },
                firstSelectedOnTop: false,
                bringToFront: false,
                sendToBack: true);
            CollectionAssert.AreEqual(new List<int> { 12, 11, 1, 2 }, plan);
            Assert.AreEqual(12, plan[0], "The first selected shape must end up bottommost.");
        }

        [TestMethod]
        public void RelativePlanPlacesMoversDirectlyAboveTheReference()
        {
            List<int> plan = ZOrderRankPlanner.PlanRelativeToReference(
                new List<int> { 11, 1, 12, 2, 3 },
                new List<int> { 11, 12 },
                referenceId: 3,
                placeAbove: true);
            CollectionAssert.AreEqual(new List<int> { 1, 2, 3, 11, 12 }, plan);
        }

        [TestMethod]
        public void RelativePlanPlacesMoversDirectlyBelowTheReferenceAndRejectsMissingAnchors()
        {
            List<int> plan = ZOrderRankPlanner.PlanRelativeToReference(
                new List<int> { 11, 1, 12, 2, 3 },
                new List<int> { 11, 12 },
                referenceId: 3,
                placeAbove: false);
            CollectionAssert.AreEqual(new List<int> { 1, 2, 11, 12, 3 }, plan);

            Assert.IsNull(ZOrderRankPlanner.PlanRelativeToReference(
                new List<int> { 11, 1 },
                new List<int> { 11 },
                referenceId: 99,
                placeAbove: true),
                "A missing reference must be rejected, not guessed.");
            Assert.IsNull(ZOrderRankPlanner.PlanRelativeToReference(
                new List<int> { 11, 1 },
                new List<int>(),
                referenceId: 1,
                placeAbove: true),
                "An empty mover set must be rejected.");
        }
    }
}
