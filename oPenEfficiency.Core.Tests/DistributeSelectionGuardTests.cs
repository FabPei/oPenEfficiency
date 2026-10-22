using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 1 acceptance coverage for the selection contract of BtnDistributeHorizontal and
    /// BtnDistributeVertical. Review round 6c measured that the third batch asserted only
    /// metadata -> SidebarStateEvaluator: rewriting the commands' own guard (`sr.Count &lt; 3 &amp;&amp;
    /// !relativeToSlide` -> `sr.Count &lt; 5`) left every scenario green. This test executes the host-neutral
    /// production rules the two Execute methods use, and the Catalog tests
    /// Wave1DistributeHorizontal|VerticalSelectionGuardIsTheProductionRule assert that they really delegate to them.
    /// Owner decision a1 for REG-DISTRIBUTE-SCOPE-001 is part of the contract: with exactly two selected shapes the
    /// primary click spreads them across the slide instead of doing nothing, while a single shape stays refused.
    /// The COM leg (ShapeRange.Distribute, the Undo entry) stays the office-automated leg.
    /// </summary>
    [TestClass]
    public sealed class DistributeSelectionGuardTests
    {
        [TestMethod]
        public void DistributeSelectionGuardMatchesTheCommandRoute()
        {
            Assert.AreEqual(3, DistributeSelectionRule.MinimumShapesForPrimaryRoute,
                "The primary route distributes interior gaps from three shapes on.");
            Assert.AreEqual(2, DistributeSelectionRule.MinimumShapesForAcrossSlideRoute,
                "The across-slide route distributes two shapes over the slide.");

            Assert.IsFalse(DistributeSelectionRule.MeetsMinimum(0, false), "An empty selection must be refused.");
            Assert.IsFalse(DistributeSelectionRule.MeetsMinimum(1, false), "One shape cannot be distributed.");
            Assert.IsFalse(DistributeSelectionRule.MeetsMinimum(2, false),
                "The primary route alone refuses two shapes - exactly why the route fallback exists.");
            Assert.IsTrue(DistributeSelectionRule.MeetsMinimum(3, false), "Three shapes must be distributed.");
            Assert.IsTrue(DistributeSelectionRule.MeetsMinimum(4, false), "Four shapes must be distributed.");

            Assert.IsFalse(DistributeSelectionRule.MeetsMinimum(0, true), "An empty selection must be refused.");
            Assert.IsFalse(DistributeSelectionRule.MeetsMinimum(1, true), "One shape cannot be distributed.");
            Assert.IsTrue(DistributeSelectionRule.MeetsMinimum(2, true),
                "The across-slide route must accept the two shapes its metadata minimum promises.");
            Assert.IsTrue(DistributeSelectionRule.MeetsMinimum(3, true), "Three shapes must be distributed.");

            Assert.IsTrue(DistributeSelectionRule.UsesAcrossSlideRoute(2, false),
                "Owner decision a1: at two selected shapes the primary click spreads them across the slide.");
            Assert.IsFalse(DistributeSelectionRule.UsesAcrossSlideRoute(3, false),
                "From three shapes the primary click distributes the interior gaps.");
            Assert.IsFalse(DistributeSelectionRule.UsesAcrossSlideRoute(6, false),
                "From three shapes the primary click distributes the interior gaps.");
            Assert.IsTrue(DistributeSelectionRule.UsesAcrossSlideRoute(2, true),
                "An explicit across-slide request keeps its route.");
            Assert.IsTrue(DistributeSelectionRule.UsesAcrossSlideRoute(6, true),
                "An explicit across-slide request keeps its route.");

            Assert.IsTrue(DistributeSelectionRule.MeetsMinimum(2, DistributeSelectionRule.UsesAcrossSlideRoute(2, false)),
                "The button is enabled at two shapes, so the resolved route must accept them (REG-DISTRIBUTE-SCOPE-001).");
            Assert.IsFalse(DistributeSelectionRule.MeetsMinimum(1, DistributeSelectionRule.UsesAcrossSlideRoute(1, false)),
                "A single selected shape must stay refused on every route.");

            // Review round 6f: the sampled values above left room for a mutated threshold or a special case in the
            // rule body. The whole range is pinned now, by contract rather than by a second copy of the numbers: on
            // the primary route the resolved route must accept exactly the counts the across-slide minimum promises,
            // and each route must accept exactly the counts its own minimum promises.
            for (var count = 0; count <= 6; count++)
            {
                var resolved = DistributeSelectionRule.UsesAcrossSlideRoute(count, false);
                Assert.AreEqual(
                    count <= DistributeSelectionRule.MinimumShapesForAcrossSlideRoute,
                    resolved,
                    $"The primary click must resolve to the across-slide route for at most {DistributeSelectionRule.MinimumShapesForAcrossSlideRoute} shapes (count {count}).");
                Assert.AreEqual(
                    count >= DistributeSelectionRule.MinimumShapesForAcrossSlideRoute,
                    DistributeSelectionRule.MeetsMinimum(count, true),
                    $"The across-slide route must accept from {DistributeSelectionRule.MinimumShapesForAcrossSlideRoute} shapes on (count {count}).");
                Assert.AreEqual(
                    count >= DistributeSelectionRule.MinimumShapesForPrimaryRoute,
                    DistributeSelectionRule.MeetsMinimum(count, false),
                    $"The primary route must accept from {DistributeSelectionRule.MinimumShapesForPrimaryRoute} shapes on (count {count}).");
                Assert.AreEqual(
                    count >= DistributeSelectionRule.MinimumShapesForAcrossSlideRoute,
                    DistributeSelectionRule.MeetsMinimum(count, resolved),
                    $"An enabled Distribute click must act from {DistributeSelectionRule.MinimumShapesForAcrossSlideRoute} shapes on (count {count}).");
                Assert.IsTrue(
                    DistributeSelectionRule.UsesAcrossSlideRoute(count, true),
                    $"An explicit across-slide request must keep its route for every count (count {count}).");
            }

            // Review round 6g: the plan is the whole decision the command applies, and it is executed here over a
            // wide range instead of a sampled one - a threshold hidden behind a local constant (measured in review
            // round 6g: `if (selectedShapeCount > Six) return true;`) cannot hide inside the pinned range any more.
            for (var count = 0; count <= 64; count++)
            {
                foreach (var requested in new[] { false, true })
                {
                    var plan = DistributeSelectionRule.ResolvePlan(count, requested);
                    Assert.AreEqual(
                        DistributeSelectionRule.UsesAcrossSlideRoute(count, requested),
                        plan.RelativeToSlide,
                        $"The plan must carry the route the rule resolved (count {count}, requested across-slide {requested}).");
                    Assert.AreEqual(
                        DistributeSelectionRule.MeetsMinimum(count, plan.RelativeToSlide),
                        plan.Distribute,
                        $"The plan must accept exactly what its own route accepts (count {count}, requested across-slide {requested}).");
                }

                var primary = DistributeSelectionRule.ResolvePlan(count, false);
                Assert.AreEqual(
                    count >= DistributeSelectionRule.MinimumShapesForAcrossSlideRoute,
                    primary.Distribute,
                    $"An enabled Distribute click must act from {DistributeSelectionRule.MinimumShapesForAcrossSlideRoute} shapes on (count {count}).");
                Assert.AreEqual(
                    count <= DistributeSelectionRule.MinimumShapesForAcrossSlideRoute,
                    primary.RelativeToSlide,
                    $"The primary click must spread at most {DistributeSelectionRule.MinimumShapesForAcrossSlideRoute} shapes across the slide (count {count}).");
                var acrossSlide = DistributeSelectionRule.ResolvePlan(count, true);
                Assert.IsTrue(acrossSlide.RelativeToSlide,
                    $"An explicit across-slide request must keep its route in the plan (count {count}).");
                Assert.AreEqual(
                    count >= DistributeSelectionRule.MinimumShapesForAcrossSlideRoute,
                    acrossSlide.Distribute,
                    $"The across-slide route must accept exactly the counts its minimum allows (count {count}) - a refusal hidden behind an arithmetic expression must not survive the oracle (review round 6i).");
            }
        }
    }
}
