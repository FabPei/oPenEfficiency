using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave-5 coverage for the margin paste plan (QA9-BtnSetMargin-002): the context
    /// menu's paste combinations, the learned-data requirement and the reset rule, used by
    /// SetMarginFeature.
    /// </summary>
    [TestClass]
    public sealed class MarginPastePolicyTests
    {
        [TestMethod]
        public void TryBuildPlanRefusesWithoutLearnedDataUnlessReset()
        {
            Assert.IsFalse(MarginPastePolicy.TryBuildPlan(true, true, false,
                hasLearnedData: false, 3f, 4f, 5f, 6f, out _));

            Assert.IsTrue(MarginPastePolicy.TryBuildPlan(true, true, true,
                hasLearnedData: false, 3f, 4f, 5f, 6f, out _));
        }

        [TestMethod]
        public void TryBuildPlanPastesAllFourLearnedValues()
        {
            Assert.IsTrue(MarginPastePolicy.TryBuildPlan(true, true, false,
                hasLearnedData: true, 3f, 4f, 5f, 6f, out MarginPastePolicy.MarginPlan plan));
            Assert.IsTrue(plan.ApplyHorizontal);
            Assert.IsTrue(plan.ApplyVertical);
            Assert.AreEqual(3f, plan.Left, 0.001);
            Assert.AreEqual(4f, plan.Right, 0.001);
            Assert.AreEqual(5f, plan.Top, 0.001);
            Assert.AreEqual(6f, plan.Bottom, 0.001);
        }

        [TestMethod]
        public void TryBuildPlanResolvesTheSingleAxisSelections()
        {
            Assert.IsTrue(MarginPastePolicy.TryBuildPlan(true, false, false,
                hasLearnedData: true, 3f, 4f, 5f, 6f, out MarginPastePolicy.MarginPlan horizontal));
            Assert.IsTrue(horizontal.ApplyHorizontal);
            Assert.IsFalse(horizontal.ApplyVertical);
            Assert.AreEqual(3f, horizontal.Left, 0.001);
            Assert.AreEqual(4f, horizontal.Right, 0.001);

            Assert.IsTrue(MarginPastePolicy.TryBuildPlan(false, true, false,
                hasLearnedData: true, 3f, 4f, 5f, 6f, out MarginPastePolicy.MarginPlan vertical));
            Assert.IsFalse(vertical.ApplyHorizontal);
            Assert.IsTrue(vertical.ApplyVertical);
            Assert.AreEqual(5f, vertical.Top, 0.001);
            Assert.AreEqual(6f, vertical.Bottom, 0.001);
        }

        [TestMethod]
        public void TryBuildPlanResetWritesZerosForTheSelectedAxesOnly()
        {
            Assert.IsTrue(MarginPastePolicy.TryBuildPlan(true, true, true,
                hasLearnedData: false, 3f, 4f, 5f, 6f, out MarginPastePolicy.MarginPlan full));
            Assert.AreEqual(0f, full.Left, 0.001);
            Assert.AreEqual(0f, full.Right, 0.001);
            Assert.AreEqual(0f, full.Top, 0.001);
            Assert.AreEqual(0f, full.Bottom, 0.001);

            // Reset honors the axis selection: a horizontal reset resolves zeros for every
            // axis but only writes the horizontal pair — ApplyVertical stays off, so the
            // target shape's top/bottom margins are never written.
            Assert.IsTrue(MarginPastePolicy.TryBuildPlan(true, false, true,
                hasLearnedData: true, 3f, 4f, 5f, 6f, out MarginPastePolicy.MarginPlan horizontal));
            Assert.AreEqual(0f, horizontal.Left, 0.001);
            Assert.AreEqual(0f, horizontal.Right, 0.001);
            Assert.IsFalse(horizontal.ApplyVertical);
        }
    }
}
