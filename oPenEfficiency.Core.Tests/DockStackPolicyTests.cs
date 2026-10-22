using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 align-and-size-wave coverage for the dock stacking family
    /// (QA9-BtnDockTop-002, QA9-BtnDockLeft-002, QA9-BtnDockRight-002, QA9-BtnDockBottom-002):
    /// the zero-gap stack cursor against the reference shape's edge, used by the four
    /// DockToFirst*Feature features.
    /// </summary>
    [TestClass]
    public sealed class DockStackPolicyTests
    {
        // Reference rectangle: left 200, top 100, 80x60.
        private const float RL = 200f, RT = 100f, RW = 80f, RH = 60f;

        [TestMethod]
        public void AboveStacksUpwardDisplacingEachShapeByItsOwnHeight()
        {
            var positions = DockStackPolicy.StackSequence(
                DockStackPolicy.DockDirection.Above, RL, RT, RW, RH, new float[] { 30f, 20f });
            Assert.AreEqual(70f, positions[0], 0.001);  // 100 - 30
            Assert.AreEqual(50f, positions[1], 0.001);  // 70 - 20 (cursor follows the placed edge)
        }

        [TestMethod]
        public void LeftStacksLeftwardDisplacingEachShapeByItsOwnWidth()
        {
            var positions = DockStackPolicy.StackSequence(
                DockStackPolicy.DockDirection.Left, RL, RT, RW, RH, new float[] { 50f, 25f });
            Assert.AreEqual(150f, positions[0], 0.001); // 200 - 50
            Assert.AreEqual(125f, positions[1], 0.001); // 150 - 25
        }

        [TestMethod]
        public void RightPlacesTheFirstShapeOnTheAnchorEdgeThenAdvances()
        {
            var positions = DockStackPolicy.StackSequence(
                DockStackPolicy.DockDirection.Right, RL, RT, RW, RH, new float[] { 40f, 10f, 5f });
            Assert.AreEqual(280f, positions[0], 0.001); // anchor right edge
            Assert.AreEqual(320f, positions[1], 0.001); // 280 + 40
            Assert.AreEqual(330f, positions[2], 0.001); // 320 + 10
        }

        [TestMethod]
        public void BelowPlacesTheFirstShapeOnTheAnchorEdgeThenAdvances()
        {
            var positions = DockStackPolicy.StackSequence(
                DockStackPolicy.DockDirection.Below, RL, RT, RW, RH, new float[] { 15f, 45f });
            Assert.AreEqual(160f, positions[0], 0.001); // anchor bottom edge
            Assert.AreEqual(175f, positions[1], 0.001); // 160 + 15
        }

        [TestMethod]
        public void SingleShapeAndMixedSizesStackWithZeroGap()
        {
            var single = DockStackPolicy.StackSequence(
                DockStackPolicy.DockDirection.Right, RL, RT, RW, RH, new float[] { 33f });
            Assert.AreEqual(280f, single[0], 0.001);

            var mixed = DockStackPolicy.StackSequence(
                DockStackPolicy.DockDirection.Above, RL, RT, RW, RH, new float[] { 10f, 10f, 10f });
            Assert.AreEqual(90f, mixed[0], 0.001);
            Assert.AreEqual(80f, mixed[1], 0.001);
            Assert.AreEqual(70f, mixed[2], 0.001);
        }

        [TestMethod]
        public void EmptySelectionYieldsNoPositions()
        {
            var positions = DockStackPolicy.StackSequence(
                DockStackPolicy.DockDirection.Below, RL, RT, RW, RH, new float[0]);
            Assert.AreEqual(0, positions.Length);
        }
    }
}
