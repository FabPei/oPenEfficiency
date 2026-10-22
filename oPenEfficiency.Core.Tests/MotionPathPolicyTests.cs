using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;
using static oPenEfficiency.Core.Services.MotionPathPolicy;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 arrange-wave coverage for Create Motion Path (QA9-BtnCreateMotionPath-002):
    /// the mode expansion, the slide-relative center delta and the invariant path string
    /// used by CreateMotionPathFeature.
    /// </summary>
    [TestClass]
    public sealed class MotionPathPolicyTests
    {
        [TestMethod]
        public void EveryModeExpandsToItsDocumentedDirections()
        {
            Assert.IsTrue(MovesFirstToSecond(MotionPathMode.FirstToSecond));
            Assert.IsFalse(MovesSecondToFirst(MotionPathMode.FirstToSecond));

            Assert.IsFalse(MovesFirstToSecond(MotionPathMode.SecondToFirst));
            Assert.IsTrue(MovesSecondToFirst(MotionPathMode.SecondToFirst));

            Assert.IsTrue(MovesFirstToSecond(MotionPathMode.BothDirections));
            Assert.IsTrue(MovesSecondToFirst(MotionPathMode.BothDirections));
        }

        [TestMethod]
        public void TheDeltaIsMeasuredBetweenTheShapeCentersInSlideUnits()
        {
            // Centers: (125,120) -> (325,240) on an 800x400 slide.
            ComputeNormalizedDelta(
                100f, 100f, 50f, 40f,
                300f, 220f, 50f, 40f,
                800f, 400f,
                out float pctX, out float pctY);

            Assert.AreEqual(0.25, pctX, 1e-6);
            Assert.AreEqual(0.3, pctY, 1e-6);
        }

        [TestMethod]
        public void AShapeMovingBackwardsProducesANegativeDelta()
        {
            ComputeNormalizedDelta(
                300f, 220f, 50f, 40f,
                100f, 100f, 50f, 40f,
                800f, 400f,
                out float pctX, out float pctY);

            Assert.AreEqual(-0.25, pctX, 1e-6);
            Assert.AreEqual(-0.3, pctY, 1e-6);
        }

        [TestMethod]
        public void ThePathStringUsesInvariantCultureCoordinates()
        {
            Assert.AreEqual("M 0 0 L 0.25 -0.125 E", BuildPathString(0.25f, -0.125f));
            Assert.AreEqual("M 0 0 L 0 0 E", BuildPathString(0f, 0f));
        }
    }
}
