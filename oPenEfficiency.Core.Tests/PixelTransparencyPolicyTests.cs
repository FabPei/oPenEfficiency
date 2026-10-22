using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 3 coverage for the Transparent Color pixel rule (QA9-BtnTransparentColor-001):
    /// the squared Cartesian RGB limit of the tolerance and the per-pixel decision, used by
    /// TransparentColorWindow.
    /// </summary>
    [TestClass]
    public sealed class PixelTransparencyPolicyTests
    {
        [TestMethod]
        public void LimitSquaredFollowsTheCartesianToleranceCurve()
        {
            Assert.AreEqual(0f, PixelTransparencyPolicy.LimitSquared(0d), 0.0001f);
            Assert.AreEqual(3f * 255f * 255f * 0.25f, PixelTransparencyPolicy.LimitSquared(0.5d), 0.0001f);
            Assert.AreEqual(3f * 255f * 255f, PixelTransparencyPolicy.LimitSquared(1d), 0.0001f);
        }

        [TestMethod]
        public void LimitSquaredClampsTheToleranceIntoTheUnitInterval()
        {
            Assert.AreEqual(0f, PixelTransparencyPolicy.LimitSquared(-3d), 0.0001f);
            Assert.AreEqual(3f * 255f * 255f, PixelTransparencyPolicy.LimitSquared(4d), 0.0001f);
        }

        [TestMethod]
        public void ExactMatchesBecomeTransparentAtZeroTolerance()
        {
            Assert.IsTrue(PixelTransparencyPolicy.ShouldMakeTransparent(10, 20, 30, 255, 10, 20, 30, 0d));
            Assert.IsFalse(PixelTransparencyPolicy.ShouldMakeTransparent(11, 20, 30, 255, 10, 20, 30, 0d));
        }

        [TestMethod]
        public void AlreadyTransparentPixelsAreNeverTouched()
        {
            Assert.IsFalse(PixelTransparencyPolicy.ShouldMakeTransparent(10, 20, 30, 0, 10, 20, 30, 1d));
        }

        [TestMethod]
        public void TheDistanceBoundaryIsInclusive()
        {
            // distance 1 on one channel only -> distSq = 1; tolerance 1/255 gives limitSq = 3 exactly
            double tolerance = 1d / 255d;
            Assert.AreEqual(3f, PixelTransparencyPolicy.LimitSquared(tolerance), 0.0001f);
            Assert.IsTrue(PixelTransparencyPolicy.ShouldMakeTransparent(11, 20, 30, 255, 10, 20, 30, tolerance));
            // distance (1,1,1) -> distSq = 3 == limit -> still inside
            Assert.IsTrue(PixelTransparencyPolicy.ShouldMakeTransparent(11, 21, 31, 255, 10, 20, 30, tolerance));
            // distance (2,0,0) -> distSq = 4 > limit -> stays opaque
            Assert.IsFalse(PixelTransparencyPolicy.ShouldMakeTransparent(12, 20, 30, 255, 10, 20, 30, tolerance));
        }

        [TestMethod]
        public void FullToleranceMakesEveryOpaquePixelTransparent()
        {
            Assert.IsTrue(PixelTransparencyPolicy.ShouldMakeTransparent(0, 0, 0, 255, 255, 255, 255, 1d));
            Assert.IsTrue(PixelTransparencyPolicy.ShouldMakeTransparent(255, 0, 128, 1, 0, 255, 64, 1d));
        }
    }
}
