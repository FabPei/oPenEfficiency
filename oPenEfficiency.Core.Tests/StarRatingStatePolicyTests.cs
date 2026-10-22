using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave-4 coverage for the Star Rating state contract (QA9-BtnStarRating-002):
    /// the persistent state codec "oPE_StarRating|{count}|{rating}|{iconMode}|{hasFrame}|{colorBgr}"
    /// and the per-star fill decision, used by StarRatingFeature.
    /// </summary>
    [TestClass]
    public sealed class StarRatingStatePolicyTests
    {
        [TestMethod]
        public void TagRoundTripsEveryField()
        {
            string tag = StarRatingStatePolicy.FormatTag(5, 3.5, 2, true, 123456);
            Assert.AreEqual("oPE_StarRating|5|3.5|2|1|123456", tag);

            Assert.IsTrue(StarRatingStatePolicy.TryParseTag(tag, out int count, out double rating, out int iconMode, out bool hasFrame, out int colorBgr));
            Assert.AreEqual(5, count);
            Assert.AreEqual(3.5, rating, 0.001);
            Assert.AreEqual(2, iconMode);
            Assert.IsTrue(hasFrame);
            Assert.AreEqual(123456, colorBgr);
        }

        [TestMethod]
        public void TryParseTagRejectsForeignOrMalformedStates()
        {
            Assert.IsFalse(StarRatingStatePolicy.TryParseTag(null, out _, out _, out _, out _, out _));
            Assert.IsFalse(StarRatingStatePolicy.TryParseTag("oPE_StarRating|5|3.5|2", out _, out _, out _, out _, out _));
            Assert.IsFalse(StarRatingStatePolicy.TryParseTag("oPE_HarveyBall|5|3.5|2|1|123456", out _, out _, out _, out _, out _));
            Assert.IsFalse(StarRatingStatePolicy.TryParseTag("oPE_StarRating|5|3.5|2|2|123456", out _, out _, out _, out _, out _));
        }

        [TestMethod]
        public void ResolveFillCoversFullDimAndPartialStars()
        {
            StarRatingStatePolicy.ResolveFill(0, 2.0, out var k0, out _);
            Assert.AreEqual(StarRatingStatePolicy.FillKind.Full, k0);
            StarRatingStatePolicy.ResolveFill(2, 2.0, out var k1, out _);
            Assert.AreEqual(StarRatingStatePolicy.FillKind.Dim, k1);
            StarRatingStatePolicy.ResolveFill(2, 2.5, out var k2, out double fraction);
            Assert.AreEqual(StarRatingStatePolicy.FillKind.Gradient, k2);
            Assert.AreEqual(0.5, fraction, 0.001);
        }

        [TestMethod]
        public void SolidFallbackMatchesTheHistoricalHalfThreshold()
        {
            Assert.IsTrue(StarRatingStatePolicy.UseSolidFallback(0.5));
            Assert.IsTrue(StarRatingStatePolicy.UseSolidFallback(0.75));
            Assert.IsFalse(StarRatingStatePolicy.UseSolidFallback(0.49));
        }
    }
}
