using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave-4 coverage for the Progress Series state contract (QA9-BtnProgressSeries-002):
    /// the persistent state codec "oPE_Series|{count}|{activeIndex}|{mode}|{layout}|{color}|{shape}"
    /// with the historical lenient partial parse and documented defaults, used by
    /// ProgressSeriesFeature.SeriesSettings.
    /// </summary>
    [TestClass]
    public sealed class ProgressSeriesStatePolicyTests
    {
        [TestMethod]
        public void TagRoundTripsEveryField()
        {
            string tag = ProgressSeriesStatePolicy.FormatTag(7, 4, "Numbers", "V", "#0078D4", "Circle");
            Assert.AreEqual("oPE_Series|7|4|Numbers|V|#0078D4|Circle", tag);

            Assert.IsTrue(ProgressSeriesStatePolicy.TryParseTag(tag, out int count, out int activeIndex, out string mode, out string layout, out string color, out string shape));
            Assert.AreEqual(7, count);
            Assert.AreEqual(4, activeIndex);
            Assert.AreEqual("Numbers", mode);
            Assert.AreEqual("V", layout);
            Assert.AreEqual("#0078D4", color);
            Assert.AreEqual("Circle", shape);
        }

        [TestMethod]
        public void PartialTagsOverwriteOnlyThePresentFields()
        {
            Assert.IsTrue(ProgressSeriesStatePolicy.TryParseTag("oPE_Series|9", out int count, out int activeIndex, out string mode, out string layout, out string color, out string shape));
            Assert.AreEqual(9, count);
            Assert.AreEqual(ProgressSeriesStatePolicy.DefaultActiveIndex, activeIndex);
            Assert.AreEqual(ProgressSeriesStatePolicy.DefaultMode, mode);
            Assert.AreEqual(ProgressSeriesStatePolicy.DefaultLayout, layout);
            Assert.AreEqual(ProgressSeriesStatePolicy.DefaultHighlightColor, color);
            Assert.AreEqual(ProgressSeriesStatePolicy.DefaultShapeType, shape);
        }

        [TestMethod]
        public void UnparsableIntegersFallToZeroLikeTheHistoricalTryParse()
        {
            Assert.IsTrue(ProgressSeriesStatePolicy.TryParseTag("oPE_Series|abc", out int count, out _, out _, out _, out _, out _));
            Assert.AreEqual(0, count);
        }

        [TestMethod]
        public void ForeignTagsReturnFalseAndKeepDefaults()
        {
            Assert.IsFalse(ProgressSeriesStatePolicy.TryParseTag(null, out int count, out _, out _, out _, out _, out _));
            Assert.AreEqual(ProgressSeriesStatePolicy.DefaultCount, count);
            Assert.IsFalse(ProgressSeriesStatePolicy.TryParseTag("oPE_StarRating|5", out _, out _, out _, out _, out _, out _));
        }
    }
}
