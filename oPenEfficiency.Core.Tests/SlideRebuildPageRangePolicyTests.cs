using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 external-data-wave coverage for the Local Slide Rebuild page-range selection
    /// (QA9-BtnLocalSlideRebuild-002): the accepted grammar, bounds and span cap, used by
    /// LocalSlideRebuildFeature.
    /// </summary>
    [TestClass]
    public sealed class SlideRebuildPageRangePolicyTests
    {
        [TestMethod]
        public void ParsesSinglePagesRangesAndMixtures()
        {
            Assert.IsTrue(SlideRebuildPageRangePolicy.TryParsePageRange("3", out HashSet<uint> single));
            CollectionAssert.AreEquivalent(new[] { 3u }, new List<uint>(single));

            Assert.IsTrue(SlideRebuildPageRangePolicy.TryParsePageRange("4-7", out HashSet<uint> range));
            CollectionAssert.AreEquivalent(new[] { 4u, 5u, 6u, 7u }, new List<uint>(range));

            Assert.IsTrue(SlideRebuildPageRangePolicy.TryParsePageRange("1, 3, 6-8", out HashSet<uint> mixed));
            CollectionAssert.AreEquivalent(new[] { 1u, 3u, 6u, 7u, 8u }, new List<uint>(mixed));
        }

        [TestMethod]
        public void DuplicatesCollapseIntoTheSet()
        {
            Assert.IsTrue(SlideRebuildPageRangePolicy.TryParsePageRange("3,3,5-5", out HashSet<uint> pages));
            CollectionAssert.AreEquivalent(new[] { 3u, 5u }, new List<uint>(pages));
        }

        [TestMethod]
        public void EnforcesThePageBoundsAndSpanCap()
        {
            Assert.IsFalse(SlideRebuildPageRangePolicy.TryParsePageRange("0", out _));
            Assert.IsFalse(SlideRebuildPageRangePolicy.TryParsePageRange("10001", out _));
            Assert.IsFalse(SlideRebuildPageRangePolicy.TryParsePageRange("5-2", out _));
            Assert.IsFalse(SlideRebuildPageRangePolicy.TryParsePageRange("1-1001", out _));  // span 1000 > 999
            Assert.IsTrue(SlideRebuildPageRangePolicy.TryParsePageRange("1-1000", out HashSet<uint> full));
            Assert.AreEqual(1000, full.Count);
            Assert.IsTrue(SlideRebuildPageRangePolicy.TryParsePageRange("9999-10000", out HashSet<uint> tail));
            CollectionAssert.AreEquivalent(new[] { 9999u, 10000u }, new List<uint>(tail));
        }

        [TestMethod]
        public void RejectsMalformedInput()
        {
            Assert.IsFalse(SlideRebuildPageRangePolicy.TryParsePageRange(null, out _));
            Assert.IsFalse(SlideRebuildPageRangePolicy.TryParsePageRange("", out _));
            Assert.IsFalse(SlideRebuildPageRangePolicy.TryParsePageRange("  ", out _));
            Assert.IsFalse(SlideRebuildPageRangePolicy.TryParsePageRange("1,,2", out _));
            Assert.IsFalse(SlideRebuildPageRangePolicy.TryParsePageRange("1-2-3", out _));
            Assert.IsFalse(SlideRebuildPageRangePolicy.TryParsePageRange("abc", out _));
            Assert.IsFalse(SlideRebuildPageRangePolicy.TryParsePageRange("-4", out _));
        }
    }
}
