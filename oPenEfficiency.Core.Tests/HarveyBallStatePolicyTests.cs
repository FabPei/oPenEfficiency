using System.Globalization;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave-4 coverage for the Harvey Ball state contract (QA9-BtnHarveyBall-002):
    /// the 0 -> 0.25 -> 0.5 -> 0.75 -> 1.0 -> 0 cycle and the AlternativeText state codec,
    /// including the culture-safety fix (legacy de-DE decks wrote "oPE_HarveyBall|0,25").
    /// </summary>
    [TestClass]
    public sealed class HarveyBallStatePolicyTests
    {
        [TestMethod]
        public void NextPercentageWalksTheFullLadderAndWraps()
        {
            Assert.AreEqual(0.25f, HarveyBallStatePolicy.NextPercentage(0f));
            Assert.AreEqual(0.5f, HarveyBallStatePolicy.NextPercentage(0.25f));
            Assert.AreEqual(0.75f, HarveyBallStatePolicy.NextPercentage(0.5f));
            Assert.AreEqual(1.0f, HarveyBallStatePolicy.NextPercentage(0.75f));
            Assert.AreEqual(0f, HarveyBallStatePolicy.NextPercentage(1.0f));
        }

        [TestMethod]
        public void NextPercentageSnapsOffLadderValuesOntoTheLadder()
        {
            Assert.AreEqual(0.5f, HarveyBallStatePolicy.NextPercentage(0.3f));
            Assert.AreEqual(0.25f, HarveyBallStatePolicy.NextPercentage(0.1f));
            Assert.AreEqual(0f, HarveyBallStatePolicy.NextPercentage(1.4f));
        }

        [TestMethod]
        public void TagRoundTripsEveryLadderValue()
        {
            foreach (var value in new[] { 0f, 0.25f, 0.5f, 0.75f, 1f })
            {
                string tag = HarveyBallStatePolicy.FormatTag(value);
                Assert.AreEqual("oPE_HarveyBall|" + value.ToString("0.####", CultureInfo.InvariantCulture), tag);
                Assert.IsTrue(HarveyBallStatePolicy.TryParseTag(tag, out float parsed));
                Assert.AreEqual(value, parsed);
            }
        }

        [TestMethod]
        public void TryParseTagAcceptsLegacyCultureFormattedStates()
        {
            var original = Thread.CurrentThread.CurrentCulture;
            try
            {
                foreach (var cultureName in new[] { "en-US", "de-DE" })
                {
                    Thread.CurrentThread.CurrentCulture = new CultureInfo(cultureName);
                    foreach (var value in new[] { "0,25", "0,5", "0,75" })
                    {
                        string invariant = value.Replace(',', '.');
                        Assert.IsTrue(HarveyBallStatePolicy.TryParseTag("oPE_HarveyBall|" + value, out float comma));
                        Assert.AreEqual(float.Parse(invariant, CultureInfo.InvariantCulture), comma, cultureName + " comma-decimal state " + value);
                        Assert.IsTrue(HarveyBallStatePolicy.TryParseTag("oPE_HarveyBall|" + invariant, out float dot));
                        Assert.AreEqual(float.Parse(invariant, CultureInfo.InvariantCulture), dot, cultureName + " invariant-decimal state " + invariant);
                    }
                }
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = original;
            }
        }

        [TestMethod]
        public void TryParseTagRejectsMixedOrRepeatedDecimalSeparators()
        {
            var original = Thread.CurrentThread.CurrentCulture;
            try
            {
                foreach (var cultureName in new[] { "en-US", "de-DE" })
                {
                    Thread.CurrentThread.CurrentCulture = new CultureInfo(cultureName);
                    Assert.IsFalse(HarveyBallStatePolicy.TryParseTag("oPE_HarveyBall|0,25.0", out _), cultureName);
                    Assert.IsFalse(HarveyBallStatePolicy.TryParseTag("oPE_HarveyBall|0,25,0", out _), cultureName);
                    Assert.IsFalse(HarveyBallStatePolicy.TryParseTag("oPE_HarveyBall|1,000", out _), cultureName);
                    Assert.IsFalse(HarveyBallStatePolicy.TryParseTag("oPE_HarveyBall|1,25", out _), cultureName);
                }
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = original;
            }
        }

        [TestMethod]
        public void TryParseTagRejectsForeignOrMalformedStates()
        {
            Assert.IsFalse(HarveyBallStatePolicy.TryParseTag(null, out _));
            Assert.IsFalse(HarveyBallStatePolicy.TryParseTag("", out _));
            Assert.IsFalse(HarveyBallStatePolicy.TryParseTag("oPE_StarRating|3", out _));
            Assert.IsFalse(HarveyBallStatePolicy.TryParseTag("oPE_HarveyBall", out _));
            Assert.IsFalse(HarveyBallStatePolicy.TryParseTag("oPE_HarveyBall|abc", out _));
        }

        [TestMethod]
        public void CodecIsCultureIndependentUnderADecimalCommaCulture()
        {
            var original = Thread.CurrentThread.CurrentCulture;
            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
                string tag = HarveyBallStatePolicy.FormatTag(0.25f);
                Assert.AreEqual("oPE_HarveyBall|0.25", tag);
                Assert.IsTrue(HarveyBallStatePolicy.TryParseTag(tag, out float parsed));
                Assert.AreEqual(0.25f, parsed);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = original;
            }
        }
    }
}
