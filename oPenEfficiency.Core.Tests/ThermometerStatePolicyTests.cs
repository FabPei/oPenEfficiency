using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave-4 coverage for the Thermometer state contract (QA9-BtnThermometer-002):
    /// the 0.25-step fill cycle with the 1.0 -> 0 wrap, the clamping rule and the persistent
    /// state codec (historical tag literal "oPen_Thermometer" preserved), used by
    /// ThermometerFeature.
    /// </summary>
    [TestClass]
    public sealed class ThermometerStatePolicyTests
    {
        [TestMethod]
        public void NextPercentageStepsAndWrapsLikeTheClickCycle()
        {
            Assert.AreEqual(0.75f, ThermometerStatePolicy.NextPercentage(0.5f));
            Assert.AreEqual(1f, ThermometerStatePolicy.NextPercentage(0.9f));
            Assert.AreEqual(0f, ThermometerStatePolicy.NextPercentage(1f));
        }

        [TestMethod]
        public void ClampPercentageKeepsValuesInTheUnitInterval()
        {
            Assert.AreEqual(0f, ThermometerStatePolicy.ClampPercentage(-2f));
            Assert.AreEqual(1f, ThermometerStatePolicy.ClampPercentage(3f));
            Assert.AreEqual(0.42f, ThermometerStatePolicy.ClampPercentage(0.42f));
        }

        [TestMethod]
        public void TagRoundTripsAndKeepsTheHistoricalLiteral()
        {
            Assert.AreEqual("oPen_Thermometer", ThermometerStatePolicy.ThermometerTag);
            string tag = ThermometerStatePolicy.FormatTag(0.75f, 255);
            Assert.AreEqual("oPen_Thermometer|0.75|255", tag);

            ThermometerStatePolicy.ParseState(tag, out float percentage, out int colorRgb);
            Assert.AreEqual(0.75f, percentage, 0.001);
            Assert.AreEqual(255, colorRgb);
        }

        [TestMethod]
        public void ParseStateIsLenientLikeTheHistoricalReader()
        {
            ThermometerStatePolicy.ParseState(null, out float defaultPercentage, out int defaultColor);
            Assert.AreEqual(0.5f, defaultPercentage, 0.001);
            Assert.AreEqual(0x0000FF, defaultColor);

            ThermometerStatePolicy.ParseState("foreign|2.5", out float clamped, out _);
            Assert.AreEqual(1f, clamped, 0.001);

            ThermometerStatePolicy.ParseState("EE4P-indicator-text", out float fallback, out _);
            Assert.AreEqual(0.5f, fallback, 0.001);
        }
    }
}
