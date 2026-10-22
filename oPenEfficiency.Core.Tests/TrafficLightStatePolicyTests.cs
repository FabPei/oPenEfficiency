using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave-4 coverage for the Traffic Light state contract (QA9-BtnTrafficLight-002):
    /// the persistent state codec "oPE_TrafficLight|{state}|H/V|0/1|{borderWeight}" with the
    /// historical strict validation, used by TrafficLightFeature.
    /// </summary>
    [TestClass]
    public sealed class TrafficLightStatePolicyTests
    {
        [TestMethod]
        public void SerializeRoundTripsEveryField()
        {
            string tag = TrafficLightStatePolicy.Serialize(2, true, false, 1.5f);
            Assert.AreEqual("oPE_TrafficLight|2|H|0|1.5", tag);

            Assert.IsTrue(TrafficLightStatePolicy.TryParse(tag, out int state, out bool horizontal, out bool hasBorder, out float borderWeight));
            Assert.AreEqual(2, state);
            Assert.IsTrue(horizontal);
            Assert.IsFalse(hasBorder);
            Assert.AreEqual(1.5f, borderWeight, 0.001);
        }

        [TestMethod]
        public void TryParseAppliesTheHistoricalStrictValidation()
        {
            Assert.IsFalse(TrafficLightStatePolicy.TryParse(null, out _, out _, out _, out _));
            Assert.IsFalse(TrafficLightStatePolicy.TryParse("oPE_TrafficLight|2|X|0|1", out _, out _, out _, out _));
            Assert.IsFalse(TrafficLightStatePolicy.TryParse("oPE_TrafficLight|2|H|2|1", out _, out _, out _, out _));
            Assert.IsFalse(TrafficLightStatePolicy.TryParse("oPE_TrafficLight|2|H|0|0", out _, out _, out _, out _));
            Assert.IsFalse(TrafficLightStatePolicy.TryParse("oPE_State|2|H|0|1", out _, out _, out _, out _));
        }

        [TestMethod]
        public void PersistedFloatAcceptsTheLegacyCommaDecimalForm()
        {
            Assert.IsTrue(TrafficLightStatePolicy.TryParsePersistedFloat("1,5", out float comma));
            Assert.AreEqual(1.5f, comma, 0.001);
            Assert.IsTrue(TrafficLightStatePolicy.TryParsePersistedFloat("1.5", out float dot));
            Assert.AreEqual(1.5f, dot, 0.001);
        }

        [TestMethod]
        public void ExtraTrailingPartsStillParseLikeTheHistoricalReader()
        {
            Assert.IsTrue(TrafficLightStatePolicy.TryParse("oPE_TrafficLight|1|V|1|2|extra", out int state, out bool horizontal, out bool hasBorder, out _));
            Assert.AreEqual(1, state);
            Assert.IsFalse(horizontal);
            Assert.IsTrue(hasBorder);
        }
    }
}
