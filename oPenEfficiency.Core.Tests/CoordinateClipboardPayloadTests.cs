using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class CoordinateClipboardPayloadTests
    {
        [TestMethod]
        public void SerializeCurrentWritesOnlyVersionedInvariantXyValues()
        {
            Assert.IsTrue(CoordinateClipboardPayload.TrySerializeCurrent(12.5f, -3.25f, out string payload));
            Assert.AreEqual("openEfficiency:XY:v1:12.5,-3.25", payload);
            Assert.AreEqual(2, payload.Split(',').Length);
        }

        [DataTestMethod]
        [DataRow("openEfficiency:XY:v1:12.5,-3.25", 12.5f, -3.25f)]
        [DataRow("12.5,-3.25", 12.5f, -3.25f)]
        [DataRow("12.5,-3.25,100,200", 12.5f, -3.25f)]
        public void ParseForPasteAcceptsCurrentAndLegacyPayloadsButExposesOnlyXy(string payload, float expectedX, float expectedY)
        {
            Assert.IsTrue(CoordinateClipboardPayload.TryParseForPaste(payload, out float x, out float y));
            Assert.AreEqual(expectedX, x);
            Assert.AreEqual(expectedY, y);
        }

        [DataTestMethod]
        [DataRow("")]
        [DataRow("12.5")]
        [DataRow("12.5,-3.25,100")]
        [DataRow("12.5,-3.25,100,200,300")]
        [DataRow("12.5,not-a-number")]
        [DataRow("12.5,-3.25,not-a-number,200")]
        [DataRow("NaN,2")]
        [DataRow("1,Infinity")]
        [DataRow("1,2,3,NaN")]
        [DataRow("openEfficiency:XY:v1:12.5,-3.25,100,200")]
        [DataRow("openEfficiency:XY:v1:NaN,2")]
        [DataRow("openEfficiency:XY:v2:12.5,-3.25")]
        [DataRow("openEfficiency:XY:12.5,-3.25")]
        public void ParseForPasteRejectsMalformedNonfiniteOrUnknownVersionPayloads(string payload)
        {
            Assert.IsFalse(CoordinateClipboardPayload.TryParseForPaste(payload, out _, out _));
        }

        [DataTestMethod]
        [DataRow("openEfficiency:XY:v1:12.5,-3.25", true)]
        [DataRow("1,2", false)]
        [DataRow("1,2,3,4", false)]
        [DataRow("openEfficiency:XY:v1:1,2,3,4", false)]
        [DataRow("openEfficiency:XY:v1:NaN,2", false)]
        [DataRow("openEfficiency:XY:v2:1,2", false)]
        public void OwnershipRecognizesOnlyStructurallyValidCurrentPayloads(string payload, bool expectedOwnership)
        {
            Assert.AreEqual(expectedOwnership, CoordinateClipboardPayload.IsOwnedCurrentPayload(payload));
        }

        [TestMethod]
        public void SerializeRejectsNonfiniteCoordinates()
        {
            Assert.IsFalse(CoordinateClipboardPayload.TrySerializeCurrent(float.NaN, 2f, out _));
            Assert.IsFalse(CoordinateClipboardPayload.TrySerializeCurrent(1f, float.PositiveInfinity, out _));
        }
    }
}
