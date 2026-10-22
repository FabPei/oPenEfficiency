using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 dialogs/workflows-wave coverage for Master Covers & Overlays
    /// (QA9-BtnFlightMode-002): the persistent tag vocabulary and the overlay transparency
    /// levels with their Light/Heavy classification boundary, used by FlightModeFeature.
    /// </summary>
    [TestClass]
    public sealed class FlightModeStatePolicyTests
    {
        [TestMethod]
        public void OverlayTransparencyLevelsClassifyLightAndHeavy()
        {
            Assert.IsTrue(FlightModeStatePolicy.IsLightOverlay(FlightModeStatePolicy.LightOverlayTransparency));
            Assert.IsFalse(FlightModeStatePolicy.IsLightOverlay(FlightModeStatePolicy.HeavyOverlayTransparency));
        }

        [TestMethod]
        public void ClassificationBoundaryIsInclusiveAtNineTenths()
        {
            Assert.IsTrue(FlightModeStatePolicy.IsLightOverlay(0.9f));
            Assert.IsTrue(FlightModeStatePolicy.IsLightOverlay(0.96f));
            Assert.IsFalse(FlightModeStatePolicy.IsLightOverlay(0.89f));
            Assert.IsFalse(FlightModeStatePolicy.IsLightOverlay(0f));
        }

        [TestMethod]
        public void WrittenLevelsRoundTripThroughTheWrittenValues()
        {
            // A freshly written Light overlay must read back as Light, a Heavy one as Heavy.
            float lightWritten = FlightModeStatePolicy.LightOverlayTransparency;
            float heavyWritten = FlightModeStatePolicy.HeavyOverlayTransparency;
            Assert.AreEqual(0.95f, lightWritten, 0.0001);
            Assert.AreEqual(0.85f, heavyWritten, 0.0001);
            Assert.IsTrue(FlightModeStatePolicy.IsLightOverlay(lightWritten));
            Assert.IsFalse(FlightModeStatePolicy.IsLightOverlay(heavyWritten));
        }

        [TestMethod]
        public void TagVocabularyIsStable()
        {
            Assert.AreEqual("oPE_FlightMode", FlightModeStatePolicy.FlightModeTag);
            Assert.AreEqual("ImageCover", FlightModeStatePolicy.ImageCoverTag);
            Assert.AreEqual("Overlay", FlightModeStatePolicy.OverlayTag);
        }
    }
}
