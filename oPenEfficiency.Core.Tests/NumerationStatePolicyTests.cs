using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave-5 coverage for the numbered marker state contract (QA9-BtnNumeration-002):
    /// the persistent AlternativeText codec with its defaults, clamps and lenient read, and the
    /// replacement placement, used by NumerationFeature.
    /// </summary>
    [TestClass]
    public sealed class NumerationStatePolicyTests
    {
        [TestMethod]
        public void DefaultsRoundTripThroughTheTagCodec()
        {
            NumerationStatePolicy.Defaults(out int number, out string shapeType, out float size, out string colorHex, out int alpha);
            Assert.AreEqual(1, number);
            Assert.AreEqual("Circle", shapeType);
            Assert.AreEqual(22.68f, size, 0.001);
            Assert.AreEqual("#005C94", colorHex);
            Assert.AreEqual(255, alpha);

            string tag = NumerationStatePolicy.FormatTag(number, shapeType, size, colorHex, alpha);
            Assert.AreEqual("oPE_Numeration|1|Circle|22.68|#005C94|255", tag);

            NumerationStatePolicy.ParseState(tag, out int readNumber, out string readShape, out float readSize, out string readColor, out int readAlpha);
            Assert.AreEqual(1, readNumber);
            Assert.AreEqual("Circle", readShape);
            Assert.AreEqual(22.68f, readSize, 0.001);
            Assert.AreEqual("#005C94", readColor);
            Assert.AreEqual(255, readAlpha);
        }

        [TestMethod]
        public void FormatTagWritesInvariantNumbersAndKeepsNullFieldsOnDefaults()
        {
            Assert.AreEqual(
                "oPE_Numeration|7|Diamond|30.5|#FF8800|10",
                NumerationStatePolicy.FormatTag(7, "Diamond", 30.5f, "#FF8800", 10));
            Assert.AreEqual(
                "oPE_Numeration|2|Circle|12|#005C94|255",
                NumerationStatePolicy.FormatTag(2, null, 12f, null, 255));
        }

        [TestMethod]
        public void ParseStateRequiresTheTagPrefixAndReportsDefaultsOtherwise()
        {
            NumerationStatePolicy.ParseState("some other marker text",
                out int number, out string shapeType, out float size, out string colorHex, out int alpha);
            Assert.AreEqual(1, number);
            Assert.AreEqual("Circle", shapeType);
            Assert.AreEqual(22.68f, size, 0.001);
            Assert.AreEqual("#005C94", colorHex);
            Assert.AreEqual(255, alpha);

            NumerationStatePolicy.ParseState(null,
                out number, out shapeType, out size, out colorHex, out alpha);
            Assert.AreEqual(1, number);
        }

        [TestMethod]
        public void ParseStateAppliesOnlyThePresentLenientFields()
        {
            NumerationStatePolicy.ParseState("oPE_Numeration|4|Diamond",
                out int number, out string shapeType, out float size, out string colorHex, out int alpha);
            Assert.AreEqual(4, number);
            Assert.AreEqual("Diamond", shapeType);
            Assert.AreEqual(22.68f, size, 0.001);
            Assert.AreEqual("#005C94", colorHex);
            Assert.AreEqual(255, alpha);
        }

        [TestMethod]
        public void ParseStateEnforcesTheFiveToTwoHundredPointSizeWindow()
        {
            NumerationStatePolicy.ParseState("oPE_Numeration|1|Circle|4.9|#005C94|255",
                out _, out _, out float tooSmall, out _, out _);
            Assert.AreEqual(22.68f, tooSmall, 0.001);

            NumerationStatePolicy.ParseState("oPE_Numeration|1|Circle|200|#005C94|255",
                out _, out _, out float atMaximum, out _, out _);
            Assert.AreEqual(200f, atMaximum, 0.001);

            NumerationStatePolicy.ParseState("oPE_Numeration|1|Circle|200.1|#005C94|255",
                out _, out _, out float tooLarge, out _, out _);
            Assert.AreEqual(22.68f, tooLarge, 0.001);
        }

        [TestMethod]
        public void ParseStateClampsAlphaAndKeepsOnlyParsableColors()
        {
            NumerationStatePolicy.ParseState("oPE_Numeration|1|Circle|20|#FF8800|300",
                out _, out _, out _, out string keptColor, out int clampedHigh);
            Assert.AreEqual("#FF8800", keptColor);
            Assert.AreEqual(255, clampedHigh);

            NumerationStatePolicy.ParseState("oPE_Numeration|1|Circle|20|#FF8800|-5",
                out _, out _, out _, out _, out int clampedLow);
            Assert.AreEqual(0, clampedLow);

            NumerationStatePolicy.ParseState("oPE_Numeration|1|Circle|20|not-a-color|100",
                out _, out _, out _, out string fallbackColor, out _);
            Assert.AreEqual("#005C94", fallbackColor);

            NumerationStatePolicy.ParseState("oPE_Numeration|1|Circle|20| |100",
                out _, out _, out _, out string emptyColor, out _);
            Assert.AreEqual("#005C94", emptyColor);
        }

        [TestMethod]
        public void CenterOnShapePlacesTheReplacementOnThePredecessorCenter()
        {
            NumerationStatePolicy.CenterOnShape(100f, 50f, 80f, 40f, 22.68f,
                out float left, out float top);
            Assert.AreEqual(128.66f, left, 0.001);
            Assert.AreEqual(58.66f, top, 0.001);

            NumerationStatePolicy.CenterOnShape(0f, 0f, 0f, 0f, 10f,
                out float zeroedLeft, out float zeroedTop);
            Assert.AreEqual(-5f, zeroedLeft, 0.001);
            Assert.AreEqual(-5f, zeroedTop, 0.001);
        }
    }
}
