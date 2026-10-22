using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave-5 coverage for the Checkbox state contract (QA9-BtnCheckbox-002): the
    /// persistent AlternativeText codec with its culture-safe write and legacy-lenient read,
    /// the Wingdings symbol map and the centered insert placement, used by CheckboxFeature.
    /// </summary>
    [TestClass]
    public sealed class CheckboxStatePolicyTests
    {
        [TestMethod]
        public void FormatTagWritesTheDocumentedTagShape()
        {
            Assert.AreEqual("oPE_Checkbox|0|True|24|0",
                CheckboxStatePolicy.FormatTag(0, true, 24f, 0x000000));
            Assert.AreEqual("oPE_Checkbox|2|False|18.5|-1",
                CheckboxStatePolicy.FormatTag(2, false, 18.5f, -1));
        }

        [TestMethod]
        public void FormatTagWritesInvariantSizesEvenInCurrentCultureDecks()
        {
            CultureInfo original = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
                // A fractional font size must never carry a locale decimal comma into the tag.
                Assert.AreEqual("oPE_Checkbox|1|True|24.5|-1",
                    CheckboxStatePolicy.FormatTag(1, true, 24.5f, -1));
            }
            finally
            {
                CultureInfo.CurrentCulture = original;
            }
        }

        [TestMethod]
        public void ParseTagFieldsReadsWrittenTagsIncludingTheColorSentinel()
        {
            CheckboxStatePolicy.ParseTagFields("oPE_Checkbox|1|True|24.5|-1",
                out string stateText, out bool hasFrame, out float fontSize, out int colorRgb);
            Assert.AreEqual("1", stateText);
            Assert.IsTrue(hasFrame);
            Assert.AreEqual(24.5f, fontSize, 0.001);
            Assert.AreEqual(-1, colorRgb);
        }

        [TestMethod]
        public void ParseTagFieldsReadsLegacyCurrentCultureSizeTags()
        {
            CultureInfo original = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
                // Decks authored before the invariant write carry a comma decimal; the lenient
                // reader falls back to the current culture so the size survives.
                CheckboxStatePolicy.ParseTagFields("oPE_Checkbox|0|False|24,5|255",
                    out _, out _, out float legacySize, out _);
                Assert.AreEqual(24.5f, legacySize, 0.001);
            }
            finally
            {
                CultureInfo.CurrentCulture = original;
            }
        }

        [TestMethod]
        public void ParseTagFieldsLeavesDefaultsForMissingOrShortTags()
        {
            CheckboxStatePolicy.ParseTagFields(null, out string nullStateText, out _, out _, out _);
            Assert.IsNull(nullStateText);

            CheckboxStatePolicy.ParseTagFields("oPE_Checkbox", out string shortStateText, out _, out _, out _);
            Assert.IsNull(shortStateText);

            CheckboxStatePolicy.ParseTagFields("oPE_Checkbox|0", out string partialStateText, out _, out _, out _);
            Assert.IsNull(partialStateText);
        }

        [TestMethod]
        public void ParseTagFieldsFallsBackToZeroValuesForMalformedFields()
        {
            CheckboxStatePolicy.ParseTagFields("oPE_Checkbox|x|yes|abc|zz",
                out string stateText, out bool hasFrame, out float fontSize, out int colorRgb);
            Assert.AreEqual("x", stateText);
            Assert.IsFalse(hasFrame);  // bool.TryParse accepts only True/False
            Assert.AreEqual(0f, fontSize, 0.001);
            Assert.AreEqual(0, colorRgb);
        }

        [TestMethod]
        public void GetSymbolMapsTheThreeStatesAndTheUnknownDefault()
        {
            Assert.AreEqual("ü", CheckboxStatePolicy.GetSymbol(0));
            Assert.AreEqual("û", CheckboxStatePolicy.GetSymbol(1));
            Assert.AreEqual(" ", CheckboxStatePolicy.GetSymbol(2));
            Assert.AreEqual("", CheckboxStatePolicy.GetSymbol(99));
        }

        [TestMethod]
        public void CenterOnSlideCentersTheInsertBox()
        {
            CheckboxStatePolicy.CenterOnSlide(960f, 540f, 50f, out float left, out float top);
            Assert.AreEqual(455f, left, 0.001);
            Assert.AreEqual(245f, top, 0.001);

            CheckboxStatePolicy.CenterOnSlide(0f, 0f, 50f, out float zeroedLeft, out float zeroedTop);
            Assert.AreEqual(-25f, zeroedLeft, 0.001);
            Assert.AreEqual(-25f, zeroedTop, 0.001);
        }

        [TestMethod]
        public void IsCheckboxTagRecognizesOnlyTheTagPrefix()
        {
            Assert.IsFalse(CheckboxStatePolicy.IsCheckboxTag(null));
            Assert.IsFalse(CheckboxStatePolicy.IsCheckboxTag(""));
            Assert.IsFalse(CheckboxStatePolicy.IsCheckboxTag("oPE_HarveyBall|0.25"));
            Assert.IsTrue(CheckboxStatePolicy.IsCheckboxTag("oPE_Checkbox|0|True|24|0"));
            Assert.IsTrue(CheckboxStatePolicy.IsCheckboxTag("oPE_CheckboxAnything"));
        }
    }
}
