using System.Globalization;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 3 coverage for the Font Panel decision rules (QA9-BtnFontPanel-001): the size
    /// validity bounds, the change detection that decides which ranges an apply touches, and the
    /// culture-dependent size parsing, used by FontPanel.
    /// </summary>
    [TestClass]
    public sealed class FontEditPolicyTests
    {
        [TestMethod]
        public void FontSizeValidityIsExclusiveOnBothBounds()
        {
            Assert.IsFalse(FontEditPolicy.IsValidFontSize(0f), "0pt is rejected.");
            Assert.IsFalse(FontEditPolicy.IsValidFontSize(-3f), "Negative sizes are rejected.");
            Assert.IsFalse(FontEditPolicy.IsValidFontSize(4000f), "4000pt is rejected (exclusive upper bound).");
            Assert.IsTrue(FontEditPolicy.IsValidFontSize(0.01f), "Tiny positive sizes are accepted.");
            Assert.IsTrue(FontEditPolicy.IsValidFontSize(3999.99f), "Sizes just below 4000 are accepted.");
        }

        [TestMethod]
        public void SizeChangeUsesTheInclusiveNoChangeBoundary()
        {
            Assert.IsFalse(FontEditPolicy.NeedsSizeChange(0f, 0.01f), "A difference of exactly the epsilon counts as unchanged.");
            Assert.IsFalse(FontEditPolicy.NeedsSizeChange(12f, 12f), "An equal size needs no change.");
            Assert.IsTrue(FontEditPolicy.NeedsSizeChange(12f, 12.01f),
                "Float reality: 12.01f rounds above the 0.01f epsilon, so production treats 12 -> 12.01 as a change.");
            Assert.IsTrue(FontEditPolicy.NeedsSizeChange(18f, 12f), "A large difference needs a change.");
        }

        [TestMethod]
        public void NameChangeIsCaseInsensitive()
        {
            Assert.IsFalse(FontEditPolicy.NeedsNameChange("Arial", "ARIAL"), "Case-only differences are unchanged.");
            Assert.IsFalse(FontEditPolicy.NeedsNameChange("Arial", "Arial"), "An equal name needs no change.");
            Assert.IsTrue(FontEditPolicy.NeedsNameChange("Arial", "Calibri"), "A different name needs a change.");
            Assert.IsTrue(FontEditPolicy.NeedsNameChange(null, "Arial"), "A missing name needs a change.");
        }

        [TestMethod]
        public void SizeParsingFollowsTheUsersCulture()
        {
            var original = Thread.CurrentThread.CurrentCulture;
            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
                Assert.IsTrue(FontEditPolicy.TryParseSize("12,5", out float comma), "de-DE users type a decimal comma.");
                Assert.AreEqual(12.5f, comma, 0.001f);

                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                Assert.IsTrue(FontEditPolicy.TryParseSize("12.5", out float dot), "en-US users type a decimal point.");
                Assert.AreEqual(12.5f, dot, 0.001f);

                Assert.IsFalse(FontEditPolicy.TryParseSize("not-a-size", out _), "Non-numeric input is rejected.");
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = original;
            }
        }
    }
}
