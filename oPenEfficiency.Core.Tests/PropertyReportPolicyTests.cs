using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 dialogs/workflows-wave coverage for Inspect Fonts &amp; Colors
    /// (QA9-BtnPropertyExtraction-002): the COLORREF hex conversion, the property
    /// normalization rules and the report layout, used by PropertyExtractionFeature.
    /// </summary>
    [TestClass]
    public sealed class PropertyReportPolicyTests
    {
        [TestMethod]
        public void ToHexDecodesTheColorrefByteOrder()
        {
            Assert.AreEqual("#000000", PropertyReportPolicy.ToHex(0x000000));
            Assert.AreEqual("#FFFFFF", PropertyReportPolicy.ToHex(0xFFFFFF));
            Assert.AreEqual("#FF0000", PropertyReportPolicy.ToHex(0x0000FF)); // pure red COLORREF
            Assert.AreEqual("#204080", PropertyReportPolicy.ToHex(0x804020));
        }

        [TestMethod]
        public void NormalizeFontsDeduplicatesSortsAndDropsBlanks()
        {
            CollectionAssert.AreEqual(
                new[] { "Arial", "Calibri", "Segoe UI" },
                PropertyReportPolicy.NormalizeFonts(new[] { "Calibri", " ", null, "Arial", "Calibri", "Segoe UI" }));
        }

        [TestMethod]
        public void NormalizeSizesAndColorsDropInvalidValues()
        {
            CollectionAssert.AreEqual(
                new float[] { 12f, 18.5f, 24f },
                PropertyReportPolicy.NormalizeSizes(new float[] { 24f, 0f, 18.5f, -3f, 12f, 24f }));
            CollectionAssert.AreEqual(
                new[] { 0x000000, 0x804020, 0xFFFFFF },
                PropertyReportPolicy.NormalizeColors(new[] { 0xFFFFFF, -1, 0x000000, 0x804020, 0x000000 }));
        }

        [TestMethod]
        public void BuildReportFormatsEveryGroupWithTheDocumentedLayout()
        {
            // The feature normalizes (sorts/dedups) before reporting; sizes render in the
            // current culture, exactly like the historical inline implementation.
            string report = PropertyReportPolicy.BuildReport(
                "Selected shapes",
                new[] { "Arial", "Calibri" },
                new float[] { 18.5f, 24f },
                new[] { 0x0000FF },
                new int[] { },
                new[] { 0x804020 });
            string nl = System.Environment.NewLine;
            Assert.AreEqual(
                "Scope: Selected shapes" + nl +
                "Fonts: Arial, Calibri" + nl +
                "Font sizes: " + 18.5f.ToString("0.##") + " pt, " + 24f.ToString("0.##") + " pt" + nl +
                "Text colors: #FF0000" + nl +
                "Fill colors: None" + nl +
                "Line colors: #204080",
                report);
        }

        [TestMethod]
        public void SizeFormattingFollowsTheCurrentCulture()
        {
            System.Globalization.CultureInfo original = System.Globalization.CultureInfo.CurrentCulture;
            try
            {
                System.Globalization.CultureInfo.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo("de-DE");
                string report = PropertyReportPolicy.BuildReport(
                    "Current slide", new string[0], new float[] { 18.5f, 24f }, new int[0], new int[0], new int[0]);
                StringAssert.Contains(report, "Font sizes: 18,5 pt, 24 pt");
            }
            finally
            {
                System.Globalization.CultureInfo.CurrentCulture = original;
            }
        }

        [TestMethod]
        public void EmptyScopeReportsNoneEverywhereAndTrimsTrailingWhitespace()
        {
            string report = PropertyReportPolicy.BuildReport(
                "Current slide", new string[0], new float[0], new int[0], new int[0], new int[0]);
            string nl = System.Environment.NewLine;
            Assert.AreEqual(
                "Scope: Current slide" + nl +
                "Fonts: None" + nl +
                "Font sizes: None" + nl +
                "Text colors: None" + nl +
                "Fill colors: None" + nl +
                "Line colors: None",
                report);
        }
    }
}
