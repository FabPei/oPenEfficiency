using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 external-data-wave coverage for the Excel cell value link
    /// (QA9-BtnInsertExcelValue-002): the newline-delimited key=value AlternativeText codec
    /// with invariant numbers and documented fallbacks, used by ExcelValueLink.
    /// </summary>
    [TestClass]
    public sealed class ExcelValueLinkStatePolicyTests
    {
        [TestMethod]
        public void BuildAndParseRoundTripAllFields()
        {
            string tag = ExcelValueLinkStatePolicy.BuildTag(
                "C:\\W.xlsx", "Sheet1", "$B$3",
                decimals: 2, thousands: true, percent: false,
                raw: 1234.5, id: "abc", start: 7, length: 9, text: "1.234,50 €");
            Assert.IsTrue(ExcelValueLinkStatePolicy.TryParseTag(tag,
                out string path, out string sheet, out string cell,
                out int decimals, out bool thousands, out bool percent,
                out double? raw, out string id, out int start, out int length, out string text));
            Assert.AreEqual("C:\\W.xlsx", path);
            Assert.AreEqual("Sheet1", sheet);
            Assert.AreEqual("$B$3", cell);
            Assert.AreEqual(2, decimals);
            Assert.IsTrue(thousands);
            Assert.IsFalse(percent);
            Assert.AreEqual(1234.5, raw.Value, 0.0001);
            Assert.AreEqual("abc", id);
            Assert.AreEqual(7, start);
            Assert.AreEqual(9, length);
            Assert.AreEqual("1.234,50 €", text);
        }

        [TestMethod]
        public void RawNumberIsWrittenInvariant()
        {
            System.Globalization.CultureInfo original = System.Globalization.CultureInfo.CurrentCulture;
            try
            {
                System.Globalization.CultureInfo.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo("de-DE");
                string tag = ExcelValueLinkStatePolicy.BuildTag("p", "s", "c", -1, false, false, 1.5, null, -1, -1, null);
                StringAssert.Contains(tag, "raw=1.5"); // never a locale decimal comma in a tag
            }
            finally
            {
                System.Globalization.CultureInfo.CurrentCulture = original;
            }
        }

        [TestMethod]
        public void UnparsableFieldsKeepTheirDocumentedDefaults()
        {
            Assert.IsTrue(ExcelValueLinkStatePolicy.TryParseTag(
                "oPE_ExcelValue\npath=p\nthou=yes\nraw=not-a-number\nstart=xx",
                out _, out _, out _,
                out int decimals, out bool thousands, out bool percent,
                out double? raw, out _, out int start, out _, out _));
            Assert.AreEqual(-1, decimals);
            Assert.IsFalse(thousands); // only exact "1" sets the flag
            Assert.IsFalse(percent);
            Assert.IsNull(raw);
            Assert.AreEqual(-1, start);
        }

        [TestMethod]
        public void ValuesMayContainEqualsSignsAfterTheFirst()
        {
            Assert.IsTrue(ExcelValueLinkStatePolicy.TryParseTag(
                "oPE_ExcelValue\ntext=a=b=c",
                out _, out _, out _, out _, out _, out _, out _, out _, out _, out _, out string text));
            Assert.AreEqual("a=b=c", text); // the value keeps everything after the first '='
        }

        [TestMethod]
        public void HandlesCarriageReturnsAndRejectsNonTags()
        {
            Assert.IsTrue(ExcelValueLinkStatePolicy.TryParseTag(
                "oPE_ExcelValue\r\npath=p\r\ncell=A1",
                out string path, out _, out string cell,
                out _, out _, out _, out _, out _, out _, out _, out _));
            Assert.AreEqual("p", path);
            Assert.AreEqual("A1", cell);

            Assert.IsFalse(ExcelValueLinkStatePolicy.TryParseTag(null, out _, out _, out _, out _, out _, out _, out _, out _, out _, out _, out _));
            Assert.IsFalse(ExcelValueLinkStatePolicy.TryParseTag("something else", out _, out _, out _, out _, out _, out _, out _, out _, out _, out _, out _));
        }

        [TestMethod]
        public void MissingOptionalKeysLeaveNullsForReferenceFields()
        {
            Assert.IsTrue(ExcelValueLinkStatePolicy.TryParseTag(
                "oPE_ExcelValue\ncell=A1",
                out string path, out string sheet, out string cell,
                out _, out _, out _, out _, out string id, out _, out _, out string text));
            Assert.IsNull(path);
            Assert.IsNull(sheet);
            Assert.AreEqual("A1", cell);
            Assert.IsNull(id);
            Assert.IsNull(text);
        }
    }
}
