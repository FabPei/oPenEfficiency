using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 external-data-wave coverage for the native linked-table tag
    /// (QA9-BtnRefreshLinkedTable-002): the oPE_TableLink codec with its optional flags
    /// token, legacy flag-less format and pipe-rejoining sheet names, used by TableLink.
    /// </summary>
    [TestClass]
    public sealed class TableLinkStatePolicyTests
    {
        [TestMethod]
        public void BuildFlagsEncodesTheThreeStates()
        {
            Assert.AreEqual("V", TableLinkStatePolicy.BuildFlags(false, false));
            Assert.AreEqual("F", TableLinkStatePolicy.BuildFlags(true, false));
            Assert.AreEqual("R", TableLinkStatePolicy.BuildFlags(false, true));
            Assert.AreEqual("FR", TableLinkStatePolicy.BuildFlags(true, true));
        }

        [TestMethod]
        public void FlagsTokenRecognitionSeparatesNewTagsFromPaths()
        {
            Assert.IsTrue(TableLinkStatePolicy.IsFlagsToken("V"));
            Assert.IsTrue(TableLinkStatePolicy.IsFlagsToken("FR"));
            Assert.IsFalse(TableLinkStatePolicy.IsFlagsToken("FRVX")); // too long
            Assert.IsFalse(TableLinkStatePolicy.IsFlagsToken("FX"));   // invalid letter
            Assert.IsFalse(TableLinkStatePolicy.IsFlagsToken(""));
            Assert.IsFalse(TableLinkStatePolicy.IsFlagsToken(null));
            // A workbook path always contains ':' or a slash, so it never parses as flags.
            Assert.IsFalse(TableLinkStatePolicy.IsFlagsToken("C:"));
        }

        [TestMethod]
        public void NewFormatRoundTripsAndRejoinsPipedSheetNames()
        {
            Assert.IsTrue(TableLinkStatePolicy.TryParseTag(
                "oPE_TableLink|FR|C:\\Reports\\Q3.xlsx|Sheet|With|Pipes|A1:C4",
                out string path, out string sheet, out string range, out bool formatted, out bool autoResize));
            Assert.AreEqual("C:\\Reports\\Q3.xlsx", path);
            Assert.AreEqual("Sheet|With|Pipes", sheet); // interior pipes belong to the sheet
            Assert.AreEqual("A1:C4", range);
            Assert.IsTrue(formatted);
            Assert.IsTrue(autoResize);
        }

        [TestMethod]
        public void LegacyFlaglessFormatStillParses()
        {
            Assert.IsTrue(TableLinkStatePolicy.TryParseTag(
                "oPE_TableLink|C:\\Data.xlsx|Tabelle1|B2:D9",
                out string path, out string sheet, out string range, out bool formatted, out bool autoResize));
            Assert.AreEqual("C:\\Data.xlsx", path);
            Assert.AreEqual("Tabelle1", sheet);
            Assert.AreEqual("B2:D9", range);
            Assert.IsFalse(formatted);
            Assert.IsFalse(autoResize);
        }

        [TestMethod]
        public void MalformedTagsAreRejected()
        {
            Assert.IsFalse(TableLinkStatePolicy.TryParseTag(null, out _, out _, out _, out _, out _));
            Assert.IsFalse(TableLinkStatePolicy.TryParseTag("", out _, out _, out _, out _, out _));
            Assert.IsFalse(TableLinkStatePolicy.TryParseTag("other|F|path|sheet|A1", out _, out _, out _, out _, out _));
            Assert.IsFalse(TableLinkStatePolicy.TryParseTag("oPE_TableLink|path|sheet", out _, out _, out _, out _, out _));
        }

        [TestMethod]
        public void BuildAndParseRoundTrip()
        {
            string tag = $"oPE_TableLink|{TableLinkStatePolicy.BuildFlags(true, false)}|C:\\W.xlsx|My|Sheet|A1:B2";
            Assert.IsTrue(TableLinkStatePolicy.TryParseTag(tag,
                out string path, out string sheet, out string range, out bool formatted, out bool autoResize));
            Assert.AreEqual("C:\\W.xlsx", path);
            Assert.AreEqual("My|Sheet", sheet);
            Assert.AreEqual("A1:B2", range);
            Assert.IsTrue(formatted);
            Assert.IsFalse(autoResize);
        }
    }
}
