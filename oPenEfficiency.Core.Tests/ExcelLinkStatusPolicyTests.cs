using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 3 coverage for the Excel Link Manager display rules (QA9-BtnExcelLinkManager-001):
    /// the status resolution from the source's existence and the last-refresh tag formatting,
    /// used by ExcelLinkManagerWindow.
    /// </summary>
    [TestClass]
    public sealed class ExcelLinkStatusPolicyTests
    {
        [TestMethod]
        public void StatusNeedsBothASourcePathAndAnExistingFile()
        {
            Assert.AreEqual(ExcelLinkStatusPolicy.LinkedStatus, ExcelLinkStatusPolicy.ResolveStatus(true, true));
            Assert.AreEqual(ExcelLinkStatusPolicy.MissingStatus, ExcelLinkStatusPolicy.ResolveStatus(true, false),
                "A source path whose file is gone reports Missing.");
            Assert.AreEqual(ExcelLinkStatusPolicy.MissingStatus, ExcelLinkStatusPolicy.ResolveStatus(false, true),
                "A shape without a source path can never be Linked.");
            Assert.AreEqual(ExcelLinkStatusPolicy.MissingStatus, ExcelLinkStatusPolicy.ResolveStatus(false, false));
        }

        [TestMethod]
        public void LastRefreshRendersThroughTheCurrentCulture()
        {
            var stamp = new DateTime(2026, 10, 17, 14, 30, 0);
            string raw = stamp.ToString(CultureInfo.CurrentCulture);

            Assert.AreEqual(stamp.ToString("g", CultureInfo.CurrentCulture),
                ExcelLinkStatusPolicy.FormatLastRefresh(raw),
                "A parseable tag renders with the culture's short date/time format.");
        }

        [TestMethod]
        public void BlankOrUnparseableTagsRenderAsEmDash()
        {
            Assert.AreEqual(ExcelLinkStatusPolicy.NoRefreshValue, ExcelLinkStatusPolicy.FormatLastRefresh(null));
            Assert.AreEqual(ExcelLinkStatusPolicy.NoRefreshValue, ExcelLinkStatusPolicy.FormatLastRefresh(""));
            Assert.AreEqual(ExcelLinkStatusPolicy.NoRefreshValue, ExcelLinkStatusPolicy.FormatLastRefresh("not-a-date"));
        }
    }
}
