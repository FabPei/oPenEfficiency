using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public class IllustrationSourceCatalogTests
    {
        [TestMethod]
        public void EveryEvaluatedSourceIsListed()
        {
            Assert.AreEqual(21, IllustrationSourceCatalog.All.Count, "The evaluated source list is the user-facing contract; update it deliberately.");
        }

        [TestMethod]
        public void EverySourceCarriesNameLinkAndLicenseFacts()
        {
            foreach (var source in IllustrationSourceCatalog.All)
            {
                Assert.IsFalse(string.IsNullOrWhiteSpace(source.Id), "source without id");
                Assert.IsFalse(string.IsNullOrWhiteSpace(source.Name), "source without name: " + source.Id);
                Assert.IsFalse(string.IsNullOrWhiteSpace(source.LicenseName), "source without license name: " + source.Id);
                Assert.IsFalse(string.IsNullOrWhiteSpace(source.StatusLabel), "source without status label: " + source.Id);
                Assert.IsTrue(
                    source.HomePageUrl != null && source.HomePageUrl.StartsWith("https://", StringComparison.Ordinal),
                    "source without https home page: " + source.Id);
                Assert.IsTrue(
                    source.LicenseUrl != null && source.LicenseUrl.StartsWith("https://", StringComparison.Ordinal),
                    "source without https license page: " + source.Id);
            }
        }

        [TestMethod]
        public void SourceIdsAreUnique()
        {
            var duplicates = IllustrationSourceCatalog.All
                .GroupBy(source => source.Id, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key)
                .ToList();

            Assert.AreEqual(0, duplicates.Count, "duplicate source ids: " + string.Join(", ", duplicates));
        }

        [TestMethod]
        public void BundledSourcesMatchTheShippedCollections()
        {
            var bundled = IllustrationSourceCatalog.Bundled;

            Assert.AreEqual(3, bundled.Count, "only the CC0 collections may ship inside the add-in");
            foreach (var source in bundled)
            {
                Assert.AreEqual(IllustrationSourceStatus.Bundled, source.Status);
                Assert.IsFalse(source.AttributionRequired, "bundled source with attribution duty: " + source.Id);
                Assert.IsTrue(source.BundledFigureCount > 0, "bundled source without figures: " + source.Id);
            }

            Assert.AreEqual(421, IllustrationSourceCatalog.BundledFigureTotal, "bundled figure total must match the shipped assets");
            CollectionAssert.AreEquivalent(
                new[] { "open-peeps", "open-doodles", "kitbitz" },
                bundled.Select(source => source.Id).ToArray());
        }

        [TestMethod]
        public void NonBundledSourcesExplainHowToGetThem()
        {
            foreach (var source in IllustrationSourceCatalog.All.Where(item => !item.IsBundled))
            {
                Assert.IsFalse(string.IsNullOrWhiteSpace(source.AcquisitionNote), "source without acquisition note: " + source.Id);
                Assert.IsFalse(string.IsNullOrWhiteSpace(source.TechnicalNote), "source without technical note: " + source.Id);
            }
        }

        [TestMethod]
        public void RestrictedSourcesAreMarkedAsSuch()
        {
            Assert.AreEqual(
                IllustrationSourceStatus.RedistributionRestricted,
                IllustrationSourceCatalog.Find("storyset").Status);
            Assert.AreEqual(
                IllustrationSourceStatus.RedistributionRestricted,
                IllustrationSourceCatalog.Find("undraw").Status);
            Assert.AreEqual(
                IllustrationSourceStatus.BundlingPermitted,
                IllustrationSourceCatalog.Find("humaaans").Status);
            Assert.IsNull(IllustrationSourceCatalog.Find("does-not-exist"));
        }

        [TestMethod]
        public void ResponsibilityNoticeNamesTheUserDuty()
        {
            string notice = IllustrationSourceCatalog.ResponsibilityNotice;

            Assert.IsTrue(notice.IndexOf("responsible", StringComparison.OrdinalIgnoreCase) >= 0, notice);
            Assert.IsTrue(notice.IndexOf("license", StringComparison.OrdinalIgnoreCase) >= 0, notice);
            Assert.IsTrue(notice.IndexOf("no legal advice", StringComparison.OrdinalIgnoreCase) >= 0, notice);
        }

        [TestMethod]
        public void AttributionDutyIsVisibleInTheLicenseSummary()
        {
            StringAssert.Contains(IllustrationSourceCatalog.Find("ira-design").LicenseSummary, "attribution required");
            StringAssert.Contains(IllustrationSourceCatalog.Find("humaaans").LicenseSummary, "no attribution");
        }
    }
}
