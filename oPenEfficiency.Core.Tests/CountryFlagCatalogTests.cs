using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public class CountryFlagCatalogTests
    {
        [TestMethod]
        public void CatalogContainsCompleteUniqueIsoSet()
        {
            Assert.AreEqual(249, CountryFlagCatalog.Entries.Count);
            Assert.AreEqual(249, CountryFlagCatalog.Entries.Select(entry => entry.Alpha2).Distinct(StringComparer.Ordinal).Count());
            Assert.AreEqual(249, CountryFlagCatalog.Entries.Select(entry => entry.Alpha3).Distinct(StringComparer.Ordinal).Count());
            Assert.IsTrue(CountryFlagCatalog.Entries.SequenceEqual(
                CountryFlagCatalog.Entries.OrderBy(entry => entry.Name, StringComparer.Ordinal)));
        }

        [TestMethod]
        public void SearchMatchesNamesCodesAliasesAndDiacritics()
        {
            Assert.AreEqual("DE", CountryFlagCatalog.Search("Germany").First().Alpha2);
            Assert.AreEqual("DE", CountryFlagCatalog.Search("de").First().Alpha2);
            Assert.AreEqual("DE", CountryFlagCatalog.Search("DEU").First().Alpha2);
            Assert.AreEqual("US", CountryFlagCatalog.Search("USA").First().Alpha2);
            Assert.AreEqual("CI", CountryFlagCatalog.Search("Ivory Coast").First().Alpha2);
            Assert.AreEqual("AX", CountryFlagCatalog.Search("Aland").First().Alpha2);
        }

        [TestMethod]
        public void RecentResolutionDropsUnknownAndDuplicateIds()
        {
            var recent = CountryFlagCatalog.ResolveRecent(new[] { "de", "missing", "DE", "jp" });
            CollectionAssert.AreEqual(new[] { "de", "jp" }, recent.Select(entry => entry.AssetId).ToArray());
        }
    }
}
