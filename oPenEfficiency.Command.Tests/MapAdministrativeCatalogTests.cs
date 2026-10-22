using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Command.Tests
{
    [TestClass]
    [DoNotParallelize]
    public sealed class MapAdministrativeCatalogTests
    {
        [DataContract]
        private sealed class SourceMap
        {
            [DataMember] public List<SourceRegion> Regions { get; set; }
        }

        // Deliberately exclude polygon coordinates: compare selection metadata independently.
        [DataContract]
        private sealed class SourceRegion
        {
            [DataMember] public string Id { get; set; }
            [DataMember] public string Name { get; set; }
            [DataMember] public string EnglishName { get; set; }
            [DataMember] public string Group { get; set; }
            [DataMember] public string ParentGroup { get; set; }
            [DataMember] public string EntityType { get; set; }
        }

        private static List<SourceRegion> ReadSource(string suffix)
        {
            Assembly assembly = typeof(MapAssetService).Assembly;
            string resource = assembly.GetManifestResourceNames().Single(name => name.EndsWith(suffix, StringComparison.Ordinal));
            using (Stream stream = assembly.GetManifestResourceStream(resource))
            using (var gzip = new GZipStream(stream, CompressionMode.Decompress))
                return ((SourceMap)new DataContractJsonSerializer(typeof(SourceMap)).ReadObject(gzip)).Regions;
        }

        [TestMethod]
        public void FirstLevelNamesAndIdsMatchBundledGeometryExactly()
        {
            foreach (var country in ReadSource("states_provinces_admin1_naturalearth_5_1_1.json.gz").GroupBy(region => region.Group))
            {
                var actual = MapAssetService.GetCountryAdministrativeRegionCatalog(country.Key).ToDictionary(region => region.Id);
                Assert.AreEqual(country.Count(), actual.Count, country.Key);
                foreach (SourceRegion expected in country)
                {
                    Assert.IsTrue(actual.ContainsKey(expected.Id), expected.Id);
                    Assert.AreEqual(expected.EnglishName ?? expected.Name ?? expected.Id, actual[expected.Id].DisplayName, expected.Id);
                    Assert.AreEqual(expected.Name, actual[expected.Id].LocalName, expected.Id);
                    Assert.AreEqual(expected.EntityType, actual[expected.Id].EntityType, expected.Id);
                }
            }
        }

        [DataTestMethod]
        [DataRow("ADM2", "administrative_regions_adm2_geoboundaries_current.json.gz")]
        [DataRow("ADM3", "administrative_regions_adm3_geoboundaries_current.json.gz")]
        public void AvailabilityCountsAndParentsMatchEveryBundledRegion(string level, string suffix)
        {
            foreach (var country in ReadSource(suffix).GroupBy(region => region.Group))
            {
                Assert.AreEqual(country.Count(), Count(level, country.Key, null), country.Key);
                foreach (var parent in country.Where(region => !string.IsNullOrWhiteSpace(region.ParentGroup)).GroupBy(region => region.ParentGroup))
                    Assert.AreEqual(parent.Count(), Count(level, country.Key, parent.Key), parent.Key);
                var parentIds = new HashSet<string>(country.Select(region => region.ParentGroup), StringComparer.OrdinalIgnoreCase);
                var expected = MapAssetService.GetCountryAdministrativeRegionCatalog(country.Key)
                    .Where(region => parentIds.Contains(region.Id)).Select(region => region.Id).ToArray();
                CollectionAssert.AreEqual(expected, MapAssetService.GetLowerLevelParentRegionCatalog(country.Key, level).Select(region => region.Id).ToArray(), country.Key);
            }
            Assert.AreEqual(0, Count(level, "NOT-A-COUNTRY", null));
            Assert.AreEqual(0, Count(level, null, null));
        }

        private static int Count(string level, string country, string parent) => level == "ADM3"
            ? MapAssetService.GetThirdLevelAdministrativeRegionCount(country, parent)
            : MapAssetService.GetSecondLevelAdministrativeRegionCount(country, parent);

        [TestMethod]
        public void AvailabilityQueriesNeverPopulateThePolygonCache()
        {
            var cache = (IDictionary)typeof(MapAssetService).GetField("AssetCache", BindingFlags.NonPublic | BindingFlags.Static).GetValue(null);
            cache.Clear();
            MapAssetService.GetSecondLevelAdministrativeRegionCount("AFG");
            MapAssetService.GetThirdLevelAdministrativeRegionCount("DEU");
            MapAssetService.GetLowerLevelParentRegionCatalog("DEU", "ADM3");
            Assert.AreEqual(0, cache.Count, "UI availability loaded full geometry into the production cache.");
        }
    }
}
