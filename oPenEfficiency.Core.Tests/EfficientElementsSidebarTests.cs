using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class EfficientElementsSidebarTests
    {
        private static string Root()
        {
            var directory = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (directory != null && !File.Exists(Path.Combine(directory.FullName, "docs/product/feature-catalog.md")))
                directory = directory.Parent;
            if (directory == null) throw new InvalidOperationException("Repository feature catalog not found.");
            return directory.FullName;
        }

        private static List<FeatureCatalogRecord> Catalog()
        {
            var root = Root();
            var records = new Dictionary<string, FeatureCatalogRecord>(StringComparer.Ordinal);
            var pattern = new Regex(@"^\|\s*\d+\s*\|\s*`([^`]+)`\s*\|\s*([^|]+)\|\s*([^|]+)\|\s*([^|]+)\|\s*([^|]+)\|\s*([^|]+)\|\s*([^|]+)\|");
            foreach (var line in File.ReadLines(Path.Combine(root, "docs/product/feature-catalog.md")))
            {
                var match = pattern.Match(line);
                if (!match.Success) continue;
                var id = match.Groups[1].Value;
                records.Add(id, new FeatureCatalogRecord
                {
                    Id = id, Name = match.Groups[2].Value.Trim(), UserVisible = true,
                    Category = (FeatureCategory)Enum.Parse(typeof(FeatureCategory), match.Groups[3].Value.Trim()),
                    Maturity = (FeatureMaturity)Enum.Parse(typeof(FeatureMaturity), match.Groups[6].Value.Trim()),
                    DefaultPlacement = (FeatureDefaultPlacement)Enum.Parse(typeof(FeatureDefaultPlacement), match.Groups[7].Value.Trim())
                });
            }
            var ordered = new List<FeatureCatalogRecord>();
            var seed = File.ReadAllText(Path.Combine(root, "oPenEfficiency/UI/FeatureLibrary.cs"));
            foreach (Match match in Regex.Matches(seed, @"new SidebarFeature\s*\{\s*Id\s*=\s*""([^""]+)"""))
                if (records.TryGetValue(match.Groups[1].Value, out var record) && !ordered.Contains(record)) ordered.Add(record);
            ordered.AddRange(records.Values.Where(record => !ordered.Contains(record)).OrderBy(record => record.Name, StringComparer.Ordinal));
            return ordered;
        }

        private static string[] Ids(SidebarConfig config) => config.Sections.SelectMany(section => section.Features).Select(feature => feature.Id).ToArray();

        [TestMethod]
        public void EveryRequiredEntryPointExistsAndIsStableInTheRealCatalog()
        {
            var records = Catalog().ToDictionary(record => record.Id);
            Assert.AreEqual(180, records.Count);
            foreach (var id in SidebarPresetFactory.EfficientElementsRequiredIds)
            {
                Assert.IsTrue(records.ContainsKey(id), "Required command is absent: " + id);
                Assert.AreEqual(FeatureMaturity.Stable, records[id].Maturity, id);
            }
        }

        [TestMethod]
        public void StandardShowsAllExistingCounterpartsWithoutLosingExistingDefaults()
        {
            var catalog = Catalog();
            var shown = Ids(SidebarPresetFactory.Create(SidebarPresetFactory.ConsultingStandardName, catalog));
            foreach (var id in SidebarPresetFactory.EfficientElementsRequiredIds) CollectionAssert.Contains(shown, id);
            foreach (var record in catalog.Where(record => record.Maturity == FeatureMaturity.Stable &&
                record.DefaultPlacement == FeatureDefaultPlacement.ProfessionalDefault &&
                record.Category != FeatureCategory.SettingsAndHelp)) CollectionAssert.Contains(shown, record.Id);
            Assert.AreEqual(shown.Length, shown.Distinct(StringComparer.Ordinal).Count());
            CollectionAssert.Contains(shown, "BtnSelectSameType");
            Assert.IsTrue(shown.Length < Ids(SidebarPresetFactory.Create(SidebarPresetFactory.AllFeaturesName, catalog)).Length);
        }

        [TestMethod]
        public void HiddenObsoleteAndPreviewRequiredEntriesAreNotPromoted()
        {
            var catalog = Catalog();
            catalog.Single(record => record.Id == "BtnSlidePaste").UserVisible = false;
            catalog.Single(record => record.Id == "BtnSpellCheck").IsObsolete = true;
            catalog.Single(record => record.Id == "BtnStretchLeft").Maturity = FeatureMaturity.Preview;
            var shown = Ids(SidebarPresetFactory.Create(SidebarPresetFactory.ConsultingStandardName, catalog));
            foreach (var id in new[] { "BtnSlidePaste", "BtnSpellCheck", "BtnStretchLeft" }) CollectionAssert.DoesNotContain(shown, id);
        }

        [TestMethod]
        public void ExistingPersonalLayoutAndItsContentArePreserved()
        {
            var config = new SidebarConfig { Sections = new List<SidebarSection> { new SidebarSection
                { Name = "Personal", Features = new List<SidebarFeature> { new SidebarFeature { Id = "BtnSettings", Name = "My Settings" } } } } };
            var resolved = SidebarPresetFactory.ResolveInitialLayout(config, Catalog());
            Assert.AreSame(config, resolved);
            Assert.AreEqual("Personal", resolved.Sections.Single().Name);
            Assert.AreEqual("My Settings", resolved.Sections.Single().Features.Single().Name);
            CollectionAssert.AreEqual(new[] { "BtnSettings" }, Ids(resolved));
        }

        [TestMethod]
        public void StandardUsesApprovedArrangeOrderAndPersistentFinalActions()
        {
            var config = SidebarPresetFactory.Create(SidebarPresetFactory.ConsultingStandardName, Catalog());
            Assert.AreEqual("BtnSearchBar", Ids(config).First());
            CollectionAssert.AreEqual(new[] { "BtnCopyXY", "BtnPasteXY", "BtnPositionPainter", "BtnSwapPositions",
                "BtnBringToFront", "BtnSendToBack", "BtnSelectSameType" },
                config.Sections.Single(section => section.Name == "ARRANGE & ORGANIZE").Features.Take(7).Select(feature => feature.Id).ToArray());
            CollectionAssert.AreEqual(new[] { "BtnSettings", "BtnExploreFeatures" }, config.Sections.Last().Features.Select(feature => feature.Id).ToArray());
            CollectionAssert.AreEquivalent(SidebarPresetFactory.MinimalRequiredIds.ToArray(), Ids(SidebarPresetFactory.Create(SidebarPresetFactory.MinimalName, Catalog())));
        }

        [TestMethod]
        public void ExportedInventoryMatchesTheRuntimePreset()
        {
            Inventory inventory;
            using (var stream = File.OpenRead(Path.Combine(Root(), "oPenEfficiency.Core.Tests/TestData/sidebar-default-inventory.json")))
                inventory = (Inventory)new DataContractJsonSerializer(typeof(Inventory)).ReadObject(stream);
            var expected = inventory.Sections.SelectMany(section => section.Features).Select(feature => feature.Id).ToArray();
            CollectionAssert.AreEqual(expected, Ids(SidebarPresetFactory.Create(SidebarPresetFactory.ConsultingStandardName, Catalog())));
        }

        [TestMethod]
        public void PublicDefaultInventoryCoversEveryRequiredExistingCommand()
        {
            // The private competitor-control matrix remains private. Pin the same public
            // required product IDs against an independent frozen default-layout snapshot.
            Inventory inventory;
            using (var stream = File.OpenRead(Path.Combine(Root(), "oPenEfficiency.Core.Tests/TestData/sidebar-default-inventory.json")))
                inventory = (Inventory)new DataContractJsonSerializer(typeof(Inventory)).ReadObject(stream);
            var recorded = inventory.Sections.SelectMany(section => section.Features).Select(feature => feature.Id).ToArray();
            var shown = Ids(SidebarPresetFactory.Create(SidebarPresetFactory.ConsultingStandardName, Catalog()));
            foreach (var id in SidebarPresetFactory.EfficientElementsRequiredIds)
            {
                CollectionAssert.Contains(recorded, id, "Required command missing from frozen public contract: " + id);
                CollectionAssert.Contains(shown, id, "Required command missing from runtime preset: " + id);
            }
        }

        [DataContract] private sealed class Inventory { [DataMember(Name = "defaultSections")] public List<Section> Sections { get; set; } }
        [DataContract] private sealed class Section { [DataMember(Name = "features")] public List<Feature> Features { get; set; } }
        [DataContract] private sealed class Feature { [DataMember(Name = "id")] public string Id { get; set; } }
        [DataContract] private sealed class Coverage { [DataMember(Name = "entries")] public List<CoverageEntry> Entries { get; set; } }
        [DataContract] private sealed class CoverageEntry
        {
            [DataMember(Name = "controlId")] public string ControlId { get; set; }
            [DataMember(Name = "featureIds")] public List<string> FeatureIds { get; set; }
            [DataMember(Name = "disposition")] public string Disposition { get; set; }
        }
    }
}
