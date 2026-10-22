using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services.Attributes;
using TestAssert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;

namespace oPenEfficiency.Catalog.Tests
{
    /// <summary>
    /// QA-S7: the acceptance inventory in docs/testing/feature-acceptance must stay tied to the
    /// resolved production catalog. Catalog.Tests owns the catalog truth, so the coverage ledger and
    /// the per-feature catalog manifests are asserted here instead of relying on a PowerShell-side
    /// feature list.
    /// </summary>
    [TestClass]
    public class AcceptanceCoverageTests
    {
        private const string CoverageRelativePath = "docs/testing/feature-acceptance/catalog-coverage.json";
        private const string CatalogManifestRelativeDirectory = "docs/testing/feature-acceptance/catalog";
        private const string ExclusionRelativePath = "docs/testing/feature-acceptance/catalog-exclusions.json";
        private const string CanonicalCatalogRelativePath = "docs/product/feature-catalog.md";

        [TestMethod]
        public void AcceptanceCoverageCoversEveryVisibleFeature()
        {
            string root = FindRepositoryRoot();
            var visibleIds = ReadVisibleCatalogIds(Path.Combine(root, CanonicalCatalogRelativePath));
            TestAssert.IsTrue(visibleIds.Count > 0, "The canonical feature catalog lists no visible Feature IDs.");

            var ledger = ReadLedger(Path.Combine(root, CoverageRelativePath));
            var exclusions = ReadExclusions(Path.Combine(root, ExclusionRelativePath));
            var excludedIds = new HashSet<string>(
                exclusions.Exclusions.Where(e => e != null && !string.IsNullOrWhiteSpace(e.FeatureId)).Select(e => e.FeatureId),
                StringComparer.Ordinal);

            TestAssert.AreEqual(
                visibleIds.Count,
                ledger.VisibleFeatureTotal,
                "The coverage ledger does not agree with the canonical feature catalog. Run eng/Generate-FeatureAcceptanceCoverage.ps1.");
            TestAssert.AreEqual(
                0,
                ledger.UncoveredFeatureIds == null ? 0 : ledger.UncoveredFeatureIds.Length,
                "The coverage ledger reports uncovered visible features: " + string.Join(", ", ledger.UncoveredFeatureIds ?? new string[0]));

            var ledgerIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var feature in ledger.Features)
            {
                TestAssert.IsTrue(visibleIds.Contains(feature.FeatureId),
                    "The coverage ledger contains the stale Feature ID '" + feature.FeatureId + "'.");
                TestAssert.IsTrue(ledgerIds.Add(feature.FeatureId),
                    "The coverage ledger lists the Feature ID '" + feature.FeatureId + "' more than once.");
                TestAssert.IsTrue(feature.Covered || feature.Excluded,
                    "Feature ID '" + feature.FeatureId + "' is neither covered nor explicitly excluded.");
                if (!feature.Covered)
                {
                    continue;
                }

                TestAssert.IsTrue(
                    feature.ScenarioIds != null && feature.ScenarioIds.Length > 0,
                    "Covered Feature ID '" + feature.FeatureId + "' names no acceptance scenario.");

                string manifestPath = Path.Combine(root, CatalogManifestRelativeDirectory, feature.FeatureId + ".json");
                TestAssert.IsTrue(File.Exists(manifestPath),
                    "Covered Feature ID '" + feature.FeatureId + "' has no catalog manifest.");
                var manifest = ReadCatalogManifest(manifestPath);
                TestAssert.AreEqual(feature.FeatureId, manifest.FeatureId,
                    "catalog/" + feature.FeatureId + ".json declares a different Feature ID.");
                TestAssert.IsTrue(manifest.Scenarios != null && manifest.Scenarios.Length > 0,
                    "catalog/" + feature.FeatureId + ".json declares no acceptance scenario.");
                foreach (string scenarioId in feature.ScenarioIds)
                {
                    TestAssert.IsTrue(
                        manifest.Scenarios.Any(s => string.Equals(s.Id, scenarioId, StringComparison.Ordinal)),
                        "catalog/" + feature.FeatureId + ".json does not declare the ledger scenario '" + scenarioId + "'.");
                }
            }

            foreach (string visibleId in visibleIds)
            {
                TestAssert.IsTrue(
                    ledgerIds.Contains(visibleId),
                    "Visible Feature ID '" + visibleId + "' is missing from the coverage ledger.");
                TestAssert.IsTrue(
                    ledger.Features.First(f => f.FeatureId == visibleId).Covered || excludedIds.Contains(visibleId),
                    "Visible Feature ID '" + visibleId + "' has no acceptance scenario and no recorded exclusion.");
            }

            TestAssert.AreEqual(ledger.CoveredFeatureTotal + ledger.ExcludedFeatureTotal, ledger.VisibleFeatureTotal,
                "The coverage ledger totals do not add up to the visible feature count.");
        }

        private static HashSet<string> ReadVisibleCatalogIds(string path)
        {
            TestAssert.IsTrue(File.Exists(path), "The canonical feature catalog is missing: " + path);
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (string line in File.ReadAllLines(path))
            {
                var match = System.Text.RegularExpressions.Regex.Match(line, @"^\|\s*\d+\s*\|\s*`(?<id>[^`]+)`\s*\|");
                if (match.Success)
                {
                    ids.Add(match.Groups["id"].Value);
                }
            }

            return ids;
        }

        private static CatalogCoverageDocument ReadLedger(string path)
        {
            TestAssert.IsTrue(File.Exists(path),
                "The catalog coverage ledger is missing. Run eng/Generate-FeatureAcceptanceCoverage.ps1.");
            return Deserialize<CatalogCoverageDocument>(path);
        }

        private static CatalogExclusionDocument ReadExclusions(string path)
        {
            TestAssert.IsTrue(File.Exists(path), "The catalog exclusion document is missing: " + path);
            return Deserialize<CatalogExclusionDocument>(path);
        }

        private static CatalogManifestDocument ReadCatalogManifest(string path) =>
            Deserialize<CatalogManifestDocument>(path);

        private static T Deserialize<T>(string path)
        {
            var serializer = new DataContractJsonSerializer(typeof(T));
            using (var stream = File.OpenRead(path))
            {
                return (T)serializer.ReadObject(stream);
            }
        }

        private static string FindRepositoryRoot()
        {
            var current = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (current != null)
            {
                if (File.Exists(Path.Combine(current.FullName, "oPenEfficiency.slnx")))
                {
                    return current.FullName;
                }

                current = current.Parent;
            }

            throw new DirectoryNotFoundException("Could not locate the repository root.");
        }

        [DataContract]
        private sealed class CatalogCoverageDocument
        {
            [DataMember(Name = "source")] public string Source { get; set; }
            [DataMember(Name = "visibleFeatureTotal")] public int VisibleFeatureTotal { get; set; }
            [DataMember(Name = "coveredFeatureTotal")] public int CoveredFeatureTotal { get; set; }
            [DataMember(Name = "excludedFeatureTotal")] public int ExcludedFeatureTotal { get; set; }
            [DataMember(Name = "scenarioTotal")] public int ScenarioTotal { get; set; }
            [DataMember(Name = "uncoveredFeatureIds")] public string[] UncoveredFeatureIds { get; set; }
            [DataMember(Name = "features")] public CatalogCoverageFeature[] Features { get; set; }
        }

        [DataContract]
        private sealed class CatalogCoverageFeature
        {
            [DataMember(Name = "featureId")] public string FeatureId { get; set; }
            [DataMember(Name = "name")] public string Name { get; set; }
            [DataMember(Name = "category")] public string Category { get; set; }
            [DataMember(Name = "covered")] public bool Covered { get; set; }
            [DataMember(Name = "excluded")] public bool Excluded { get; set; }
            [DataMember(Name = "scenarioIds")] public string[] ScenarioIds { get; set; }
        }

        [DataContract]
        private sealed class CatalogManifestDocument
        {
            [DataMember(Name = "featureId")] public string FeatureId { get; set; }
            [DataMember(Name = "title")] public string Title { get; set; }
            [DataMember(Name = "sourcePaths")] public string[] SourcePaths { get; set; }
            [DataMember(Name = "scenarios")] public CatalogManifestScenario[] Scenarios { get; set; }
        }

        [DataContract]
        private sealed class CatalogManifestScenario
        {
            [DataMember(Name = "id")] public string Id { get; set; }
            [DataMember(Name = "tier")] public string Tier { get; set; }
            [DataMember(Name = "automationStatus")] public string AutomationStatus { get; set; }
            [DataMember(Name = "changedPathPatterns")] public string[] ChangedPathPatterns { get; set; }
        }

        [DataContract]
        private sealed class CatalogExclusionDocument
        {
            [DataMember(Name = "exclusions")] public CatalogExclusionEntry[] Exclusions { get; set; }
        }

        [DataContract]
        private sealed class CatalogExclusionEntry
        {
            [DataMember(Name = "featureId")] public string FeatureId { get; set; }
            [DataMember(Name = "category")] public string Category { get; set; }
            [DataMember(Name = "reason")] public string Reason { get; set; }
        }
    }
}
