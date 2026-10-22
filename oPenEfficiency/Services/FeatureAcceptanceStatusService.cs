using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace oPenEfficiency.Services
{
    /// <summary>One acceptance scenario as the dev/QA window shows it.</summary>
    public sealed class FeatureAcceptanceScenarioStatus
    {
        public string ScenarioId { get; set; }
        public string FeatureId { get; set; }
        public string Tier { get; set; }
        public string AutomationStatus { get; set; }
        /// <summary>Green / Red / Missing / NotApplicable (visual and owner tiers have no auto evidence).</summary>
        public string EvidenceState { get; set; }
        public DateTime? EvidenceCompletedUtc { get; set; }
        public double EvidenceAgeHours { get; set; }
        /// <summary>failed / missing-evidence / stale / none - derived, documented in the QA-S6 section.</summary>
        public string Severity { get; set; }
        public string RunnerPath { get; set; }
        public string RunnerArguments { get; set; }
        public bool HasRunner
        {
            get { return !string.IsNullOrWhiteSpace(RunnerPath); }
        }
    }

    /// <summary>
    /// Reads the acceptance manifests, the local TestResults evidence and derives one status row
    /// per scenario. Dev-only by construction: it needs the repository checkout (manifests and
    /// TestResults live in the repo), and every command the window offers launches the existing
    /// eng/ scripts - this service never re-implements gate logic.
    /// </summary>
    public static class FeatureAcceptanceStatusService
    {
        private static string _cachedRepositoryRoot;
        private static bool _rootResolved;

        // The DTO fields below are populated exclusively by DataContractJsonSerializer, never by
        // code in this file, so CS0649 would fire on all of them; the serializer contract wins.
#pragma warning disable CS0649
        [DataContract]
        private sealed class ManifestDocument
        {
            [DataMember(Name = "featureId")]
            public string FeatureId;
            [DataMember(Name = "scenarios")]
            public ScenarioDocument[] Scenarios;
        }

        [DataContract]
        private sealed class ScenarioDocument
        {
            [DataMember(Name = "id")]
            public string Id;
            [DataMember(Name = "tier")]
            public string Tier;
            [DataMember(Name = "automationStatus")]
            public string AutomationStatus;
            [DataMember(Name = "runner")]
            public RunnerDocument Runner;
        }

        [DataContract]
        private sealed class RunnerDocument
        {
            [DataMember(Name = "kind")]
            public string Kind;
            [DataMember(Name = "path")]
            public string Path;
            [DataMember(Name = "arguments")]
            public string[] Arguments;
        }

        [DataContract]
        private sealed class EvidenceDocument
        {
            [DataMember(Name = "scenarioId")]
            public string ScenarioId;
            [DataMember(Name = "passed")]
            public bool Passed;
            [DataMember(Name = "completedUtc")]
            public string CompletedUtc;
        }
#pragma warning restore CS0649

        /// <summary>Dev gate: the QA window and its ribbon command exist only where the repository
        /// checkout (docs/testing/feature-acceptance) is reachable from the deployed assembly.</summary>
        public static bool IsAvailable
        {
            get { return TryFindRepositoryRoot() != null; }
        }

        public static string TryFindRepositoryRoot()
        {
            if (_rootResolved) return _cachedRepositoryRoot;
            try
            {
                var marker = Path.Combine("docs", "testing", "feature-acceptance");
                var directory = AppDomain.CurrentDomain.BaseDirectory;
                while (!string.IsNullOrWhiteSpace(directory))
                {
                    if (Directory.Exists(Path.Combine(directory, marker)))
                    {
                        _cachedRepositoryRoot = directory;
                        break;
                    }
                    directory = Path.GetDirectoryName(directory.TrimEnd(Path.DirectorySeparatorChar));
                }
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "FeatureAcceptanceStatusService.TryFindRepositoryRoot");
                _cachedRepositoryRoot = null;
            }
            _rootResolved = true;
            return _cachedRepositoryRoot;
        }

        private static T ReadJson<T>(string path)
        {
            // StreamReader strips the utf-8-sig BOM the acceptance manifests carry; the serializer
            // then reads the cleaned text (DataContractJsonSerializer would trip over the BOM).
            using (var reader = new StreamReader(path, Encoding.UTF8, true))
            using (var memory = new MemoryStream(Encoding.UTF8.GetBytes(reader.ReadToEnd())))
            {
                var serializer = new DataContractJsonSerializer(typeof(T));
                return (T)serializer.ReadObject(memory);
            }
        }

        public static List<FeatureAcceptanceScenarioStatus> LoadStatuses(out string repositoryRoot)
        {
            var rows = new List<FeatureAcceptanceScenarioStatus>();
            repositoryRoot = TryFindRepositoryRoot();
            if (repositoryRoot == null) return rows;

            var manifestDirectory = Path.Combine(repositoryRoot, "docs", "testing", "feature-acceptance");
            var evidenceDirectory = Path.Combine(repositoryRoot, "TestResults");
            var evidenceByScenario = new Dictionary<string, EvidenceDocument>(StringComparer.OrdinalIgnoreCase);

            try
            {
                if (Directory.Exists(evidenceDirectory))
                {
                    foreach (var evidencePath in Directory.GetFiles(evidenceDirectory, "feat-*.json"))
                    {
                        try
                        {
                            var evidence = ReadJson<EvidenceDocument>(evidencePath);
                            if (evidence != null && !string.IsNullOrWhiteSpace(evidence.ScenarioId))
                                evidenceByScenario[evidence.ScenarioId] = evidence;
                        }
                        catch (Exception ex)
                        {
                            Utils.ExceptionLogger.Log(ex, "FeatureAcceptanceStatusService.LoadStatuses.Evidence");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "FeatureAcceptanceStatusService.LoadStatuses.EvidenceDir");
            }

            var nowUtc = DateTime.UtcNow;
            foreach (var manifestPath in Directory.GetFiles(manifestDirectory, "feat-*.json").OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
            {
                ManifestDocument manifest;
                try
                {
                    manifest = ReadJson<ManifestDocument>(manifestPath);
                }
                catch (Exception ex)
                {
                    Utils.ExceptionLogger.Log(ex, "FeatureAcceptanceStatusService.LoadStatuses.Manifest");
                    continue;
                }
                if (manifest == null || manifest.Scenarios == null) continue;

                foreach (var scenario in manifest.Scenarios)
                {
                    if (scenario == null || string.IsNullOrWhiteSpace(scenario.Id)) continue;
                    var row = new FeatureAcceptanceScenarioStatus
                    {
                        ScenarioId = scenario.Id,
                        // featureId lives at manifest level in the acceptance files (scenarios only
                        // carry id/tier/status/runner), so the document-level value is the source.
                        FeatureId = manifest.FeatureId,
                        Tier = scenario.Tier ?? string.Empty,
                        AutomationStatus = scenario.AutomationStatus ?? string.Empty,
                        EvidenceState = "Missing",
                        Severity = "none"
                    };

                    EvidenceDocument evidence;
                    if (evidenceByScenario.TryGetValue(scenario.Id, out evidence))
                    {
                        row.EvidenceState = evidence.Passed ? "Green" : "Red";
                        DateTime completed;
                        if (DateTime.TryParse(evidence.CompletedUtc, null, System.Globalization.DateTimeStyles.RoundtripKind, out completed))
                        {
                            row.EvidenceCompletedUtc = completed.ToUniversalTime();
                            row.EvidenceAgeHours = Math.Round((nowUtc - row.EvidenceCompletedUtc.Value).TotalHours, 1);
                        }
                    }
                    else if (scenario.Tier == "visual-review" || scenario.Tier == "owner-manual")
                    {
                        row.EvidenceState = "NotApplicable";
                    }

                    if (scenario.Runner != null && !string.IsNullOrWhiteSpace(scenario.Runner.Path))
                    {
                        row.RunnerPath = scenario.Runner.Path;
                        row.RunnerArguments = scenario.Runner.Arguments == null
                            ? string.Empty
                            : string.Join(" ", scenario.Runner.Arguments.Select(QuoteIfNeeded));
                    }

                    row.Severity = DeriveSeverity(row);
                    rows.Add(row);
                }
            }
            return rows;
        }

        private static string QuoteIfNeeded(string value)
        {
            if (string.IsNullOrEmpty(value)) return "\"\"";
            return value.IndexOfAny(new[] { ' ', '\t' }) >= 0 ? "\"" + value + "\"" : value;
        }

        /// <summary>Issue severity for the window's severity navigation: failed evidence outranks
        /// missing evidence, which outranks stale evidence; visual/owner tiers are never an
        /// evidence issue because their contract is human review.</summary>
        private static string DeriveSeverity(FeatureAcceptanceScenarioStatus row)
        {
            if (row.EvidenceState == "Red") return "failed";
            var automatedTier = row.Tier == "core-automated" || row.Tier == "office";
            if (automatedTier && row.EvidenceState == "Missing" &&
                (row.AutomationStatus == "ready" || row.AutomationStatus == "running"))
                return "missing-evidence";
            if (row.EvidenceState == "Green" && row.EvidenceAgeHours >= 24 * 7) return "stale";
            return "none";
        }

        public static string BuildSummary(IEnumerable<FeatureAcceptanceScenarioStatus> rows, string repositoryRoot)
        {
            var list = rows.ToList();
            var builder = new StringBuilder();
            builder.AppendLine("openEfficiency feature acceptance status");
            builder.AppendLine("repository: " + repositoryRoot);
            builder.AppendLine("generated (UTC): " + DateTime.UtcNow.ToString("o"));
            builder.AppendLine("scenarios: " + list.Count +
                " | green: " + list.Count(r => r.EvidenceState == "Green") +
                " | red: " + list.Count(r => r.EvidenceState == "Red") +
                " | missing: " + list.Count(r => r.EvidenceState == "Missing") +
                " | issues: " + list.Count(r => r.Severity != "none"));
            foreach (var row in list.Where(r => r.Severity != "none").OrderBy(r => r.Severity).ThenByDescending(r => r.EvidenceAgeHours))
            {
                builder.AppendLine(string.Format("[{0}] {1} ({2}, {3}, age {4}h)",
                    row.Severity, row.ScenarioId, row.Tier, row.EvidenceState, row.EvidenceAgeHours));
            }
            return builder.ToString();
        }
    }
}
