using System;
using System.Collections.Generic;
using System.Linq;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Services
{
    public static class FeatureCatalogValidator
    {
        public static IReadOnlyList<string> ValidateCatalog(IEnumerable<FeatureCatalogRecord> records)
        {
            var catalog = (records ?? Enumerable.Empty<FeatureCatalogRecord>()).ToList();
            var errors = new List<string>();

            foreach (var duplicate in catalog
                .Where(record => record != null && !string.IsNullOrWhiteSpace(record.Id))
                .GroupBy(record => record.Id, StringComparer.Ordinal)
                .Where(group => group.Count() > 1))
            {
                errors.Add($"Duplicate feature ID '{duplicate.Key}'.");
            }

            foreach (var record in catalog)
            {
                if (record == null)
                {
                    errors.Add("Catalog contains a null record.");
                    continue;
                }

                if (string.IsNullOrWhiteSpace(record.Id))
                    errors.Add("Catalog contains an empty feature ID.");

                if (record.UserVisible)
                {
                    Require(record.Id, "Name", record.Name, errors);
                    Require(record.Id, "Description", record.Description, errors);
                    Require(record.Id, "DetailedHelpText", record.DetailedHelpText, errors);

                    string normalizedTooltip = NormalizeDisplayText(record.Tooltip);
                    if (!string.IsNullOrEmpty(normalizedTooltip))
                    {
                        if (normalizedTooltip == NormalizeDisplayText(record.Id))
                            errors.Add($"Feature '{record.Id}' uses its stable ID as tooltip copy.");
                        if (normalizedTooltip == NormalizeDisplayText(record.Name))
                            errors.Add($"Feature '{record.Id}' duplicates its Name in Tooltip.");
                        if ((record.SearchAliases ?? Array.Empty<string>())
                            .Any(alias => normalizedTooltip == NormalizeDisplayText(alias)))
                            errors.Add($"Feature '{record.Id}' uses a legacy/search alias as Tooltip.");
                        if (record.Tooltip.IndexOf("right-click", StringComparison.OrdinalIgnoreCase) >= 0)
                            errors.Add($"Feature '{record.Id}' embeds a right-click hint in Tooltip: '{record.Tooltip}'.");
                    }
                }

                if (!Enum.IsDefined(typeof(FeatureCategory), record.Category) ||
                    record.Category == FeatureCategory.Unknown)
                    errors.Add($"Feature '{record.Id}' has an invalid category.");

                if (!Enum.IsDefined(typeof(FeatureMaturity), record.Maturity) ||
                    record.Maturity == FeatureMaturity.Unknown)
                    errors.Add($"Feature '{record.Id}' has an invalid maturity.");

                if (!Enum.IsDefined(typeof(FeatureDefaultPlacement), record.DefaultPlacement) ||
                    record.DefaultPlacement == FeatureDefaultPlacement.Unknown)
                    errors.Add($"Feature '{record.Id}' has an invalid default placement.");

                if (record.MinSelection < 0 || record.MaxSelection < 0 ||
                    (record.MaxSelection > 0 && record.MaxSelection < record.MinSelection))
                    errors.Add($"Feature '{record.Id}' has contradictory selection counts.");

                if (record.RequiresTable &&
                    record.RequiredType != OfficeSelectionType.Shapes &&
                    record.RequiredType != OfficeSelectionType.None)
                    errors.Add($"Feature '{record.Id}' requires a table but has incompatible selection type '{record.RequiredType}'.");

                if (record.IsManual && !string.IsNullOrWhiteSpace(record.ManualOverrideReason))
                    continue;
            }

            return errors;
        }

        private static string NormalizeDisplayText(string value)
        {
            return new string((value ?? string.Empty)
                .Where(char.IsLetterOrDigit)
                .Select(char.ToLowerInvariant)
                .ToArray());
        }

        public static IReadOnlyList<string> ValidatePresets(
            IEnumerable<FeatureCatalogRecord> records,
            IEnumerable<FeaturePresetDefinition> presets)
        {
            var catalog = (records ?? Enumerable.Empty<FeatureCatalogRecord>()).ToList();
            var byId = catalog
                .Where(record => record != null && !string.IsNullOrWhiteSpace(record.Id))
                .GroupBy(record => record.Id, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            var errors = new List<string>();

            foreach (var preset in presets ?? Enumerable.Empty<FeaturePresetDefinition>())
            {
                if (preset == null)
                {
                    errors.Add("Preset collection contains a null definition.");
                    continue;
                }

                var ids = (preset.FeatureIds ?? Array.Empty<string>()).ToList();
                foreach (var id in ids)
                {
                    if (!byId.TryGetValue(id ?? string.Empty, out var record))
                        errors.Add($"Preset '{preset.Name}' references unknown feature ID '{id}'.");
                    else if (record.IsObsolete)
                        errors.Add($"Preset '{preset.Name}' references obsolete feature ID '{id}'.");
                }

                if (preset.IsAutomaticallyGenerated)
                {
                    foreach (var duplicate in ids.GroupBy(id => id, StringComparer.Ordinal).Where(group => group.Count() > 1))
                        errors.Add($"Automatically generated preset '{preset.Name}' contains duplicate feature ID '{duplicate.Key}'.");
                }

                if (preset.IsAllFeaturesPreset)
                {
                    var included = new HashSet<string>(ids, StringComparer.Ordinal);
                    foreach (var record in catalog.Where(record =>
                        record != null && record.UserVisible && !record.IsObsolete))
                    {
                        if (!included.Contains(record.Id))
                            errors.Add($"User-visible feature '{record.Id}' is absent from All Features preset '{preset.Name}'.");
                    }
                }

                if (preset.IsProfessionalDefault)
                {
                    foreach (var id in ids)
                    {
                        if (byId.TryGetValue(id ?? string.Empty, out var record) &&
                            record.Maturity == FeatureMaturity.Experimental)
                            errors.Add($"Experimental feature '{id}' appears in professional default preset '{preset.Name}'.");
                    }
                }

                if (preset.IsLabsPreset)
                {
                    foreach (var id in ids)
                    {
                        if (byId.TryGetValue(id ?? string.Empty, out var record) &&
                            record.Maturity != FeatureMaturity.Preview &&
                            record.Maturity != FeatureMaturity.Experimental)
                            errors.Add($"Stable feature '{id}' appears in Labs preset '{preset.Name}'.");
                    }
                }

                var includedIds = new HashSet<string>(ids, StringComparer.Ordinal);
                foreach (var requiredId in preset.RequiredFeatureIds ?? Array.Empty<string>())
                {
                    if (!includedIds.Contains(requiredId))
                        errors.Add($"Required feature '{requiredId}' is absent from preset '{preset.Name}'.");
                }
            }

            return errors;
        }

        public static IReadOnlyList<string> ValidateManualShadowing(
            IEnumerable<string> attributedIds,
            IEnumerable<FeatureCatalogRecord> manualRecords)
        {
            var attributed = new HashSet<string>(attributedIds ?? Enumerable.Empty<string>(), StringComparer.Ordinal);
            return (manualRecords ?? Enumerable.Empty<FeatureCatalogRecord>())
                .Where(record => record != null &&
                    attributed.Contains(record.Id) &&
                    string.IsNullOrWhiteSpace(record.ManualOverrideReason))
                .Select(record => $"Manual catalog record '{record.Id}' shadows attributed metadata without an override reason.")
                .ToList();
        }

        public static IReadOnlyList<string> ValidateCommands(
            IEnumerable<FeatureCatalogRecord> records,
            IEnumerable<FeatureCommandContract> commands)
        {
            var catalog = (records ?? Enumerable.Empty<FeatureCatalogRecord>())
                .Where(record => record != null && !string.IsNullOrWhiteSpace(record.Id))
                .ToList();
            var commandList = (commands ?? Enumerable.Empty<FeatureCommandContract>())
                .Where(command => command != null)
                .ToList();
            var errors = new List<string>();
            var catalogById = catalog
                .GroupBy(record => record.Id, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

            foreach (var duplicate in commandList
                .Where(command => !string.IsNullOrWhiteSpace(command.Id))
                .GroupBy(command => command.Id, StringComparer.Ordinal)
                .Where(group => group.Count() > 1))
                errors.Add($"Duplicate command ID '{duplicate.Key}'.");

            foreach (var command in commandList)
            {
                if (string.IsNullOrWhiteSpace(command.Id))
                {
                    errors.Add("Command registry contains an empty ID.");
                    continue;
                }
                if (!catalogById.TryGetValue(command.Id, out var record))
                {
                    errors.Add($"Command registry contains unknown feature ID '{command.Id}'.");
                    continue;
                }
                if (record.ExecutionKind == FeatureExecutionKind.Widget)
                    errors.Add($"Widget feature '{command.Id}' must not be registered as a command.");
                if (record.IsToggle != command.IsToggle)
                    errors.Add($"Feature '{command.Id}' toggle metadata does not match its command kind.");
                if (!command.HasValidPrimaryArguments)
                    errors.Add($"Command '{command.Id}' has invalid primary arguments.");
            }

            var commandIds = new HashSet<string>(
                commandList.Where(command => !string.IsNullOrWhiteSpace(command.Id))
                    .Select(command => command.Id),
                StringComparer.Ordinal);
            foreach (var record in catalog.Where(record =>
                record.UserVisible && !record.IsObsolete &&
                record.ExecutionKind == FeatureExecutionKind.Command))
            {
                if (!commandIds.Contains(record.Id))
                    errors.Add($"User-visible command feature '{record.Id}' has no registered command.");
            }

            return errors;
        }

        private static void Require(string id, string field, string value, ICollection<string> errors)
        {
            if (string.IsNullOrWhiteSpace(value))
                errors.Add($"User-visible feature '{id}' lacks required {field}.");
        }
    }
}
