using System;
using System.Collections.Generic;
using System.Linq;
using oPenEfficiency.Models;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Services
{
    public static class SidebarPresetFactory
    {
        public const string ConsultingStandardName = "Consulting Standard";
        public const string AllFeaturesName = "All Features";
        public const string MinimalName = "Minimal";
        public const string LabsName = "Labs";

        private static readonly IReadOnlyList<string> PresetNamesValue =
            new[]
            {
                ConsultingStandardName,
                AllFeaturesName,
                MinimalName,
                LabsName
            };

        private static readonly IReadOnlyList<string> MinimalRequiredIdsValue =
            new[]
            {
                "BtnSearchBar",
                "BtnSizePanel",
                "BtnMatchSize",
                "BtnAlignLeft",
                "BtnAlignCenter",
                "BtnAlignRight",
                "BtnAlignTop",
                "BtnAlignCenterVertical",
                "BtnAlignBottom",
                "BtnDistributeHorizontal",
                "BtnDistributeVertical",
                "BtnFontPanel",
                "BtnAssetManager",
                "BtnSettings",
                "BtnExploreFeatures"
            };

        // Owner-approved V0.9 floor: expose existing EE counterpart entry points.
        // This is visibility coverage, not a claim of complete behavioral parity.
        private static readonly IReadOnlyList<string> EfficientElementsRequiredIdsValue =
            Array.AsReadOnly(new[]
            {
                "BtnAgendaWizard", "BtnAssetManager", "BtnTemplateManager",
                "BtnFlightMode", "BtnInsertText", "BtnSpecialChars",
                "BtnTextFormatPainter", "BtnPositionPainter", "BtnReplaceFont",
                "BtnMapWizard", "BtnCountryFlags", "BtnGanttChart",
                "BtnHarveyBall", "BtnTrafficLight", "BtnProgressSeries",
                "BtnStarRating", "BtnThermometer", "BtnExportWizard",
                "BtnStyleCheck", "BtnCleaner",
                "BtnAlignLeft", "BtnAlignCenter", "BtnAlignRight",
                "BtnAlignTop", "BtnAlignCenterVertical", "BtnAlignBottom",
                "BtnDistributeHorizontal", "BtnDistributeVertical",
                "BtnSwapPositions", "BtnArrangeGrid", "BtnAlignToTableCell",
                "BtnDockLeft", "BtnDockRight", "BtnDockTop", "BtnDockBottom",
                "BtnRemoveHorizontalGaps", "BtnRemoveVerticalGaps",
                "BtnMatchWidth", "BtnMatchHeight", "BtnMatchSize",
                "BtnStretchLeft", "BtnStretchRight", "BtnStretchTop", "BtnStretchBottom",
                "BtnStretchLeftEdge", "BtnStretchRightEdge", "BtnStretchTopEdge", "BtnStretchBottomEdge",
                "BtnMagicResizer", "BtnAlignShapeAdjustments", "BtnSmartCorners",
                "BtnSplitShape", "BtnRepeatShape", "BtnShapeFill",
                "BtnPickColor", "BtnPickFillColor", "BtnPickLineColor", "BtnPickTextColor", "BtnThemeColor",
                "BtnLearnMargin", "BtnSetMargin", "BtnFitFormToText",
                "BtnSplitByParagraphs", "BtnSpellCheck",
                "BtnHideSelected", "BtnShowHidden", "BtnHideMasterObjects",
                "BtnSlidePaste", "BtnSlideSize", "BtnConsolidateMasters",
                "BtnSelectSameType", "BtnOptimizeFreeForm", "BtnDeleteText",
                "BtnApplyTextTool", "BtnConvertTableToShapes", "BtnCreateMotionPath",
                "BtnTagInspector", "BtnSizePanel"
            });

        private static readonly IReadOnlyList<string> ConsultingArrangeOrderValue =
            Array.AsReadOnly(new[]
            {
                "BtnCopyXY", "BtnPasteXY", "BtnPositionPainter", "BtnSwapPositions",
                "BtnBringToFront", "BtnSendToBack", "BtnSelectSameType"
            });

        private static readonly IReadOnlyDictionary<FeatureCategory, string> CategoryNames =
            new Dictionary<FeatureCategory, string>
            {
                [FeatureCategory.AlignAndSize] = "ALIGN & SIZE",
                [FeatureCategory.ArrangeAndOrganize] = "ARRANGE & ORGANIZE",
                [FeatureCategory.TextAndFormat] = "TEXT & FORMAT",
                [FeatureCategory.Tables] = "TABLES",
                [FeatureCategory.VisualElements] = "VISUAL ELEMENTS",
                [FeatureCategory.StorylineAndReview] = "STORYLINE & REVIEW",
                [FeatureCategory.DataAndAutomation] = "DATA & AUTOMATION",
                [FeatureCategory.Libraries] = "LIBRARIES",
                [FeatureCategory.SettingsAndHelp] = "SETTINGS & HELP",
                [FeatureCategory.Labs] = "LABS"
            };

        private static readonly IReadOnlyDictionary<FeatureCategory, string> CategoryColors =
            new Dictionary<FeatureCategory, string>
            {
                [FeatureCategory.AlignAndSize] = "#10B981",
                [FeatureCategory.ArrangeAndOrganize] = "#F59E0B",
                [FeatureCategory.TextAndFormat] = "#FDE047",
                [FeatureCategory.Tables] = "#3B82F6",
                [FeatureCategory.VisualElements] = "#EF4444",
                [FeatureCategory.StorylineAndReview] = "#FDE047",
                [FeatureCategory.DataAndAutomation] = "#06B6D4",
                [FeatureCategory.Libraries] = "#8B5CF6",
                [FeatureCategory.SettingsAndHelp] = "#94A3B8",
                [FeatureCategory.Labs] = "#F97316"
            };

        public static IReadOnlyList<string> PresetNames => PresetNamesValue;
        public static IReadOnlyList<string> MinimalRequiredIds => MinimalRequiredIdsValue;
        public static IReadOnlyList<string> EfficientElementsRequiredIds => EfficientElementsRequiredIdsValue;

        public static SidebarConfig ResolveInitialLayout(
            SidebarConfig existingConfig,
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            if (existingConfig != null)
            {
                if (existingConfig.Sections == null)
                    existingConfig.Sections = new List<SidebarSection>();
                return existingConfig;
            }

            return Create(ConsultingStandardName, catalog);
        }

        public static SidebarConfig Create(
            string presetName,
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));

            switch (presetName)
            {
                case ConsultingStandardName:
                    return CreateConsultingStandard(catalog);
                case AllFeaturesName:
                    return CreateAllFeatures(catalog);
                case MinimalName:
                    return CreateMinimal(catalog);
                case LabsName:
                    return CreateLabs(catalog);
                default:
                    throw new ArgumentException($"Unknown sidebar preset '{presetName}'.", nameof(presetName));
            }
        }

        public static IReadOnlyList<FeaturePresetDefinition> CreateDefinitions(
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            return PresetNames
                .Select(name =>
                {
                    var definition = new FeaturePresetDefinition
                    {
                        Name = name,
                        FeatureIds = Flatten(Create(name, catalog)),
                        IsAutomaticallyGenerated = true,
                        IsAllFeaturesPreset = name == AllFeaturesName,
                        IsProfessionalDefault = name == ConsultingStandardName,
                        IsLabsPreset = name == LabsName
                    };
                    if (name == MinimalName)
                        definition.RequiredFeatureIds = MinimalRequiredIds;
                    return definition;
                })
                .ToList()
                .AsReadOnly();
        }

        public static string GetDescription(string presetName)
        {
            switch (presetName)
            {
                case ConsultingStandardName:
                    return "Stable commands grouped by category, including existing Efficient Elements counterpart entry points.";
                case AllFeaturesName:
                    return "Every available command, with Preview and Experimental commands clearly separated.";
                case MinimalName:
                    return "Core search, sizing, alignment, distribution, font, library, and help capabilities.";
                case LabsName:
                    return "Preview and Experimental commands only. Included behavior may change.";
                default:
                    return string.Empty;
            }
        }

        private static SidebarConfig CreateConsultingStandard(
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            var requiredIds = new HashSet<string>(EfficientElementsRequiredIdsValue, StringComparer.Ordinal);
            var included = catalog.Where(record =>
                IsAvailable(record) &&
                record.Maturity == FeatureMaturity.Stable &&
                (record.DefaultPlacement == FeatureDefaultPlacement.ProfessionalDefault ||
                 requiredIds.Contains(record.Id)));

            var config = CreateCategorizedLayout(included, includeSearchFirst: true, includePersistentFinal: true);
            var arrange = config.Sections.FirstOrDefault(section =>
                section.Name == CategoryNames[FeatureCategory.ArrangeAndOrganize]);
            if (arrange != null)
            {
                var positions = ConsultingArrangeOrderValue
                    .Select((id, index) => new { id, index })
                    .ToDictionary(item => item.id, item => item.index, StringComparer.Ordinal);
                arrange.Features = arrange.Features.OrderBy(feature =>
                    positions.TryGetValue(feature.Id, out var position) ? position : int.MaxValue).ToList();
            }
            return config;
        }

        private static SidebarConfig CreateAllFeatures(
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            var config = CreateCategorizedLayout(
                catalog.Where(record =>
                    IsAvailable(record) &&
                    record.Maturity == FeatureMaturity.Stable),
                includeSearchFirst: false,
                includePersistentFinal: false);

            AddMaturitySection(
                config,
                "PREVIEW — BEHAVIOR MAY CHANGE",
                "#6366F1",
                catalog.Where(record => IsAvailable(record) && record.Maturity == FeatureMaturity.Preview));
            AddMaturitySection(
                config,
                "LABS — EXPERIMENTAL",
                "#F97316",
                catalog.Where(record => IsAvailable(record) && record.Maturity == FeatureMaturity.Experimental));
            MovePersistentActionsToFinalSection(config, catalog);
            return config;
        }

        private static SidebarConfig CreateMinimal(
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            var byId = CatalogById(catalog);
            var config = new SidebarConfig();

            AddSection(config, "COMMAND SEARCH", "#10B981", Features(byId, "BtnSearchBar"));
            AddSection(
                config,
                CategoryNames[FeatureCategory.AlignAndSize],
                CategoryColors[FeatureCategory.AlignAndSize],
                Features(
                    byId,
                    "BtnSizePanel", "BtnMatchSize",
                    "BtnAlignLeft", "BtnAlignCenter", "BtnAlignRight",
                    "BtnAlignTop", "BtnAlignCenterVertical", "BtnAlignBottom",
                    "BtnDistributeHorizontal", "BtnDistributeVertical"));
            AddSection(
                config,
                CategoryNames[FeatureCategory.TextAndFormat],
                CategoryColors[FeatureCategory.TextAndFormat],
                Features(byId, "BtnFontPanel"));
            AddSection(
                config,
                CategoryNames[FeatureCategory.Libraries],
                CategoryColors[FeatureCategory.Libraries],
                Features(byId, "BtnAssetManager"));
            AddSection(
                config,
                CategoryNames[FeatureCategory.SettingsAndHelp],
                CategoryColors[FeatureCategory.SettingsAndHelp],
                Features(byId, "BtnSettings", "BtnExploreFeatures"));
            return config;
        }

        private static SidebarConfig CreateLabs(
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            var config = new SidebarConfig();
            AddMaturitySection(
                config,
                "PREVIEW — BEHAVIOR MAY CHANGE",
                "#6366F1",
                catalog.Where(record => IsAvailable(record) && record.Maturity == FeatureMaturity.Preview));
            AddMaturitySection(
                config,
                "EXPERIMENTAL — BEHAVIOR MAY CHANGE",
                "#F97316",
                catalog.Where(record => IsAvailable(record) && record.Maturity == FeatureMaturity.Experimental));
            return config;
        }

        private static SidebarConfig CreateCategorizedLayout(
            IEnumerable<FeatureCatalogRecord> records,
            bool includeSearchFirst,
            bool includePersistentFinal)
        {
            var ordered = records.ToList();
            var config = new SidebarConfig();
            var excludedIds = new HashSet<string>(StringComparer.Ordinal);

            if (includeSearchFirst)
            {
                AddSection(
                    config,
                    "COMMAND SEARCH",
                    "#10B981",
                    ordered.Where(record => record.Id == "BtnSearchBar"));
                excludedIds.Add("BtnSearchBar");
            }

            if (includePersistentFinal)
            {
                excludedIds.Add("BtnSettings");
                excludedIds.Add("BtnExploreFeatures");
            }

            foreach (var category in CategoryNames.Keys.Where(category =>
                category != FeatureCategory.SettingsAndHelp &&
                category != FeatureCategory.Labs))
            {
                AddSection(
                    config,
                    CategoryNames[category],
                    CategoryColors[category],
                    ordered.Where(record =>
                        record.Category == category && !excludedIds.Contains(record.Id)));
            }

            if (!includePersistentFinal)
            {
                AddSection(
                    config,
                    CategoryNames[FeatureCategory.SettingsAndHelp],
                    CategoryColors[FeatureCategory.SettingsAndHelp],
                    ordered.Where(record =>
                        record.Category == FeatureCategory.SettingsAndHelp &&
                        !excludedIds.Contains(record.Id)));
            }
            else
            {
                var byId = CatalogById(ordered);
                AddSection(
                    config,
                    CategoryNames[FeatureCategory.SettingsAndHelp],
                    CategoryColors[FeatureCategory.SettingsAndHelp],
                    Features(byId, "BtnSettings", "BtnExploreFeatures"));
            }

            return config;
        }

        private static void MovePersistentActionsToFinalSection(
            SidebarConfig config,
            IReadOnlyList<FeatureCatalogRecord> catalog)
        {
            var persistentIds = new HashSet<string>(
                new[] { "BtnSettings", "BtnExploreFeatures" },
                StringComparer.Ordinal);
            foreach (var section in config.Sections)
                section.Features.RemoveAll(feature => persistentIds.Contains(feature.Id));
            config.Sections.RemoveAll(section => section.Features.Count == 0);

            var byId = CatalogById(catalog);
            AddSection(
                config,
                CategoryNames[FeatureCategory.SettingsAndHelp],
                CategoryColors[FeatureCategory.SettingsAndHelp],
                Features(byId, "BtnSettings", "BtnExploreFeatures"));
        }

        private static void AddMaturitySection(
            SidebarConfig config,
            string name,
            string color,
            IEnumerable<FeatureCatalogRecord> records)
        {
            AddSection(
                config,
                name,
                color,
                records.OrderBy(record => record.Category).ThenBy(record => record.Name, StringComparer.Ordinal));
        }

        private static void AddSection(
            SidebarConfig config,
            string name,
            string color,
            IEnumerable<FeatureCatalogRecord> records)
        {
            var features = records
                .Where(record => record != null)
                .GroupBy(record => record.Id, StringComparer.Ordinal)
                .Select(group => group.First())
                .Select(ToSidebarFeature)
                .ToList();
            if (features.Count == 0) return;

            config.Sections.Add(new SidebarSection
            {
                Name = name,
                Color = color,
                Features = features
            });
        }

        private static IEnumerable<FeatureCatalogRecord> Features(
            IReadOnlyDictionary<string, FeatureCatalogRecord> byId,
            params string[] ids)
        {
            foreach (var id in ids)
            {
                if (!byId.TryGetValue(id, out var record))
                    throw new InvalidOperationException($"Preset requires unknown catalog feature '{id}'.");
                if (!IsAvailable(record))
                    throw new InvalidOperationException($"Preset requires unavailable catalog feature '{id}'.");
                yield return record;
            }
        }

        private static IReadOnlyDictionary<string, FeatureCatalogRecord> CatalogById(
            IEnumerable<FeatureCatalogRecord> catalog)
        {
            return catalog
                .Where(record => record != null && !string.IsNullOrWhiteSpace(record.Id))
                .GroupBy(record => record.Id, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        }

        private static IReadOnlyList<string> Flatten(SidebarConfig config)
        {
            return config.Sections
                .SelectMany(section => section.Features)
                .Select(feature => feature.Id)
                .ToList()
                .AsReadOnly();
        }

        private static SidebarFeature ToSidebarFeature(FeatureCatalogRecord record)
        {
            return new SidebarFeature { Id = record.Id, Name = record.Name };
        }

        private static bool IsAvailable(FeatureCatalogRecord record)
        {
            return record != null && record.UserVisible && !record.IsObsolete;
        }
    }
}
