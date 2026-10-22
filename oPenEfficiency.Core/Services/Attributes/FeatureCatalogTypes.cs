using System;
using System.Collections.Generic;

namespace oPenEfficiency.Services.Attributes
{
    public enum FeatureCategory
    {
        Unknown = 0,
        AlignAndSize,
        ArrangeAndOrganize,
        TextAndFormat,
        Tables,
        VisualElements,
        StorylineAndReview,
        DataAndAutomation,
        Libraries,
        SettingsAndHelp,
        Labs
    }

    public enum FeatureMaturity
    {
        Unknown = 0,
        Stable,
        Preview,
        Experimental
    }

    public enum FeatureDefaultPlacement
    {
        Unknown = 0,
        CatalogOnly,
        ProfessionalDefault,
        Labs
    }

    public enum FeatureExecutionKind
    {
        Command,
        Widget
    }

    public sealed class FeatureCommandContract
    {
        public string Id { get; set; }
        public bool IsToggle { get; set; }
        public bool HasValidPrimaryArguments { get; set; } = true;
    }

    public sealed class FeatureCatalogRecord
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Tooltip { get; set; }
        public string Description { get; set; }
        public string DetailedHelpText { get; set; }
        public string IconData { get; set; }
        public string Color { get; set; }
        public string HelpImagePath { get; set; }
        public FeatureCategory Category { get; set; }
        public IReadOnlyList<string> SearchAliases { get; set; } = Array.Empty<string>();
        public int MinSelection { get; set; }
        public int MaxSelection { get; set; }
        public OfficeSelectionType RequiredType { get; set; }
        public bool RequiresTable { get; set; }
        public bool RequiresHiddenShapes { get; set; }
        public bool IsToggle { get; set; }
        public FeatureMaturity Maturity { get; set; }
        public bool UserVisible { get; set; }
        public FeatureDefaultPlacement DefaultPlacement { get; set; }
        public bool SupportsUserShortcut { get; set; }
        public IReadOnlyList<string> Variants { get; set; } = Array.Empty<string>();
        public bool IsObsolete { get; set; }
        public bool IsManual { get; set; }
        public string ManualOverrideReason { get; set; }
        public FeatureExecutionKind ExecutionKind { get; set; } = FeatureExecutionKind.Command;
    }

    public sealed class FeaturePresetDefinition
    {
        public string Name { get; set; }
        public IReadOnlyList<string> FeatureIds { get; set; } = Array.Empty<string>();
        public IReadOnlyList<string> RequiredFeatureIds { get; set; } = Array.Empty<string>();
        public bool IsAutomaticallyGenerated { get; set; }
        public bool IsAllFeaturesPreset { get; set; }
        public bool IsProfessionalDefault { get; set; }
        public bool IsLabsPreset { get; set; }
    }
}
