using System;

namespace oPenEfficiency.Services.Attributes
{
    /// <summary>
    /// Metadata attribute for auto-discovering and registering features.
    /// Apply to static feature classes to eliminate manual registration.
    ///
    /// Host-neutral: lives in oPenEfficiency.Core so PowerPoint, Excel, Word and
    /// OneNote add-ins can all reuse the same feature-discovery mechanism.
    /// Selection requirements use <see cref="OfficeSelectionType"/>; each host maps
    /// it to its native selection enum at the boundary.
    /// </summary>
    /// <example>
    /// [FeatureMetadata(
    ///     Id = "BtnAlignLeft",
    ///     Name = "Align Left",
    ///     Tooltip = "Align left edges",
    ///     Icon = "←",
    ///     Color = "#10B981",
    ///     MinSelection = 2)]
    /// public static class AlignLeftFeature { ... }
    /// </example>
    [AttributeUsage(AttributeTargets.Class, Inherited = false)]
    public class FeatureMetadataAttribute : Attribute
    {
        // Required
        public string Id { get; set; }
        public string Name { get; set; }

        // Display
        public string Tooltip { get; set; }
        public string IconData { get; set; }
        public string Color { get; set; }
        public string Description { get; set; }
        public string DetailedHelpText { get; set; }
        public string HelpImagePath { get; set; }
        public FeatureCategory Category { get; set; } = FeatureCategory.Unknown;
        public string[] SearchAliases { get; set; } = new string[0];

        // Selection requirements
        public int MinSelection { get; set; } = 0;
        public int MaxSelection { get; set; } = 0; // 0 = unlimited
        public OfficeSelectionType RequiredType { get; set; } = OfficeSelectionType.Shapes;

        // Feature type
        public bool RequiresTable { get; set; }
        public bool RequiresHiddenShapes { get; set; }
        public bool IsToggle { get; set; }

        // Catalog policy
        public FeatureMaturity Maturity { get; set; } = FeatureMaturity.Stable;
        public bool UserVisible { get; set; } = true;
        public FeatureDefaultPlacement DefaultPlacement { get; set; } = FeatureDefaultPlacement.CatalogOnly;
        public bool SupportsUserShortcut { get; set; } = true;
        public string[] Variants { get; set; } = new string[0];
    }
}
