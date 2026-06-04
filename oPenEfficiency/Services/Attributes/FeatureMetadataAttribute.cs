using System;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Services.Attributes
{
    /// <summary>
    /// Metadata attribute for auto-discovering and registering features.
    /// Apply to static feature classes to eliminate manual registration.
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

        // Selection requirements
        public int MinSelection { get; set; } = 0;
        public int MaxSelection { get; set; } = 0; // 0 = unlimited
        public PpSelectionType RequiredType { get; set; } = PpSelectionType.ppSelectionShapes;

        // Feature type
        public bool RequiresTable { get; set; }
        public bool RequiresHiddenShapes { get; set; }
        public bool IsToggle { get; set; }

        // Search
        public string Keywords { get; set; }
    }
}
