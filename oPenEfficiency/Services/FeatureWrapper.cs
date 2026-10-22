using System.Reflection;
using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Runtime view of metadata declared through <see cref="FeatureMetadataAttribute"/>.
    /// Feature execution is handled exclusively by the command registry.
    /// </summary>
    public class FeatureWrapper
    {
        public string Id { get; }
        public string Name { get; }
        public FeatureDisplayInfo DisplayInfo { get; }

        public int MinSelection { get; }
        public int MaxSelection { get; }
        public PpSelectionType RequiredType { get; }
        public bool RequiresTable { get; }
        public bool RequiresHiddenShapes { get; }
        public bool IsToggle { get; }

        public FeatureWrapper(Type featureType)
        {
            // Get the metadata attribute
            var metadata = featureType.GetCustomAttribute<FeatureMetadataAttribute>();
            if (metadata == null)
            {
                throw new InvalidOperationException(
                    $"Type {featureType.Name} must have [FeatureMetadata] attribute");
            }

            Id = metadata.Id;
            Name = metadata.Name;
            MinSelection = metadata.MinSelection;
            MaxSelection = metadata.MaxSelection;
            RequiredType = ToPpSelectionType(metadata.RequiredType);
            RequiresTable = metadata.RequiresTable;
            RequiresHiddenShapes = metadata.RequiresHiddenShapes;
            IsToggle = metadata.IsToggle;

            // Build DisplayInfo from attribute
            DisplayInfo = new FeatureDisplayInfo
            {
                Tooltip = metadata.Tooltip ?? metadata.Name,
                IconData = metadata.IconData ?? "",
                Color = metadata.Color ?? "#6366F1",
                Description = metadata.Description ?? metadata.Tooltip ?? metadata.Name,
                DetailedHelpText = metadata.DetailedHelpText,
                HelpImagePath = metadata.HelpImagePath,
                SearchAliases = metadata.SearchAliases ?? Array.Empty<string>(),
                MinSelection = metadata.MinSelection,
                MaxSelection = metadata.MaxSelection,
                RequiredType = ToPpSelectionType(metadata.RequiredType),
                RequiresTable = metadata.RequiresTable,
                RequiresHiddenShapes = metadata.RequiresHiddenShapes,
                IsToggle = metadata.IsToggle
            };

        }

        /// <summary>
        /// Validates if the feature can be executed with current selection.
        /// </summary>
        public bool Validate(Selection selection, out string reason, SelectionMetadata meta = null)
        {
            reason = string.Empty;

            // 1. Hidden Shapes Check (Slide scope)
            if (RequiresHiddenShapes)
            {
                bool hasHidden = meta?.HasHiddenShapesOnSlide ?? false;
                
                // If no meta provided, do manual check
                if (meta == null)
                {
                    try
                    {
                        var slide = selection.Application.ActiveWindow.View.Slide as Slide;
                        if (slide != null)
                        {
                            for (int i = 1; i <= slide.Shapes.Count; i++)
                            {
                                if (slide.Shapes[i].Visible == Microsoft.Office.Core.MsoTriState.msoFalse)
                                {
                                    hasHidden = true;
                                    break;
                                }
                            }
                        }
                    }
                    catch { /* A stale shape is ignored while validating the remaining selection. */ }
                }

                if (!hasHidden)
                {
                    reason = "Requires hidden shapes on slide";
                    return false;
                }
            }

            // 2. Selection Type Check
            var selType = meta?.Type ?? (selection?.Type ?? PpSelectionType.ppSelectionNone);

            if (selection == null || selType == PpSelectionType.ppSelectionNone)
            {
                if (MinSelection > 0 || RequiresTable)
                {
                    reason = "No selection";
                    return false;
                }
                return true;
            }

            if (RequiredType != PpSelectionType.ppSelectionNone && selType != RequiredType)
            {
                reason = $"Requires {RequiredType}";
                return false;
            }

            // 3. Shape Count & Table Checks
            if (selType == PpSelectionType.ppSelectionShapes)
            {
                int count = 0;
                bool hasTable = false;

                if (meta != null)
                {
                    count = meta.ShapeCount;
                    hasTable = meta.HasTable;
                }
                else
                {
                    ShapeRange shapes = null;
                    try { shapes = selection.ShapeRange; } catch { /* Text-edit selections may not expose a ShapeRange. */ }
                    count = shapes?.Count ?? 0;

                    if (RequiresTable && shapes != null)
                    {
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            if (shapes[i].Type == Microsoft.Office.Core.MsoShapeType.msoTable)
                            {
                                hasTable = true;
                                break;
                            }
                        }
                    }
                }

                if (MinSelection > 0 && count < MinSelection)
                {
                    reason = $"Select at least {MinSelection} shape{(MinSelection > 1 ? "s" : "")}";
                    return false;
                }

                if (MaxSelection > 0 && count > MaxSelection)
                {
                    reason = $"Select no more than {MaxSelection} shapes";
                    return false;
                }

                if (RequiresTable && !hasTable)
                {
                    reason = "Select a table";
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Maps the host-neutral <see cref="OfficeSelectionType"/> from feature metadata
        /// to PowerPoint's native <see cref="PpSelectionType"/> used at runtime.
        /// This is the PowerPoint host boundary; other hosts (Excel, Word) provide their own mapping.
        /// </summary>
        private static PpSelectionType ToPpSelectionType(OfficeSelectionType type)
        {
            switch (type)
            {
                case OfficeSelectionType.None: return PpSelectionType.ppSelectionNone;
                case OfficeSelectionType.Shapes: return PpSelectionType.ppSelectionShapes;
                case OfficeSelectionType.Text: return PpSelectionType.ppSelectionText;
                case OfficeSelectionType.Slides: return PpSelectionType.ppSelectionSlides;
                // Cells has no PowerPoint equivalent; treat as "no specific requirement".
                case OfficeSelectionType.Cells: return PpSelectionType.ppSelectionNone;
                default: return PpSelectionType.ppSelectionShapes;
            }
        }
    }
}
