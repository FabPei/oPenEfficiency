using System.Reflection;
using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using oPenEfficiency.UI;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Wrapper that uses reflection to execute static feature classes.
    /// Reads metadata from [FeatureMetadata] attribute.
    /// </summary>
    public class FeatureWrapper
    {
        private readonly Type _featureType;
        private readonly MethodInfo _executeMethod;

        public string Id { get; }
        public string Name { get; }
        public FeatureDisplayInfo DisplayInfo { get; }

        public int MinSelection { get; }
        public int MaxSelection { get; }
        public PpSelectionType RequiredType { get; }
        public bool RequiresTable { get; }
        public bool RequiresHiddenShapes { get; }
        public bool IsToggle { get; }
        public string Keywords { get; }

        public FeatureWrapper(Type featureType)
        {
            _featureType = featureType;

            // Get the metadata attribute
            var metadata = featureType.GetCustomAttribute<FeatureMetadataAttribute>();
            if (metadata == null)
            {
                throw new InvalidOperationException(
                    $"Type {featureType.Name} must have [FeatureMetadata] attribute");
            }

            Id = metadata.Id;
            Name = metadata.Name;
            Keywords = metadata.Keywords;
            MinSelection = metadata.MinSelection;
            MaxSelection = metadata.MaxSelection;
            RequiredType = metadata.RequiredType;
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
                MinSelection = metadata.MinSelection,
                MaxSelection = metadata.MaxSelection,
                RequiredType = metadata.RequiredType,
                RequiresTable = metadata.RequiresTable,
                RequiresHiddenShapes = metadata.RequiresHiddenShapes,
                IsToggle = metadata.IsToggle,
                Keywords = metadata.Keywords
            };

            // Find Execute method (must exist)
            // Try ToggleButton version first if it's a toggle feature
            if (IsToggle)
            {
                _executeMethod = featureType.GetMethod(
                    "Execute",
                    BindingFlags.Public | BindingFlags.Static,
                    null,
                    new[] { typeof(PowerPointManager), typeof(System.Windows.Controls.Primitives.ToggleButton) },
                    null);
            }

            if (_executeMethod == null)
            {
                _executeMethod = featureType.GetMethod(
                    "Execute",
                    BindingFlags.Public | BindingFlags.Static,
                    null,
                    new[] { typeof(PowerPointManager) },
                    null);
            }

            if (_executeMethod == null)
            {
                throw new InvalidOperationException(
                    $"Feature {featureType.Name} must have a public static Execute(PowerPointManager) method");
            }
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
                    catch { }
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
                    try { shapes = selection.ShapeRange; } catch { }
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
        /// Executes the feature via reflection.
        /// </summary>
        /// <param name="manager">PowerPointManager instance</param>
        /// <param name="toggleBtn">Optional ToggleButton (for toggle features)</param>
        /// <returns>True if successful, false otherwise</returns>
        public bool Execute(PowerPointManager manager, System.Windows.Controls.Primitives.ToggleButton toggleBtn = null)
        {
            try
            {
                var parameters = _executeMethod.GetParameters();
                object[] args;
                if (parameters.Length == 2 && parameters[1].ParameterType == typeof(System.Windows.Controls.Primitives.ToggleButton))
                {
                    args = new object[] { manager, toggleBtn };
                }
                else
                {
                    args = new object[] { manager };
                }

                var result = _executeMethod.Invoke(null, args);
                return result is bool b && b;
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, $"FeatureWrapper.Execute({Id})");
                return false;
            }
        }
    }
}
