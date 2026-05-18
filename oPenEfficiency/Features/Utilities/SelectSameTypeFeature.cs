using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Select similar shapes based on the first selected shape's properties.
    /// Delegates to AdvancedSelectionFeature with a pre-defined criteria.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSelectSameType",
        Name = "Select Same Type",
        Tooltip = "Select same type",
        IconData = "M19,3H5C3.9,3 3,3.9 3,5V19C3,20.1 3.9,21 5,21H19C20.1,21 21,20.1 21,19V5C21,3.9 20.1,3 19,3M19,19H5V5H19V19M7,7H10V10H7V7M11,7H14V10H11V7M15,7H18V10H15V7M7,11H10V14H7V11M11,11H14V14H11V11M15,11H18V14H15V11M7,15H10V18H7V15M11,15H14V18H11V15M15,15H18V18H15V15Z",
        Color = "#8E44AD",
        Description = "Instantly selects all other objects on the slide that match the type of your current selection.",
        DetailedHelpText = "### Select Same Type\nSelects all shapes on the current slide that share the same AutoShapeType or fill color as the currently selected shape, enabling bulk formatting.",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class SelectSameTypeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return AdvancedSelectionFeature.Execute(manager, AdvancedSelectionFeature.SelectSameCriteria.Type);
        }

        public static bool ExecuteWithCriteria(PowerPointManager manager, int criteriaIndex)
        {
            var criteria = (AdvancedSelectionFeature.SelectSameCriteria)criteriaIndex;
            return AdvancedSelectionFeature.Execute(manager, criteria);
        }

        public static bool Execute(PowerPointManager manager, AdvancedSelectionFeature.SelectSameCriteria criteria)
        {
            return AdvancedSelectionFeature.Execute(manager, criteria);
        }
    }
}
