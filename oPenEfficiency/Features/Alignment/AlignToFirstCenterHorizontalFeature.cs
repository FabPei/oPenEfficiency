using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Align selected shapes horizontally centered relative to the first/master object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAlignCenter",
        Name = "Align Center",
        Tooltip = "Align center horizontal",
        IconData = "M11,2H13V6H21V8H13V16H18V18H13V22H11V18H6V16H11V8H3V6H11V2Z",
        Color = "#10B981",
        Description = "Centers all selected objects horizontally relative to the anchor object's center point. With single selection, centers on slide.",
        DetailedHelpText = "### Align Center\nCenters all selected shapes horizontally relative to the horizontal midpoint of the first shape.",
        Keywords = "center horizontally, midpoint, center align, horizontal alignment",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class AlignToFirstCenterHorizontalFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return AlignmentHelpers.ExecuteAlignment(manager, "centerh");
        }
    }
}
