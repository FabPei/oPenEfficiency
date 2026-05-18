using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Align selected shapes vertically centered relative to the first/master object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAlignCenterVertical",
        Name = "Align Center Vertical",
        Tooltip = "Align center vertical",
        IconData = "M2,11V13H6V21H8V13H16V18H18V13H22V11H18V6H16V11H8V3H6V11H2Z",
        Color = "#10B981",
        Description = "Centers objects vertically relative to the anchor object's center point. With single selection, centers vertically on slide.",
        DetailedHelpText = "### Align Center Vertical\nCenters all selected shapes vertically relative to the vertical midpoint of the first shape.",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class AlignToFirstCenterVerticalFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return AlignmentHelpers.ExecuteAlignment(manager, "centerv");
        }
    }
}
