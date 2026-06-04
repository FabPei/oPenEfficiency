using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnShowHidden",
        Name = "Show Hidden",
        Tooltip = "Show hidden",
        IconData = "M12,4.5C7,4.5 2.73,7.61 1,12C2.73,16.39 7,19.5 12,19.5C17,19.5 21.27,16.39 23,12C21.27,7.61 17,4.5 12,4.5M12,17.5A5.5,5.5 0 0,1 6.5,12A5.5,5.5 0 0,1 12,6.5A5.5,5.5 0 0,1 17.5,12A5.5,5.5 0 0,1 12,17.5M12,9A3,3 0 0,0 9,12A3,3 0 0,0 12,15A3,3 0 0,0 15,12A3,3 0 0,0 12,9Z",
        Color = "#D946EF",
        Description = "Makes all currently hidden shapes on the slide visible again.",
        DetailedHelpText = "### Show Hidden\nRestores all shapes previously hidden by the Hide Selected feature, making them visible again at their original position and Z-order.",
        Keywords = "show, hidden, unhide, reveal, visible",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class ShowHiddenFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return VisibilityFeature.ShowAllHiddenShapes(manager);
        }
    }
}
