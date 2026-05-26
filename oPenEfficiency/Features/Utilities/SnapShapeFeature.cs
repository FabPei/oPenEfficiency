using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnSnapShape",
        Name = "Snap Shape",
        Tooltip = "Snap shapes together (Object Connector)",
        IconData = "M17 19a1 1 0 0 1-1-1v-2a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v2a1 1 0 0 1-1 1z M17 21v-2 M19 14V6.5a1 1 0 0 0-7 0v11a1 1 0 0 1-7 0V10 M21 21v-2 M3 5V3 M4 10a2 2 0 0 1-2-2V6a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2a2 2 0 0 1-2 2z M7 5V3",
        Color = "#FBBF24",
        Description = "Opens the Object Connector window to snap shape edges and centers together with precision.",
        DetailedHelpText = "### Snap Shape\nOpens the Object Connector floating toolbar. Use it to precisely snap the edges or centers of shapes together and to create visual connector lines between objects.",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone,
        IsToggle = true)]
    public static class SnapShapeFeature
    {
        public static bool Execute(PowerPointManager manager) => false;
    }
}
