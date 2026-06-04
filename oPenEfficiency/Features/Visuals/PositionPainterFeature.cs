using System;
using oPenEfficiency.Services.Attributes;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Visuals
{
    [FeatureMetadata(
        Id = "BtnPositionPainter",
        Name = "Position Painter",
        Tooltip = "Paint position/size properties to other shapes",
        IconData = "M17,12V3H7V12H17M14,10H10V5H14V10M20,7V14A2,2 0 0,1 18,16H7V14H18V7H20M5,12V21H7V12H5Z",
        Color = "#10B981",
        Description = "Learns position and dimensions of the selected shape and applies selected properties to others.",
        DetailedHelpText = "Opens the Position Painter.",
        Keywords = "position, painter, copy, size, dimensions, format",
        MinSelection = 1,
        MaxSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class PositionPainterFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            var window = new oPenEfficiency.UI.PositionPainterWindow(manager);
            window.Show(); // Non-modal so user can click on presentation
            return true;
        }
    }
}