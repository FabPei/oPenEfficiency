using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI;
using System.Windows;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnTransparentColor",
        Name = "Color Eraser",
        Tooltip = "Make color transparent",
        IconData = "M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12A10,10 0 0,0 12,2M12,4A8,8 0 0,1 20,12A8,8 0 0,1 12,20A8,8 0 0,1 4,12A8,8 0 0,1 12,4M12,6A6,6 0 0,0 6,12A6,6 0 0,0 12,18A6,6 0 0,0 18,12A6,6 0 0,0 12,6Z",
        Color = "#06B6D4",
        Description = "Opens a color picker tool to select and erase (make transparent) a specific color from images in your slide.",
        DetailedHelpText = "### Color Eraser\nRemoves the fill color from all selected shapes, making them fully transparent while preserving their border and text.",
        Keywords = "transparent, eraser, color, remove, clear, image",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class TransparentColorFeature
    {
        public static bool Execute(PowerPointManager powerPointManager)
        {
            var win = new TransparentColorWindow(powerPointManager);
            win.ShowDialog();
            return true;
        }
    }
}
