using System;
using Microsoft.Office.Interop.PowerPoint;
using System.Globalization;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Copy X/Y coordinates of the single selected shape to the clipboard.
    /// Only works when exactly one shape is selected.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnCopyXY",
        Name = "Copy XY-Coordinates",
        Tooltip = "Copy XY-Coordinates",
        IconData = "M16 1H4C2.9 1 2 1.9 2 3v14h2V3h12V1zM19 5H8C6.9 5 6 5.9 6 7v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z",
        Color = "#F59E0B",
        Description = "Grabs the exact top-left coordinates and size of the selected object. Use 'Paste XY' to apply them to another object.",
        DetailedHelpText = "### Copy XY-Coordinates\nCopies the exact X position, Y position, Width, and Height of the selected shape to an internal clipboard. Use Paste XY-Coordinates to apply these dimensions to other shapes.",
        MinSelection = 1,
        MaxSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class CopyCoordinatesFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                var selection = app.ActiveWindow.Selection;
                if (selection == null || selection.Type != PpSelectionType.ppSelectionShapes) return false;
                var shapes = selection.ShapeRange;
                if (shapes.Count != 1) return false;

                float x = shapes[1].Left;
                float y = shapes[1].Top;
                float width = shapes[1].Width;
                float height = shapes[1].Height;

                // Use InvariantCulture to ensure dot as decimal separator
                string coords = string.Format(CultureInfo.InvariantCulture, "{0},{1},{2},{3}", x, y, width, height);

                System.Windows.Forms.Clipboard.SetText(coords);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CopyCoordinatesFeature.Execute");
                return false;
            }
        }
    }
}
