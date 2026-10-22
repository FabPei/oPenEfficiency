using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Services;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Copy X/Y coordinates of the single selected shape to the clipboard.
    /// Only works when exactly one shape is selected.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnCopyXY",
        Name = "Copy XY-Coordinates",
        SearchAliases = new[] { "Copy Coordinates", "Copy Position" },
        Tooltip = "Copy the selected shape's X/Y position",
        IconData = "M16 1H4C2.9 1 2 1.9 2 3v14h2V3h12V1zM19 5H8C6.9 5 6 5.9 6 7v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z",
        Color = "#F59E0B",
        Description = "Copies the exact top-left X/Y position from one selected shape for Paste XY-Coordinates.",
        DetailedHelpText = "### Copy XY-Coordinates\nSelect exactly one shape, then copy its Left and Top position as a versioned openEfficiency X/Y payload for Paste XY-Coordinates. This command never copies Width or Height. Paste XY-Coordinates also accepts legacy two- or four-value X/Y payloads, but right-click Copy XY-Coordinates clears only a valid current openEfficiency payload and never unrelated numeric clipboard text. Use Copy/Paste Position & Size when you need to transfer dimensions.",
        MinSelection = 1,
        MaxSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
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
                if (!CoordinateClipboardPayload.TrySerializeCurrent(x, y, out string coords)) return false;

                System.Windows.Forms.Clipboard.SetText(coords);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CopyCoordinatesFeature.Execute");
                return false;
            }
        }

        public static bool ClearCopiedGeometry()
        {
            try
            {
                string clipboardText = System.Windows.Forms.Clipboard.GetText();
                if (!CoordinateClipboardPayload.IsOwnedCurrentPayload(clipboardText)) return false;

                System.Windows.Forms.Clipboard.Clear();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CopyCoordinatesFeature.ClearCopiedGeometry");
                return false;
            }
        }
    }
}
