using Microsoft.Office.Interop.PowerPoint;
using System.Globalization;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Pastes X and Y coordinates from the clipboard and applies them to the selected shape.
    /// Expects clipboard text in format "x,y".
    /// </summary>
    [FeatureMetadata(
        Id = "BtnPasteXY",
        Name = "Paste XY-Coordinates",
        Tooltip = "Paste XY-Coordinates",
        Color = "#F59E0B",
        Description = "Applies copied coordinates or size to the selected shape. Right-click for more options.",
        DetailedHelpText = "### Paste XY-Coordinates\nApplies geometric properties from the clipboard to the selection.\n\n**Usage:**\n1. Copy coordinates from a source shape using 'Copy XY'.\n2. Select target shapes.\n3. Click to paste both Position and Size.\n\n**Right-Click Options:**\n* **Paste Coordinates**: Apply X/Y position only.\n* **Paste Size**: Apply Width/Height only.",
        MinSelection = 1,
        MaxSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class PasteCoordinatesFeature
    {
        public enum PasteMode { Coordinates, Size, Both }

        // Required zero-arg overload for FeatureDiscovery auto-registration.
        public static bool Execute(PowerPointManager manager)
            => Execute(manager, PasteMode.Coordinates);

        public static bool Execute(PowerPointManager manager, PasteMode mode = PasteMode.Coordinates)
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

                string clipText = System.Windows.Forms.Clipboard.GetText();
                if (string.IsNullOrEmpty(clipText)) return false;

                // Split only by comma, assuming the string was created by CopyCoordinates
                var parts = clipText.Split(',');
                if (parts.Length < 2) return false;

                bool success = false;

                if ((mode == PasteMode.Coordinates || mode == PasteMode.Both) && parts.Length >= 2)
                {
                    if (float.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out float x) &&
                        float.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out float y))
                    {
                        shapes[1].Left = x;
                        shapes[1].Top = y;
                        success = true;
                    }
                }

                if ((mode == PasteMode.Size || mode == PasteMode.Both) && parts.Length >= 4)
                {
                    if (float.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out float width) &&
                        float.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out float height))
                    {
                        shapes[1].Width = width;
                        shapes[1].Height = height;
                        success = true;
                    }
                }

                return success;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "PasteCoordinatesFeature.Execute");
                return false;
            }
        }
    }
}
