using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Services;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Pastes X and Y coordinates from the clipboard and applies them to the selected shape.
    /// Expects clipboard text in format "x,y".
    /// </summary>
    [FeatureMetadata(
        Id = "BtnPasteXY",
        Name = "Paste XY-Coordinates",
        SearchAliases = new[] { "Paste Coordinates", "Paste Position" },
        Tooltip = "Apply copied X/Y position to the selected shape",
        Color = "#F59E0B",
        Description = "Applies a copied top-left X/Y position to one selected shape without changing its size.",
        DetailedHelpText = "### Paste XY-Coordinates\nSelect exactly one target shape, then paste the Left and Top position copied by Copy XY-Coordinates. Width, Height, rotation, text, and formatting stay unchanged. Current payloads are versioned and openEfficiency-owned; legacy two- and four-value X/Y payloads remain paste-compatible, with four-value Width and Height ignored. Right-click Paste XY-Coordinates clears only a valid current openEfficiency payload, never unrelated numeric clipboard text. Use Copy/Paste Position & Size when you need to transfer dimensions.",
        MinSelection = 1,
        MaxSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class PasteCoordinatesFeature
    {
        // Required zero-arg overload for FeatureDiscovery auto-registration.
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

                string clipText = System.Windows.Forms.Clipboard.GetText();
                if (string.IsNullOrEmpty(clipText)) return false;

                if (!CoordinateClipboardPayload.TryParseForPaste(clipText, out float x, out float y)) return false;

                shapes[1].Left = x;
                shapes[1].Top = y;
                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "PasteCoordinatesFeature.Execute");
                return false;
            }
        }

        public static bool ClearCopiedGeometry()
        {
            return CopyCoordinatesFeature.ClearCopiedGeometry();
        }
    }
}
