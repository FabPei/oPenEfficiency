using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI;
using oPenEfficiency.Utils;
using System;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnTransparentColor",
        Name = "Color Eraser",
        Tooltip = "Erase a picked color from one selected shape",
        SearchAliases = new[] { "Magic Color Erase" },
        IconData = "M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12A10,10 0 0,0 12,2M12,4A8,8 0 0,1 20,12A8,8 0 0,1 12,20A8,8 0 0,1 4,12A8,8 0 0,1 12,4M12,6A6,6 0 0,0 6,12A6,6 0 0,0 12,18A6,6 0 0,0 18,12A6,6 0 0,0 12,6Z",
        Color = "#06B6D4",
        Description = "Rasterizes one selected shape and makes pixels matching a picked screen color transparent.",
        DetailedHelpText = "### Color Eraser\nSelect exactly one shape, pick a color from the screen, adjust the tolerance, and erase matching pixels. The selected shape is exported and replaced with a PNG, so editable text, vector geometry, animations, links, and other shape-specific properties are flattened.",
        MinSelection = 1,
        MaxSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class TransparentColorFeature
    {
        public static bool Execute(PowerPointManager powerPointManager)
        {
            return Execute(powerPointManager, 10);
        }

        public static bool Execute(PowerPointManager powerPointManager, double initialTolerance)
        {
            if (powerPointManager == null) return false;

            try
            {
                var selectedShapes = powerPointManager.GetSelectedShapes();
                if (selectedShapes == null || selectedShapes.Count != 1) return false;

                var win = new TransparentColorWindow(powerPointManager, initialTolerance);
                win.ShowDialog();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TransparentColorFeature.Execute");
                throw;
            }
        }
    }
}
