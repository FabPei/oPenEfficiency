using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Visuals
{
    [FeatureMetadata(Id = "BtnFlipHorizontal", Name = "Flip Horizontal", Tooltip = "Flip Horizontal", IconData = "M15 21l6-9l-6-9v18M9 21L3 12l6-9v18M12 2v20", Color = "#D946EF", Description = "Flips the selected shapes horizontally.", Keywords = "flip, horizontal, mirror, reverse", MinSelection = 1, RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class FlipHorizontalFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try 
            { 
                var app = manager.GetApplication();
                if (app != null) app.CommandBars.ExecuteMso("ObjectsFlipHorizontal"); 
                return true; 
            } 
            catch 
            { 
                // Fallback
                var shapes = manager.GetSelectedShapes();
                if (shapes == null || shapes.Count == 0) return false;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    shapes[i].Flip(Office.MsoFlipCmd.msoFlipHorizontal);
                }
                return true;
            }
        }
    }

    [FeatureMetadata(Id = "BtnFlipVertical", Name = "Flip Vertical", Tooltip = "Flip Vertical", IconData = "M21 9l-9-6l-9 6h18M21 15l-9 6l-9-6h18M2 12h20", Color = "#D946EF", Description = "Flips the selected shapes vertically.", Keywords = "flip, vertical, mirror, reverse", MinSelection = 1, RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class FlipVerticalFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try 
            { 
                var app = manager.GetApplication();
                if (app != null) app.CommandBars.ExecuteMso("ObjectsFlipVertical"); 
                return true; 
            } 
            catch 
            { 
                // Fallback
                var shapes = manager.GetSelectedShapes();
                if (shapes == null || shapes.Count == 0) return false;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    shapes[i].Flip(Office.MsoFlipCmd.msoFlipVertical);
                }
                return true;
            }
        }
    }

    [FeatureMetadata(Id = "BtnRotateShape", Name = "Rotate", Tooltip = "Rotate (Right-click for options)", IconData = "M12 2v3a7 7 0 0 1 7 7h3a10 10 0 0 0-10-10m7 10a7 7 0 0 1-7 7v3a10 10 0 0 0 10-10h-3M12 22v-3a7 7 0 0 1-7-7H2a10 10 0 0 0 10 10M5 12a7 7 0 0 1 7-7V2A10 10 0 0 0 2 12h3M12 2l-3 4h6l-3-4M22 12l-4-3v6l4-3M12 22l3-4H9l3 4M2 12l4 3V9l-4 3z", Color = "#D946EF", Description = "Rotates the selected shapes. Left-click to rotate 90 degrees right; right-click for custom angles.", Keywords = "rotate, turn, spin, angle", MinSelection = 1, RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class RotateShapeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            // Left click is 90 right
            return ExecuteSpecific(manager, 90f);
        }

        public static bool ExecuteSpecific(PowerPointManager manager, float degrees)
        {
            if (manager == null) return false;
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;

            try
            {
                for (int i = 1; i <= shapes.Count; i++)
                {
                    shapes[i].IncrementRotation(degrees);
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "RotateShapeFeature.ExecuteSpecific");
                return false;
            }
        }
    }
}

