using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnBringToFront",
        Name = "Bring to Front",
        Tooltip = "Bring selection to the very front",
        IconData = "M14 2h1a2 2 0 0 1 2 2v1 M7 2H6a2 2 0 0 0-2 2v1 M2 7V6a2 2 0 0 1 2-2h1 M22 7V6a2 2 0 0 0-2-2h-1 M14 22h1a2 2 0 0 0 2-2v-1 M7 22H6a2 2 0 0 1-2-2v-1 M2 17v1a2 2 0 0 0 2 2h1 M22 17v1a2 2 0 0 1-2 2h-1 M8 8h8v8h-8z",
        Color = "#10B981",
        Description = "Brings the selected shape to the absolute front layer.",
        MinSelection = 1)]
    public static class BringToFrontFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;

            try
            {
                shapes.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                return true;
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "BringToFrontFeature.Execute");
                return false;
            }
        }

        public static bool ExecuteRanked(PowerPointManager manager)
        {
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;

            try
            {
                // To have 1st selected on top, 2nd below it, etc.
                // We move them to front in REVERSE selection order
                for (int i = shapes.Count; i >= 1; i--)
                {
                    shapes[i].ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                }
                return true;
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "BringToFrontFeature.ExecuteRanked");
                return false;
            }
        }
    }

    [FeatureMetadata(
        Id = "BtnSendToBack",
        Name = "Send to Back",
        Tooltip = "Send selection to the very back",
        IconData = "M7 2h1a2 2 0 0 1 2 2v1 M2 7V6a2 2 0 0 1 2-2h1 M22 14v1a2 2 0 0 1-2 2h-1 M14 22h1a2 2 0 0 0 2-2v-1 M14 14h8v8h-8z M2 2h8v8h-8z",
        Color = "#10B981",
        Description = "Sends the selected shape to the absolute back layer.",
        MinSelection = 1)]
    public static class SendToBackFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;

            try
            {
                shapes.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                return true;
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "SendToBackFeature.Execute");
                return false;
            }
        }

        public static bool ExecuteRanked(PowerPointManager manager)
        {
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;

            try
            {
                // To have 1st selected furthest back, 2nd above it, etc.
                // We send them to back in REVERSE selection order
                // Logic: 3rd sent to back, then 2nd sent to back (pushes 3rd forward), then 1st sent to back (pushes 2nd and 3rd forward)
                // Result: 1st is back, 2nd is above it, 3rd is above that.
                for (int i = shapes.Count; i >= 1; i--)
                {
                    shapes[i].ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                }
                return true;
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "SendToBackFeature.ExecuteRanked");
                return false;
            }
        }
    }

    [FeatureMetadata(
        Id = "BtnBringForward",
        Name = "Bring Forward",
        Tooltip = "Bring selection one step forward",
        IconData = "M6 2h12a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z M12 14V6 M9 9l3-3 3 3",
        Color = "#10B981",
        Description = "Brings the selected shape forward one layer.",
        MinSelection = 1)]
    public static class BringForwardFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;

            try
            {
                shapes.ZOrder(Office.MsoZOrderCmd.msoBringForward);
                return true;
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "BringForwardFeature.Execute");
                return false;
            }
        }

        public static bool ExecuteRanked(PowerPointManager manager)
        {
            // For Bring Forward / Send Backward, "Ranked" usually means 
            // reordering the selection relative to each other based on selection rank
            // without necessarily moving to front/back of the entire slide.
            
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count < 2) return Execute(manager);

            try
            {
                // Get the current Z-positions of all shapes in selection
                var list = new List<Shape>();
                for (int i = 1; i <= shapes.Count; i++) list.Add(shapes[i]);
                
                var positions = list.Select(s => s.ZOrderPosition).OrderBy(p => p).ToList();
                
                // Assign these positions to shapes in order of selection (reverse)
                // First selected gets the highest position, last gets lowest
                var selectionInOrder = list; // shapes[1] is 1st selected
                
                // We can't set ZOrderPosition directly, so we use the normalization trick
                // Move them all to front in reverse selection order ensures they are ranked correctly
                // but this moves them to the VERY front.
                
                // If we want to maintain the "stack" location but just reorder them:
                // 1. Move them to front in reverse order
                // 2. Move the whole range back until the lowest one is at the original lowest position
                int originalLowest = positions[0];
                
                for (int i = selectionInOrder.Count; i >= 1; i--)
                {
                    selectionInOrder[i-1].ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                }
                
                var updatedShapes = manager.GetSelectedShapes();
                while (updatedShapes.ZOrderPosition > originalLowest)
                {
                    updatedShapes.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                }
                
                return true;
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "BringForwardFeature.ExecuteRanked");
                return false;
            }
        }
    }

    [FeatureMetadata(
        Id = "BtnSendBackward",
        Name = "Send Backward",
        Tooltip = "Send selection one step backward",
        IconData = "M6 2h12a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z M12 6v8 M9 11l3 3 3-3",
        Color = "#10B981",
        Description = "Sends the selected shape backward one layer.",
        MinSelection = 1)]
    public static class SendBackwardFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;

            try
            {
                shapes.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                return true;
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "SendBackwardFeature.Execute");
                return false;
            }
        }

        public static bool ExecuteRanked(PowerPointManager manager)
        {
            // Ranked Send Backward: 1st selected is furthest back among selection
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count < 2) return Execute(manager);

            try
            {
                var list = new List<Shape>();
                for (int i = 1; i <= shapes.Count; i++) list.Add(shapes[i]);
                var positions = list.Select(s => s.ZOrderPosition).OrderBy(p => p).ToList();
                int originalLowest = positions[0];

                // Normalize: 1st selected bottom, last selected top
                for (int i = list.Count; i >= 1; i--)
                {
                    list[i-1].ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                }

                var updatedShapes = manager.GetSelectedShapes();
                while (updatedShapes.ZOrderPosition < originalLowest)
                {
                    updatedShapes.ZOrder(Office.MsoZOrderCmd.msoBringForward);
                }

                return true;
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "SendBackwardFeature.ExecuteRanked");
                return false;
            }
        }
    }
}
