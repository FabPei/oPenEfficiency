using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Shapes
{
    [FeatureMetadata(
        Id = "BtnOptimizeFreeForm",
        Name = "Optimize Free Form",
        Tooltip = "Optimize Free Form Path",
        IconData = "M20.71,7.04C21.1,6.65 21.1,6 20.71,5.63L18.37,3.29C18,2.9 17.35,2.9 16.96,3.29L15.12,5.12L18.87,8.87M3,17.25V21H6.75L17.81,9.93L14.06,6.18L3,17.25Z", // Pen edit icon
        Color = "#FBBF24",
        MinSelection = 1,
        RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes,
        Description = "Traverses the path vertices of hand-drawn/freeform shapes and removes mathematical outliers and redundant points to smooth the path globally.",
        Keywords = "simplify path, reduce vertices, smooth freehand, clean up svg",
        DetailedHelpText = "### Optimize Free Form\n\nReduces the number of vertex nodes in a freehand vector path without visually distorting the shape. Useful for cleaning up imported SVGs or hand-drawn paths."
    )]
    public static class OptimizeFreeFormFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var selection = manager.GetApplication().ActiveWindow.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return false;

                bool modified = false;

                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    modified |= OptimizeShape(shape);
                }

                return modified;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex.Message, "OptimizeFreeFormFeature");
                return false;
            }
        }

        private static bool OptimizeShape(PowerPoint.Shape shape)
        {
            if (shape.Type == MsoShapeType.msoGroup)
            {
                bool grpModified = false;
                foreach (PowerPoint.Shape child in shape.GroupItems)
                {
                    grpModified |= OptimizeShape(child);
                }
                return grpModified;
            }

            if (shape.Type != MsoShapeType.msoFreeform || shape.Nodes.Count <= 3)
                return false;

            int initialNodes = shape.Nodes.Count;
            // Iterate backwards to safely delete nodes
            for (int i = shape.Nodes.Count - 1; i >= 2; i--)
            {
                try
                {
                    var p1 = (Array)shape.Nodes[i].Points;
                    var p2 = (Array)shape.Nodes[i - 1].Points;

                    // Depending on version, it might be float[,] or object[,]
                    float x1 = Convert.ToSingle(p1.GetValue(1, 1));
                    float y1 = Convert.ToSingle(p1.GetValue(1, 2));

                    float x2 = Convert.ToSingle(p2.GetValue(1, 1));
                    float y2 = Convert.ToSingle(p2.GetValue(1, 2));

                    float distSq = (x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2);

                    // Threshold: Delete if distance between points is less than ~5 pts
                    if (distSq < 25)
                    {
                        shape.Nodes.Delete(i);
                    }
                }
                catch
                {
                    // Ignore node errors (could be a curve node or path sub-break)
                }
            }

            return initialNodes != shape.Nodes.Count;
        }
    }
}
