using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Shapes
{
    [FeatureMetadata(
        Id = "BtnOptimizeFreeForm",
        Name = "Simplify Freeform Nodes",
        SearchAliases = new[] { "Optimize Free Form" },
        Tooltip = "Simplify freeform paths by removing nearby nodes",
        IconData = "M20.71,7.04C21.1,6.65 21.1,6 20.71,5.63L18.37,3.29C18,2.9 17.35,2.9 16.96,3.29L15.12,5.12L18.87,8.87M3,17.25V21H6.75L17.81,9.93L14.06,6.18L3,17.25Z", // Pen edit icon
        Color = "#FBBF24",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes,
        Description = "Removes adjacent freeform nodes that are less than five points apart.",
        DetailedHelpText = "### Simplify Freeform Nodes\n\nCleans selected freeform shapes, including freeforms inside selected groups, by removing adjacent nodes whose distance is less than 5 pt. Endpoints are retained and the shape object is edited in place. Curve or subpath nodes that PowerPoint does not allow to be deleted are skipped and logged. The command does not perform global smoothing or mathematical outlier detection."
    )]
    public static class OptimizeFreeFormFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, 5f);
        }

        public static bool Execute(PowerPointManager manager, float minimumDistance)
        {
            if (manager == null || minimumDistance <= 0) return false;
            try
            {
                var application = manager.GetApplication();
                var selection = application?.ActiveWindow?.Selection;
                if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                    return false;

                try { application.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "OptimizeFreeFormFeature.StartNewUndoEntry"); }

                bool modified = false;
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    try
                    {
                        modified |= OptimizeShape(shape, minimumDistance);
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "OptimizeFreeFormFeature.OptimizeShape");
                    }
                }

                return modified;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "OptimizeFreeFormFeature.Execute");
                return false;
            }
        }

        private static bool OptimizeShape(PowerPoint.Shape shape, float minimumDistance)
        {
            if (shape.Type == MsoShapeType.msoGroup)
            {
                bool grpModified = false;
                foreach (PowerPoint.Shape child in shape.GroupItems)
                {
                    grpModified |= OptimizeShape(child, minimumDistance);
                }
                return grpModified;
            }

            if (shape.Type != MsoShapeType.msoFreeform || shape.Nodes.Count <= 3)
                return false;

            int initialNodes = shape.Nodes.Count;

            // Read all node points up front; the deletion set is computed host-free from the
            // original adjacent pairs (deletions at descending indices never shift lower ones).
            var points = new List<(float X, float Y)>(initialNodes);
            for (int n = 1; n <= initialNodes; n++)
            {
                var p = (Array)shape.Nodes[n].Points;
                points.Add((Convert.ToSingle(p.GetValue(1, 1)), Convert.ToSingle(p.GetValue(1, 2))));
            }

            foreach (int nodeIndex in oPenEfficiency.Core.Services.FreeformDecimationPolicy.ComputeDeletableNodeIndices(points, minimumDistance))
            {
                try
                {
                    shape.Nodes.Delete(nodeIndex);
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "OptimizeFreeFormFeature.OptimizeNode");
                }
            }

            return initialNodes != shape.Nodes.Count;
        }
    }
}
