using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stack selected shapes on top of the first object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnDockTop",
        Name = "Stack Above",
        Tooltip = "Stack above",
        SearchAliases = new[] { "Dock Top", "Stack Top" },
        IconData = "M4,10V6H20V10H4M2,12V4H22V12H2M4,14V20H20V14H4Z",
        Color = "#F59E0B",
        Description = "Moves selected shapes above the anchor shape, stacking them vertically.",
        DetailedHelpText = "### Stack Above\n\nKeeps the first or last selected anchor object configured in Settings fixed, then places every other selected shape directly above it in selection order with zero vertical gap. Right-click any Stack button to choose left, right, above, or below directly. Horizontal positions, sizes, and formatting remain unchanged.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class DockToFirstTopFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var originalPositions = new System.Collections.Generic.List<System.Tuple<Shape, float, float>>();
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var refShape = manager.GetReferenceShape();
                if (refShape == null) return false;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (shape.Id != refShape.Id)
                        originalPositions.Add(System.Tuple.Create(shape, shape.Left, shape.Top));
                }

                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (System.Exception ex) { ExceptionLogger.Log(ex, "DockToFirstTopFeature.StartNewUndoEntry"); }

                var orderedSizes = new System.Collections.Generic.List<float>();
                var orderedShapes = new System.Collections.Generic.List<Shape>();
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (shape.Id == refShape.Id) continue;
                    orderedShapes.Add(shape);
                    orderedSizes.Add(shape.Height);
                }
                var positions = oPenEfficiency.Core.Services.DockStackPolicy.StackSequence(
                    oPenEfficiency.Core.Services.DockStackPolicy.DockDirection.Above,
                    refShape.Left, refShape.Top, refShape.Width, refShape.Height,
                    orderedSizes.ToArray());
                for (int i = 0; i < orderedShapes.Count; i++)
                {
                    orderedShapes[i].Top = positions[i];
                }
                return true;
            }
            catch (System.Exception ex)
            {
                foreach (var original in originalPositions)
                {
                    try { original.Item1.Left = original.Item2; original.Item1.Top = original.Item3; }
                    catch (System.Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "DockToFirstTopFeature.Rollback"); }
                }
                ExceptionLogger.Log(ex, "DockToFirstTopFeature.Execute");
                return false;
            }
        }
    }
}
