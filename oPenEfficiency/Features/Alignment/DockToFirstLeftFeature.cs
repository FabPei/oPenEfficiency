using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Dock selected shapes to the left of the first selected shape.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnDockLeft",
        Name = "Stack Left",
        Tooltip = "Stack shapes to the left of the anchor object",
        SearchAliases = new[] { "Dock Left" },
        IconData = "M10,4V20H6V4H10M12,2H4V22H12V2M14,4H20V20H14V4Z",
        Color = "#F59E0B",
        Description = "Moves selected shapes to the left of the anchor object, stacking them horizontally.",
        DetailedHelpText = "### Stack Left\n\nKeeps the first or last selected anchor object configured in Settings fixed, then places every other selected shape directly to its left in selection order with zero horizontal gap. Right-click any Stack button to choose left, right, above, or below directly. Vertical positions, sizes, and formatting remain unchanged.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class DockToFirstLeftFeature
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
                catch (System.Exception ex) { ExceptionLogger.Log(ex, "DockToFirstLeftFeature.StartNewUndoEntry"); }

                var orderedSizes = new System.Collections.Generic.List<float>();
                var orderedShapes = new System.Collections.Generic.List<Shape>();
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (shape.Id == refShape.Id) continue;
                    orderedShapes.Add(shape);
                    orderedSizes.Add(shape.Width);
                }
                var positions = oPenEfficiency.Core.Services.DockStackPolicy.StackSequence(
                    oPenEfficiency.Core.Services.DockStackPolicy.DockDirection.Left,
                    refShape.Left, refShape.Top, refShape.Width, refShape.Height,
                    orderedSizes.ToArray());
                for (int i = 0; i < orderedShapes.Count; i++)
                {
                    orderedShapes[i].Left = positions[i];
                }
                return true;
            }
            catch (System.Exception ex)
            {
                foreach (var original in originalPositions)
                {
                    try { original.Item1.Left = original.Item2; original.Item1.Top = original.Item3; }
                    catch (System.Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "DockToFirstLeftFeature.Rollback"); }
                }
                ExceptionLogger.Log(ex, "DockToFirstLeftFeature.Execute");
                return false;
            }
        }
    }
}
