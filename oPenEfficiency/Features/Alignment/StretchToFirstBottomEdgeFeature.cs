using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stretch selected shapes so their Bottom edge matches the first object's Bottom edge.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStretchBottomEdge",
        Name = "Stretch Bottom Edge",
        Tooltip = "Stretch bottom edge",
        IconData = "M2,2V4H22V2H2M12,13L8,9H11V2H13V9H16L12,13Z",
        Color = "#6366F1",
        Description = "Resizes selected shapes so their bottom edges match the configured anchor's bottom edge.",
        DetailedHelpText = "### Stretch Bottom Edge\n\nKeeps the first or last selected anchor configured in Settings fixed. Each eligible shape keeps its top edge fixed and stretches its bottom edge to the anchor's bottom edge. Shapes whose top edge is at or beyond the target, and shapes already aligned to it, are skipped to avoid inversion and no-op changes. Right-click any Stretch button to choose a gap direction or move a specific edge to the anchor.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class StretchToFirstBottomEdgeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var originalStates = new System.Collections.Generic.List<System.Tuple<Shape, float, float, float, float, Microsoft.Office.Core.MsoTriState>>();
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var refShape = manager.GetReferenceShape();
                if (refShape == null) return false;
                float targetBottom = refShape.Top + refShape.Height;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (AlignmentHelpers.IsStretchTargetEligible("BtnStretchBottomEdge", shape, refShape))
                        originalStates.Add(System.Tuple.Create(shape, shape.Left, shape.Top, shape.Width, shape.Height, shape.LockAspectRatio));
                }

                if (originalStates.Count == 0) return false;
                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (System.Exception ex) { ExceptionLogger.Log(ex, "StretchToFirstBottomEdgeFeature.StartNewUndoEntry"); }

                foreach (var state in originalStates)
                {
                    try
                    {
                        state.Item1.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        state.Item1.Height = targetBottom - state.Item3;
                    }
                    finally
                    {
                        state.Item1.LockAspectRatio = state.Item6;
                    }
                }
                return true;
            }
            catch (System.Exception ex)
            {
                foreach (var state in originalStates)
                {
                    try
                    {
                        state.Item1.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        state.Item1.Width = state.Item4;
                        state.Item1.Height = state.Item5;
                        state.Item1.Left = state.Item2;
                        state.Item1.Top = state.Item3;
                        state.Item1.LockAspectRatio = state.Item6;
                    }
                    catch (System.Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "StretchToFirstBottomEdgeFeature.Rollback"); }
                }
                ExceptionLogger.Log(ex, "StretchToFirstBottomEdgeFeature.Execute");
                return false;
            }
        }
    }
}
