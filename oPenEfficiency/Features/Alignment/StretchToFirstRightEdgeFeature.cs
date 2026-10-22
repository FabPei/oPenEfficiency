using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stretch selected shapes so their Right edge matches the first object's Right edge.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStretchRightEdge",
        Name = "Stretch Right Edge",
        Tooltip = "Stretch right edge",
        IconData = "M2,2H4V22H2V2M13,12L9,8V11H22V13H9V16L13,12Z",
        Color = "#6366F1",
        Description = "Resizes selected shapes so their right edges match the configured anchor's right edge.",
        DetailedHelpText = "### Stretch Right Edge\n\nKeeps the first or last selected anchor configured in Settings fixed. Each eligible shape keeps its left edge fixed and stretches its right edge to the anchor's right edge. Shapes whose left edge is at or beyond the target, and shapes already aligned to it, are skipped to avoid inversion and no-op changes. Right-click any Stretch button to choose a gap direction or move a specific edge to the anchor.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class StretchToFirstRightEdgeFeature
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
                float targetRight = refShape.Left + refShape.Width;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (AlignmentHelpers.IsStretchTargetEligible("BtnStretchRightEdge", shape, refShape))
                        originalStates.Add(System.Tuple.Create(shape, shape.Left, shape.Top, shape.Width, shape.Height, shape.LockAspectRatio));
                }

                if (originalStates.Count == 0) return false;
                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (System.Exception ex) { ExceptionLogger.Log(ex, "StretchToFirstRightEdgeFeature.StartNewUndoEntry"); }

                foreach (var state in originalStates)
                {
                    try
                    {
                        state.Item1.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        state.Item1.Width = targetRight - state.Item2;
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
                    catch (System.Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "StretchToFirstRightEdgeFeature.Rollback"); }
                }
                ExceptionLogger.Log(ex, "StretchToFirstRightEdgeFeature.Execute");
                return false;
            }
        }
    }
}
