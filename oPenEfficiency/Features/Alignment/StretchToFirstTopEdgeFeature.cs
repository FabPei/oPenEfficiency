using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stretch selected shapes so their Top edge matches the first object's Top edge.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStretchTopEdge",
        Name = "Stretch Top Edge",
        Tooltip = "Stretch top edge",
        IconData = "M2,20V22H22V20H2M12,11L8,15H11V22H13V15H16L12,11Z",
        Color = "#6366F1",
        Description = "Resizes selected shapes so their top edges match the anchor's top edge.",
        DetailedHelpText = "### Stretch Top Edge\n\nKeeps the first or last selected anchor configured in Settings fixed. Each other shape whose bottom edge lies below the anchor's top edge is resized from the top until both top edges match; its original bottom edge, width, horizontal position, and aspect-ratio lock state remain unchanged. Shapes entirely above the target edge are skipped to prevent inversion. Right-click any Stretch button to choose a gap direction or move a specific edge to the anchor.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class StretchToFirstTopEdgeFeature
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
                float targetTop = refShape.Top;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (AlignmentHelpers.IsStretchTargetEligible("BtnStretchTopEdge", shape, refShape))
                        originalStates.Add(System.Tuple.Create(shape, shape.Left, shape.Top, shape.Width, shape.Height, shape.LockAspectRatio));
                }

                if (originalStates.Count == 0) return false;
                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (System.Exception ex) { ExceptionLogger.Log(ex, "StretchToFirstTopEdgeFeature.StartNewUndoEntry"); }

                foreach (var state in originalStates)
                {
                    var shape = state.Item1;
                    var originalLock = state.Item6;
                    try
                    {
                        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        float currentBottom = state.Item3 + state.Item5;
                        shape.Top = targetTop;
                        shape.Height = currentBottom - targetTop;
                    }
                    finally
                    {
                        shape.LockAspectRatio = originalLock;
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
                    catch (System.Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "StretchToFirstTopEdgeFeature.Rollback"); }
                }
                ExceptionLogger.Log(ex, "StretchToFirstTopEdgeFeature.Execute");
                return false;
            }
        }
    }
}
