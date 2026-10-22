using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stretch selected shapes so their Left edge matches the first object's Left edge.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStretchLeftEdge",
        Name = "Stretch Left Edge",
        Tooltip = "Stretch left edge",
        IconData = "M20,2H22V22H20V2M11,12L15,8V11H2V13H15V16L11,12Z",
        Color = "#6366F1",
        Description = "Resizes selected shapes so their left edges match the anchor's left edge.",
        DetailedHelpText = "### Stretch Left Edge\n\nKeeps the first or last selected anchor configured in Settings fixed. Each other shape whose right edge lies beyond the anchor's left edge is resized from the left until both left edges match; its original right edge, height, vertical position, and aspect-ratio lock state remain unchanged. Shapes entirely left of the target edge are skipped to prevent inversion. Right-click any Stretch button to choose a gap direction or move a specific edge to the anchor.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class StretchToFirstLeftEdgeFeature
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
                float targetLeft = refShape.Left;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (AlignmentHelpers.IsStretchTargetEligible("BtnStretchLeftEdge", shape, refShape))
                        originalStates.Add(System.Tuple.Create(shape, shape.Left, shape.Top, shape.Width, shape.Height, shape.LockAspectRatio));
                }

                if (originalStates.Count == 0) return false;
                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (System.Exception ex) { ExceptionLogger.Log(ex, "StretchToFirstLeftEdgeFeature.StartNewUndoEntry"); }

                foreach (var state in originalStates)
                {
                    var shape = state.Item1;
                    var originalLock = state.Item6;
                    try
                    {
                        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        float currentRight = state.Item2 + state.Item4;
                        shape.Left = targetLeft;
                        shape.Width = currentRight - targetLeft;
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
                    catch (System.Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "StretchToFirstLeftEdgeFeature.Rollback"); }
                }
                ExceptionLogger.Log(ex, "StretchToFirstLeftEdgeFeature.Execute");
                return false;
            }
        }
    }
}
