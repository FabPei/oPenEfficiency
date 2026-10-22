using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stretch the width of selected shapes LEFT so their Left touches the Right of the first object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStretchLeft",
        Name = "Stretch Left",
        Tooltip = "Stretch left",
        IconData = "M22,13H10V16L6,12L10,8V11H22V13M2,2H4V22H2V2Z",
        Color = "#6366F1",
        Description = "Extends shapes leftward across the gap until they touch the anchor's right edge.",
        DetailedHelpText = "### Stretch Left\n\nKeeps the first or last selected anchor configured in Settings fixed. Shapes entirely to its right extend their left edge to the anchor's right edge while their original right edge, height, position on the other axis, and aspect-ratio lock state remain unchanged. Overlapping shapes and shapes on the other side are skipped. Right-click any Stretch button to choose a gap direction or move a specific edge to the anchor.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class StretchToFirstLeftFeature
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
                float refRight = refShape.Left + refShape.Width;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (AlignmentHelpers.IsStretchTargetEligible("BtnStretchLeft", shape, refShape))
                        originalStates.Add(System.Tuple.Create(shape, shape.Left, shape.Top, shape.Width, shape.Height, shape.LockAspectRatio));
                }

                if (originalStates.Count == 0) return false;
                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (System.Exception ex) { ExceptionLogger.Log(ex, "StretchToFirstLeftFeature.StartNewUndoEntry"); }

                foreach (var state in originalStates)
                {
                    var shape = state.Item1;
                    var originalLock = state.Item6;
                    try
                    {
                        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        float originalRight = state.Item2 + state.Item4;
                        shape.Left = refRight;
                        shape.Width = originalRight - refRight;
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
                    catch (System.Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "StretchToFirstLeftFeature.Rollback"); }
                }
                ExceptionLogger.Log(ex, "StretchToFirstLeftFeature.Execute");
                return false;
            }
        }
    }
}
