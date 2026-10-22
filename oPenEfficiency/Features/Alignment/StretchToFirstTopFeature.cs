using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stretch the height of selected shapes UP so their Top touches the Bottom of the first object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStretchTop",
        Name = "Stretch Top",
        Tooltip = "Stretch top",
        IconData = "M13,22V10H16L12,6L8,10H11V22H13M22,2V4H2V2H22Z",
        Color = "#6366F1",
        Description = "Extends shapes upward across the gap until they touch the anchor's bottom edge.",
        DetailedHelpText = "### Stretch Top\n\nKeeps the first or last selected anchor configured in Settings fixed. Shapes entirely below it extend their top edge to the anchor's bottom edge while their original bottom edge, width, horizontal position, and aspect-ratio lock state remain unchanged. Overlapping shapes and shapes on the other side are skipped. Right-click any Stretch button to choose a gap direction or move a specific edge to the anchor.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class StretchToFirstTopFeature
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
                float refBottom = refShape.Top + refShape.Height;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (AlignmentHelpers.IsStretchTargetEligible("BtnStretchTop", shape, refShape))
                        originalStates.Add(System.Tuple.Create(shape, shape.Left, shape.Top, shape.Width, shape.Height, shape.LockAspectRatio));
                }

                if (originalStates.Count == 0) return false;
                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (System.Exception ex) { ExceptionLogger.Log(ex, "StretchToFirstTopFeature.StartNewUndoEntry"); }

                foreach (var state in originalStates)
                {
                    var shape = state.Item1;
                    var originalLock = state.Item6;
                    try
                    {
                        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        float originalBottom = state.Item3 + state.Item5;
                        shape.Top = refBottom;
                        shape.Height = originalBottom - refBottom;
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
                    catch (System.Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "StretchToFirstTopFeature.Rollback"); }
                }
                ExceptionLogger.Log(ex, "StretchToFirstTopFeature.Execute");
                return false;
            }
        }
    }
}
