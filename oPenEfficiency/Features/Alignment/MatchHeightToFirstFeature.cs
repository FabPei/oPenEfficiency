using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Match the height of all selected shapes to the first selected shape.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnMatchHeight",
        Name = "Match Height",
        Tooltip = "Match height to the first selected shape",
        IconData = "M5,3V5H19V3H5M5,19V21H19V19H5M12,6V18M10,8L12,6L14,8M10,16L12,18L14,16",
        Color = "#10B981",
        Description = "Matches selected shapes to the height of the configured first or last selected anchor while preserving width and position.",
        DetailedHelpText = "### Match Height\n\nKeeps the first or last selected anchor configured in Settings fixed and resizes the remaining shapes to its exact height. Right-click any Match button to choose width, height, or both dimensions directly. With exactly two selected shapes, **Match Height and Swap Positions** first matches the non-anchor height and then exchanges both center positions in one Undo step. Width and existing aspect-ratio lock state are preserved.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class MatchHeightToFirstFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return manager != null && AlignmentHelpers.MatchSize(manager, matchWidth: false, matchHeight: true);
        }

        public static bool ExecuteAndSwapPositions(PowerPointManager manager)
        {
            if (manager == null) return false;
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count != 2) return false;

            var first = shapeRange[1];
            var second = shapeRange[2];
            var anchor = manager.GetReferenceShape();
            if (anchor == null) return false;
            var resized = first.Id == anchor.Id ? second : first;

            float firstLeft = first.Left;
            float firstTop = first.Top;
            float firstHeight = first.Height;
            float secondLeft = second.Left;
            float secondTop = second.Top;
            float secondHeight = second.Height;
            var firstLock = first.LockAspectRatio;
            var secondLock = second.LockAspectRatio;
            var resizedLock = resized.LockAspectRatio;

            try
            {
                try { manager.GetApplication()?.StartNewUndoEntry(); } catch { /* Resize remains valid when undo grouping is unavailable. */ }

                resized.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                resized.Height = anchor.Height;
                resized.LockAspectRatio = resizedLock;

                float firstCenterX = first.Left + first.Width / 2f;
                float firstCenterY = first.Top + first.Height / 2f;
                float secondCenterX = second.Left + second.Width / 2f;
                float secondCenterY = second.Top + second.Height / 2f;

                first.Left = secondCenterX - first.Width / 2f;
                first.Top = secondCenterY - first.Height / 2f;
                second.Left = firstCenterX - second.Width / 2f;
                second.Top = firstCenterY - second.Height / 2f;
                return true;
            }
            catch (System.Exception ex)
            {
                try
                {
                    first.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                    first.Height = firstHeight;
                    first.Left = firstLeft;
                    first.Top = firstTop;
                    first.LockAspectRatio = firstLock;
                }
                catch (System.Exception rollbackEx) { oPenEfficiency.Utils.ExceptionLogger.Log(rollbackEx, "MatchHeightToFirstFeature.RollbackFirst"); }
                try
                {
                    second.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                    second.Height = secondHeight;
                    second.Left = secondLeft;
                    second.Top = secondTop;
                    second.LockAspectRatio = secondLock;
                }
                catch (System.Exception rollbackEx) { oPenEfficiency.Utils.ExceptionLogger.Log(rollbackEx, "MatchHeightToFirstFeature.RollbackSecond"); }
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "MatchHeightToFirstFeature.ExecuteAndSwapPositions");
                return false;
            }
        }
    }
}
