using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Alignment
{
    [FeatureMetadata(Id = "BtnDistributeHorizontal", Name = "Distribute Horizontally", Tooltip = "Space selected shapes evenly from left to right", IconData = "M4 22H2V2h2v20zM22 2v20h-2V2h2zm-4 4H6v12h12V6z", Color = "#10B981", Description = "Distributes selected shapes horizontally within the selection or across the slide.", DetailedHelpText = "### Distribute Horizontally\n\nThe primary action spaces three or more shapes evenly between the outer shapes; with exactly two selected shapes it spreads them across the slide instead, because there is no interior gap to distribute. Right-click to choose the across-slide route explicitly or to align their vertical middles before distribution. Each action creates one Undo entry.", MinSelection = 2, RequiredType = OfficeSelectionType.Shapes)]
    public static class DistributeHorizontalFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, false, false);
        }

        public static bool Execute(PowerPointManager manager, bool relativeToSlide, bool alignMiddles = false)
        {
            var app = manager?.GetApplication();
            if (app?.ActiveWindow?.Selection?.Type != PpSelectionType.ppSelectionShapes) return false;
            try
            {
                var sr = app.ActiveWindow.Selection.ShapeRange;
                if (sr == null) return false;
                // Owner decision a1 for REG-DISTRIBUTE-SCOPE-001: the sidebar enables this button from
                // MinSelection = 2, so the primary click must act at two shapes as well - it spreads them across the
                // slide, because distributing an interior gap needs a shape between the outer two. From three
                // shapes the primary route distributes the gaps. The whole decision comes from one host-neutral
                // call, and the route it resolves is the value the distribution receives (review round 6g).
                var plan = DistributeSelectionRule.ResolvePlan(sr.Count, relativeToSlide);
                if (!plan.Distribute) return false;

                var originals = new Tuple<Shape, float, float>[sr.Count];
                for (int i = 1; i <= sr.Count; i++) originals[i - 1] = Tuple.Create(sr[i], sr[i].Left, sr[i].Top);

                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "DistributeFeatures.Horizontal.StartUndoEntry"); }
                try
                {
                    sr.Distribute(Office.MsoDistributeCmd.msoDistributeHorizontally, plan.RelativeToSlide ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse);
                    if (alignMiddles) sr.Align(Office.MsoAlignCmd.msoAlignMiddles, Office.MsoTriState.msoFalse);
                }
                catch
                {
                    foreach (var original in originals)
                    {
                        try { original.Item1.Left = original.Item2; original.Item1.Top = original.Item3; }
                        catch (Exception rollbackEx) { ExceptionPolicy.ReportOptionalComFailure(rollbackEx, "DistributeFeatures.Horizontal.Rollback"); }
                    }
                    throw;
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DistributeHorizontalFeature.Execute");
                return false;
            }
        }
    }

    [FeatureMetadata(Id = "BtnDistributeVertical", Name = "Distribute Vertically", Tooltip = "Space selected shapes evenly from top to bottom", IconData = "M22 2v2H2V2h20zm0 18v2H2v-2h20zm-4-4H6V8h12v8z", Color = "#10B981", Description = "Distributes selected shapes vertically within the selection or across the slide.", DetailedHelpText = "### Distribute Vertically\n\nThe primary action spaces three or more shapes evenly between the outer shapes; with exactly two selected shapes it spreads them across the slide instead, because there is no interior gap to distribute. Right-click to choose the across-slide route explicitly or to align their horizontal centers before distribution. Each action creates one Undo entry.", MinSelection = 2, RequiredType = OfficeSelectionType.Shapes)]
    public static class DistributeVerticalFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, false, false);
        }

        public static bool Execute(PowerPointManager manager, bool relativeToSlide, bool alignCenters = false)
        {
            var app = manager?.GetApplication();
            if (app?.ActiveWindow?.Selection?.Type != PpSelectionType.ppSelectionShapes) return false;
            try
            {
                var sr = app.ActiveWindow.Selection.ShapeRange;
                if (sr == null) return false;
                // Owner decision a1 for REG-DISTRIBUTE-SCOPE-001: the sidebar enables this button from
                // MinSelection = 2, so the primary click must act at two shapes as well - it spreads them across the
                // slide, because distributing an interior gap needs a shape between the outer two. From three
                // shapes the primary route distributes the gaps. The whole decision comes from one host-neutral
                // call, and the route it resolves is the value the distribution receives (review round 6g).
                var plan = DistributeSelectionRule.ResolvePlan(sr.Count, relativeToSlide);
                if (!plan.Distribute) return false;

                var originals = new Tuple<Shape, float, float>[sr.Count];
                for (int i = 1; i <= sr.Count; i++) originals[i - 1] = Tuple.Create(sr[i], sr[i].Left, sr[i].Top);

                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "DistributeFeatures.Vertical.StartUndoEntry"); }
                try
                {
                    sr.Distribute(Office.MsoDistributeCmd.msoDistributeVertically, plan.RelativeToSlide ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse);
                    if (alignCenters) sr.Align(Office.MsoAlignCmd.msoAlignCenters, Office.MsoTriState.msoFalse);
                }
                catch
                {
                    foreach (var original in originals)
                    {
                        try { original.Item1.Left = original.Item2; original.Item1.Top = original.Item3; }
                        catch (Exception rollbackEx) { ExceptionPolicy.ReportOptionalComFailure(rollbackEx, "DistributeFeatures.Vertical.Rollback"); }
                    }
                    throw;
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DistributeVerticalFeature.Execute");
                return false;
            }
        }
    }
}
