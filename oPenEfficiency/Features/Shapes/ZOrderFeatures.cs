using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    internal static class ZOrderFeatureHelper
    {
        public static bool Execute(PowerPointManager manager, Office.MsoZOrderCmd command, string context)
        {
            var shapes = manager?.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;
            var selected = ToShapeList(shapes);
            var slide = selected[0].Parent as Slide;
            if (slide == null) return false;
            var originalOrder = CaptureOrder(slide);

            try
            {
                TryStartUndo(manager, context);
                shapes.ZOrder(command);
                return true;
            }
            catch (Exception ex)
            {
                TryRestore(originalOrder, context);
                Utils.ExceptionLogger.Log(ex, context);
                return false;
            }
        }

        public static bool ExecuteRanked(PowerPointManager manager, Office.MsoZOrderCmd command, bool firstSelectedOnTop, string context)
        {
            var shapes = manager?.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;
            if (shapes.Count == 1) return Execute(manager, command, context);

            var selected = ToShapeList(shapes);
            var slide = selected[0].Parent as Slide;
            if (slide == null) return false;
            var originalOrder = CaptureOrder(slide);

            try
            {
                TryStartUndo(manager, context);
                List<Shape> workingOrder;

                if (command == Office.MsoZOrderCmd.msoBringToFront ||
                    command == Office.MsoZOrderCmd.msoSendToBack)
                {
                    workingOrder = new List<Shape>(originalOrder);
                }
                else
                {
                    shapes.ZOrder(command);
                    workingOrder = CaptureOrder(slide);
                }

                var byId = workingOrder.ToDictionary(shape => shape.Id);
                var plan = ZOrderRankPlanner.PlanRanked(
                    workingOrder.Select(shape => shape.Id).ToList(),
                    selected.Select(shape => shape.Id).ToList(),
                    firstSelectedOnTop,
                    command == Office.MsoZOrderCmd.msoBringToFront,
                    command == Office.MsoZOrderCmd.msoSendToBack);
                List<Shape> desiredOrder = plan.Select(id => byId[id]).ToList();

                ApplyOrder(desiredOrder);
                return true;
            }
            catch (Exception ex)
            {
                TryRestore(originalOrder, context);
                Utils.ExceptionLogger.Log(ex, context);
                return false;
            }
        }

        public static bool ExecuteRelativeToLastSelected(PowerPointManager manager, bool placeAbove, string context)
        {
            var shapes = manager?.GetSelectedShapes();
            if (shapes == null || shapes.Count < 2) return false;

            var selected = ToShapeList(shapes);
            var reference = selected[selected.Count - 1];
            var slide = reference.Parent as Slide;
            if (slide == null) return false;

            var originalOrder = CaptureOrder(slide);
            var moverIds = new HashSet<int>(selected
                .Take(selected.Count - 1)
                .Select(shape => shape.Id));
            var byId = originalOrder.GroupBy(shape => shape.Id).ToDictionary(group => group.Key, group => group.First());
            var plan = ZOrderRankPlanner.PlanRelativeToReference(
                originalOrder.Select(shape => shape.Id).ToList(),
                moverIds,
                reference.Id,
                placeAbove);
            if (plan == null) return false;
            List<Shape> desiredOrder = plan.Select(id => byId[id]).ToList();

            try
            {
                TryStartUndo(manager, context);
                ApplyOrder(desiredOrder);
                return true;
            }
            catch (Exception ex)
            {
                TryRestore(originalOrder, context);
                Utils.ExceptionLogger.Log(ex, context);
                return false;
            }
        }

        private static List<Shape> ToShapeList(ShapeRange shapes)
        {
            var result = new List<Shape>();
            for (int i = 1; i <= shapes.Count; i++) result.Add(shapes[i]);
            return result;
        }

        private static List<Shape> CaptureOrder(Slide slide)
        {
            var result = new List<Shape>();
            for (int i = 1; i <= slide.Shapes.Count; i++) result.Add(slide.Shapes[i]);
            return result.OrderBy(shape => shape.ZOrderPosition).ToList();
        }

        private static void ApplyOrder(IList<Shape> bottomToTop)
        {
            for (int i = 0; i < bottomToTop.Count; i++)
                bottomToTop[i].ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }

        private static void TryStartUndo(PowerPointManager manager, string context)
        {
            try { manager.GetApplication()?.StartNewUndoEntry(); }
            catch (Exception ex) { Utils.ExceptionLogger.Log(ex, context + ".StartNewUndoEntry"); }
        }

        private static void TryRestore(IList<Shape> originalOrder, string context)
        {
            try { ApplyOrder(originalOrder); }
            catch (Exception ex) { Utils.ExceptionLogger.Log(ex, context + ".Rollback"); }
        }
    }

    [FeatureMetadata(
        Id = "BtnBringToFront",
        Name = "Bring to Front",
        Tooltip = "Move the selection in front of every other shape",
        IconData = "M14 2h1a2 2 0 0 1 2 2v1 M7 2H6a2 2 0 0 0-2 2v1 M2 7V6a2 2 0 0 1 2-2h1 M22 7V6a2 2 0 0 0-2-2h-1 M14 22h1a2 2 0 0 0 2-2v-1 M7 22H6a2 2 0 0 1-2-2v-1 M2 17v1a2 2 0 0 0 2 2h1 M22 17v1a2 2 0 0 1-2 2h-1 M8 8h8v8h-8z",
        Color = "#10B981",
        Description = "Brings the selected shapes to the absolute front layer.",
        DetailedHelpText = "### Bring to Front\n\nMoves the selected shapes above every other object on the slide. Right-click variants:\n\n- **Rank Selected Shapes (First Selected on Top)** makes the first selected shape topmost.\n- **Place Directly Above Last Selected** uses the last selected shape as the reference.\n- **Place Directly Below Last Selected** uses the last selected shape as the reference.\n\nRelative placement preserves the movers' current internal layer order. Unselected shapes keep their relative order.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class BringToFrontFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return ZOrderFeatureHelper.Execute(manager, Office.MsoZOrderCmd.msoBringToFront, "BringToFrontFeature.Execute");
        }

        public static bool ExecuteRanked(PowerPointManager manager)
        {
            return ZOrderFeatureHelper.ExecuteRanked(manager, Office.MsoZOrderCmd.msoBringToFront, true, "BringToFrontFeature.ExecuteRanked");
        }
    }

    [FeatureMetadata(
        Id = "BtnSendToBack",
        Name = "Send to Back",
        Tooltip = "Move the selection behind every other shape",
        IconData = "M7 2h1a2 2 0 0 1 2 2v1 M2 7V6a2 2 0 0 1 2-2h1 M22 14v1a2 2 0 0 1-2 2h-1 M14 22h1a2 2 0 0 0 2-2v-1 M14 14h8v8h-8z M2 2h8v8h-8z",
        Color = "#10B981",
        Description = "Sends the selected shapes to the absolute back layer.",
        DetailedHelpText = "### Send to Back\n\nMoves the selected shapes behind every other object on the slide. Right-click variants:\n\n- **Rank Selected Shapes (First Selected at Bottom)** makes the first selected shape bottommost.\n- **Place Directly Above Last Selected** uses the last selected shape as the reference.\n- **Place Directly Below Last Selected** uses the last selected shape as the reference.\n\nRelative placement preserves the movers' current internal layer order. Unselected shapes keep their relative order.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class SendToBackFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return ZOrderFeatureHelper.Execute(manager, Office.MsoZOrderCmd.msoSendToBack, "SendToBackFeature.Execute");
        }

        public static bool ExecuteRanked(PowerPointManager manager)
        {
            return ZOrderFeatureHelper.ExecuteRanked(manager, Office.MsoZOrderCmd.msoSendToBack, false, "SendToBackFeature.ExecuteRanked");
        }
    }

    [FeatureMetadata(
        Id = "BtnBringForward",
        Name = "Bring Forward",
        Tooltip = "Move the selection one layer forward",
        IconData = "M6 2h12a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z M12 14V6 M9 9l3-3 3 3",
        Color = "#10B981",
        Description = "Brings the selected shapes forward by one layer.",
        DetailedHelpText = "### Bring Forward\n\nMoves the selection forward by one position in the stacking order. Right-click variants:\n\n- **Rank Selected Shapes (First Selected on Top)** performs the same one-step move and puts the first selected shape on top.\n- **Place Directly Above Last Selected** uses the last selected shape as the reference.\n- **Place Directly Below Last Selected** uses the last selected shape as the reference.\n\nRelative placement preserves the movers' current internal layer order and the relative order of unselected shapes.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class BringForwardFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return ZOrderFeatureHelper.Execute(manager, Office.MsoZOrderCmd.msoBringForward, "BringForwardFeature.Execute");
        }

        public static bool ExecuteRanked(PowerPointManager manager)
        {
            return ZOrderFeatureHelper.ExecuteRanked(manager, Office.MsoZOrderCmd.msoBringForward, true, "BringForwardFeature.ExecuteRanked");
        }
    }

    [FeatureMetadata(
        Id = "BtnSendBackward",
        Name = "Send Backward",
        Tooltip = "Move the selection one layer backward",
        IconData = "M6 2h12a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z M12 6v8 M9 11l3 3 3-3",
        Color = "#10B981",
        Description = "Sends the selected shapes backward by one layer.",
        DetailedHelpText = "### Send Backward\n\nMoves the selection backward by one position in the stacking order. Right-click variants:\n\n- **Rank Selected Shapes (First Selected at Bottom)** performs the same one-step move and puts the first selected shape at the bottom.\n- **Place Directly Above Last Selected** uses the last selected shape as the reference.\n- **Place Directly Below Last Selected** uses the last selected shape as the reference.\n\nRelative placement preserves the movers' current internal layer order and the relative order of unselected shapes.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class SendBackwardFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return ZOrderFeatureHelper.Execute(manager, Office.MsoZOrderCmd.msoSendBackward, "SendBackwardFeature.Execute");
        }

        public static bool ExecuteRanked(PowerPointManager manager)
        {
            return ZOrderFeatureHelper.ExecuteRanked(manager, Office.MsoZOrderCmd.msoSendBackward, false, "SendBackwardFeature.ExecuteRanked");
        }
    }
}
