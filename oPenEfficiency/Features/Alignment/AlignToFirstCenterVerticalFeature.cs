using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

using System;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Align selected shapes vertically centered relative to the first/master object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAlignCenterVertical",
        Name = "Align Center Vertical",
        IconData = "M2,11V13H6V21H8V13H16V18H18V13H22V11H18V6H16V11H8V3H6V11H2Z",
        Color = "#10B981",
        Description = "Multi-select: centers shapes vertically on the configured first/last anchor. Single select: snaps to a horizontal guide or slide middle.",
        DetailedHelpText = "### Align Center Vertical\n\nWith multiple selected shapes, keeps the anchor object configured in Settings (**First selected** or **Last selected**) fixed and moves every other shape so their vertical centers match. With one shape, snaps its vertical center to the nearest horizontal presentation, master, or layout guide; the slide middle is the fallback. Right-click to explicitly align within the selection, center every shape on the slide, snap each shape to its nearest guide, or use the configured anchor. Each action creates a PowerPoint Undo entry.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class AlignToFirstCenterVerticalFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return AlignmentHelpers.ExecuteAlignment(manager, "centerh");
        }

        public static bool ExecuteWithinSelection(PowerPointManager manager)
        {
            try { manager?.GetApplication()?.StartNewUndoEntry(); }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "AlignToFirstCenterVerticalFeature.AlignSelection.StartUndoEntry"); }
            return AlignmentHelpers.AlignShapes(manager, "middle");
        }

        public static bool ExecuteToSlide(PowerPointManager manager)
        {
            try { manager?.GetApplication()?.StartNewUndoEntry(); }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "AlignToFirstCenterVerticalFeature.AlignToSlide.StartUndoEntry"); }
            return AlignmentHelpers.AlignShapesToSlide(manager, "middle");
        }

        public static bool ExecuteToNearestGuide(PowerPointManager manager)
        {
            try { manager?.GetApplication()?.StartNewUndoEntry(); }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "AlignToFirstCenterVerticalFeature.SnapToGuide.StartUndoEntry"); }
            return AlignmentHelpers.SnapSelectionToNearestGuide(manager, "centerh");
        }

        public static bool ExecuteToAnchor(PowerPointManager manager)
        {
            try { manager?.GetApplication()?.StartNewUndoEntry(); }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "AlignToFirstCenterVerticalFeature.AlignToFirst.StartUndoEntry"); }
            return AlignmentHelpers.AlignToFirst(manager, "centerh");
        }
    }
}
