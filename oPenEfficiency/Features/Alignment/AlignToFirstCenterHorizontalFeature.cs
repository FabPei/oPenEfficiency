using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

using System;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Align selected shapes horizontally centered relative to the first/master object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAlignCenter",
        Name = "Align Center",
        IconData = "M11,2H13V6H21V8H13V16H18V18H13V22H11V18H6V16H11V8H3V6H11V2Z",
        Color = "#10B981",
        Description = "Multi-select: centers shapes horizontally on the configured first/last anchor. Single select: snaps to a vertical guide or slide center.",
        DetailedHelpText = "### Align Center\n\nWith multiple selected shapes, keeps the anchor object configured in Settings (**First selected** or **Last selected**) fixed and moves every other shape so their horizontal centers match. With one shape, snaps its horizontal center to the nearest vertical presentation, master, or layout guide; the slide center is the fallback. Right-click to explicitly align within the selection, center every shape on the slide, snap each shape to its nearest guide, or use the configured anchor. Each action creates a PowerPoint Undo entry.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class AlignToFirstCenterHorizontalFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return AlignmentHelpers.ExecuteAlignment(manager, "centerv");
        }

        public static bool ExecuteWithinSelection(PowerPointManager manager)
        {
            try { manager?.GetApplication()?.StartNewUndoEntry(); }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "AlignToFirstCenterHorizontalFeature.AlignSelection.StartUndoEntry"); }
            return AlignmentHelpers.AlignShapes(manager, "center");
        }

        public static bool ExecuteToSlide(PowerPointManager manager)
        {
            try { manager?.GetApplication()?.StartNewUndoEntry(); }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "AlignToFirstCenterHorizontalFeature.AlignToSlide.StartUndoEntry"); }
            return AlignmentHelpers.AlignShapesToSlide(manager, "center");
        }

        public static bool ExecuteToNearestGuide(PowerPointManager manager)
        {
            try { manager?.GetApplication()?.StartNewUndoEntry(); }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "AlignToFirstCenterHorizontalFeature.SnapToGuide.StartUndoEntry"); }
            return AlignmentHelpers.SnapSelectionToNearestGuide(manager, "centerv");
        }

        public static bool ExecuteToAnchor(PowerPointManager manager)
        {
            try { manager?.GetApplication()?.StartNewUndoEntry(); }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "AlignToFirstCenterHorizontalFeature.AlignToFirst.StartUndoEntry"); }
            return AlignmentHelpers.AlignToFirst(manager, "centerv");
        }
    }
}
