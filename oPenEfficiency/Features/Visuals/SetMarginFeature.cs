using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Visuals
{
    [FeatureMetadata(
        Id = "BtnSetMargin",
        Name = "Paste Text Margins",
        Tooltip = "Paste copied internal text margins onto the selected shapes",
        SearchAliases = new[] { "Set Margin", "Set Text Margins", "Apply All Learned Margins", "Apply Horizontal Margins", "Apply Vertical Margins" },
        IconData = "M5,2H17A3,3 0 0 1 20,5V11H18V5A1,1 0 0 0 17,4H5A1,1 0 0 0 4,5V19A1,1 0 0 0 5,20H11V22H5A3,3 0 0 1 2,19V5A3,3 0 0 1 5,2Z M6,6H15V8H8V17H6Z M10,10H15V12H10Z M10,14H12V16H10Z M23,17H16.4L18.7,14.7L17.3,13.3L12.6,18L17.3,22.7L18.7,21.3L16.4,19H23Z",
        Color = "#3B82F6",
        Description = "Pastes the internal text margins saved by Copy Text Margins onto the selected shapes with text frames.",
        DetailedHelpText = "### Paste Text Margins\nUse **Copy Text Margins** on one source shape first, then select the target shapes. This command applies all four copied internal text margins and skips shapes without text frames. Values come from the current add-in session, not the Windows clipboard. Right-click to paste all margins, only left and right, only top and bottom, or reset every margin to zero. Reset does not require copied values. Each action creates one Undo entry.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class SetMarginFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, applyHorizontal: true, applyVertical: true, resetToZero: false);
        }

        public static bool Execute(PowerPointManager manager, bool applyHorizontal, bool applyVertical, bool resetToZero)
        {
            if (manager == null) return false;
            if (!oPenEfficiency.Core.Services.MarginPastePolicy.TryBuildPlan(
                    applyHorizontal, applyVertical, resetToZero,
                    MarginState.HasData, MarginState.Left, MarginState.Right, MarginState.Top, MarginState.Bottom,
                    out oPenEfficiency.Core.Services.MarginPastePolicy.MarginPlan plan))
            {
                System.Windows.MessageBox.Show("Use 'Copy Text Margins' on a source shape first.", "Paste Text Margins", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                var selection = app.ActiveWindow.Selection;
                if (selection == null || selection.Type != PpSelectionType.ppSelectionShapes) return false;
                try { app.StartNewUndoEntry(); } catch { /* Margin changes remain valid when undo grouping is unavailable. */ }
                bool changed = false;
                foreach (Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        if (plan.ApplyHorizontal)
                        {
                            shape.TextFrame.MarginLeft = plan.Left;
                            shape.TextFrame.MarginRight = plan.Right;
                        }
                        if (plan.ApplyVertical)
                        {
                            shape.TextFrame.MarginTop = plan.Top;
                            shape.TextFrame.MarginBottom = plan.Bottom;
                        }
                        changed = true;
                    }
                }
                return changed;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SetMarginFeature");
                return false;
            }
        }
    }
}
