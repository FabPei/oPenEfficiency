using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Visuals
{
    [FeatureMetadata(
        Id = "BtnSetMargin",
        Name = "Set Margin",
        Tooltip = "Set Text Margins",
        IconData = "M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10h-4v4h-2v-4H7v-2h4V7h2v4h4v2z",
        Color = "#3B82F6",
        Description = "Applies the learned text margins to the selected shapes.",
        DetailedHelpText = "Applies the learned text margins to the selected shapes.",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class SetMarginFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            if (!MarginState.HasData)
            {
                System.Windows.MessageBox.Show("Please use 'Learn Margin' first.", "Set Margin", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return false;
            }
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                var selection = app.ActiveWindow.Selection;
                if (selection == null || selection.Type != PpSelectionType.ppSelectionShapes) return false;
                
                foreach (Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        shape.TextFrame.MarginLeft = MarginState.Left;
                        shape.TextFrame.MarginRight = MarginState.Right;
                        shape.TextFrame.MarginTop = MarginState.Top;
                        shape.TextFrame.MarginBottom = MarginState.Bottom;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SetMarginFeature");
                return false;
            }
        }
    }
}