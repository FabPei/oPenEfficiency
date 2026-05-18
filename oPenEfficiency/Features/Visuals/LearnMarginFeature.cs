using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Visuals
{
    public static class MarginState
    {
        public static bool HasData = false;
        public static float Left = 0;
        public static float Right = 0;
        public static float Top = 0;
        public static float Bottom = 0;
    }

    [FeatureMetadata(
        Id = "BtnLearnMargin",
        Name = "Learn Margin",
        Tooltip = "Learn Text Margins",
        IconData = "M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10h-4v4h-2v-4H7v-2h4V7h2v4h4v2z",
        Color = "#3B82F6",
        Description = "Learns the internal text margins from the selected shape.",
        DetailedHelpText = "Learns the internal text margins from the selected shape so they can be applied to others.",
        MinSelection = 1,
        MaxSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class LearnMarginFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                var selection = app.ActiveWindow.Selection;
                if (selection == null || selection.Type != PpSelectionType.ppSelectionShapes) return false;
                var shapes = selection.ShapeRange;
                if (shapes.Count != 1) return false;

                var shape = shapes[1];
                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    MarginState.Left = shape.TextFrame.MarginLeft;
                    MarginState.Right = shape.TextFrame.MarginRight;
                    MarginState.Top = shape.TextFrame.MarginTop;
                    MarginState.Bottom = shape.TextFrame.MarginBottom;
                    MarginState.HasData = true;
                    
                    System.Windows.MessageBox.Show("Margins learned successfully. Select another shape and use 'Set Margin'.", "Learn Margin", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                }
                else
                {
                    System.Windows.MessageBox.Show("Selected shape does not have a text frame.", "Learn Margin", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LearnMarginFeature");
                return false;
            }
        }
    }
}