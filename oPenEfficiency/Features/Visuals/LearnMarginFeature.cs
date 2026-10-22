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
        Name = "Copy Text Margins",
        Tooltip = "Copy all four internal text margins from the selected shape",
        SearchAliases = new[] { "Learn Margin", "Learn Text Margins", "Learn Margins from Selection", "Clear Learned Margins" },
        IconData = "M5,2H17A3,3 0 0 1 20,5V11H18V5A1,1 0 0 0 17,4H5A1,1 0 0 0 4,5V19A1,1 0 0 0 5,20H11V22H5A3,3 0 0 1 2,19V5A3,3 0 0 1 5,2Z M6,6H15V8H8V17H6Z M10,10H15V12H10Z M10,14H12V16H10Z M13,17H19.6L17.3,14.7L18.7,13.3L23.4,18L18.7,22.7L17.3,21.3L19.6,19H13Z",
        Color = "#3B82F6",
        Description = "Copies the left, right, top, and bottom internal text margins from one shape for use with Paste Text Margins.",
        DetailedHelpText = "### Copy Text Margins\nSelect exactly one shape with a text frame to copy its left, right, top, and bottom internal text margins. These are the distances between the text and the shape edges, not gaps between shapes. The values are stored for the current add-in session without changing the Windows clipboard. Select target shapes and use **Paste Text Margins** to apply them. Right-click to copy again or clear the copied values.",
        MinSelection = 1,
        MaxSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
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
                    
                    System.Windows.MessageBox.Show("Text margins copied. Select the target shapes and use 'Paste Text Margins'.", "Copy Text Margins", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                }
                else
                {
                    System.Windows.MessageBox.Show("Selected shape does not have a text frame.", "Copy Text Margins", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LearnMarginFeature");
                return false;
            }
        }

        public static bool ClearLearnedMargins()
        {
            bool hadData = MarginState.HasData;
            MarginState.HasData = false;
            MarginState.Left = MarginState.Right = MarginState.Top = MarginState.Bottom = 0;
            return hadData;
        }
    }
}
