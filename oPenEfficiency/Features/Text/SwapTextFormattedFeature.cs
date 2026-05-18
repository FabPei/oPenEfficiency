using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Swap text between two shapes while preserving their formatting.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSwapTextFormatted",
        Name = "Swap Text (Formatted)",
        Tooltip = "Swap text (formatted)",
        IconData = "M19,4H15.5L14.5,3H9.5L8.5,4H5V6H19M6,19A2,2 0 0,0 8,21H16A2,2 0 0,0 18,19V7H6V19Z",
        Color = "#6B7280",
        Description = "Swaps text between two selected shapes while preserving each shape's formatting.",
        DetailedHelpText = "### Swap Text (Formatted)\nExchanges the text content between two selected text boxes while preserving the original formatting of each destination box.",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class SwapTextFormattedFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count != 2) return false;

            try
            {
                var s1 = shapes[1];
                var s2 = shapes[2];
                if (s1.HasTextFrame != Office.MsoTriState.msoTrue || s2.HasTextFrame != Office.MsoTriState.msoTrue) return false;

                // Create a temporary shape to hold text content
                var slide = s1.Parent as Slide;
                var temp = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, -1000, -1000, 10, 10);
                
                s1.TextFrame.TextRange.Copy();
                temp.TextFrame.TextRange.Paste();

                s2.TextFrame.TextRange.Copy();
                s1.TextFrame.TextRange.Paste();

                temp.TextFrame.TextRange.Copy();
                s2.TextFrame.TextRange.Paste();

                temp.Delete();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SwapTextFormattedFeature.Execute");
                return false;
            }
        }
    }
}
