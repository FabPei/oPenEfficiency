using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Swap plain text between two shapes, ignoring formatting.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSwapTextPlain",
        Name = "Swap Text (Plain)",
        Tooltip = "Swap text (plain)",
        IconData = "M19,4H15.5L14.5,3H9.5L8.5,4H5V6H19M6,19A2,2 0 0,0 8,21H16A2,2 0 0,0 18,19V7H6V19Z",
        Color = "#6B7280",
        Description = "Swaps the plain text content between two selected shapes, ignoring formatting.",
        DetailedHelpText = "### Swap Text (Plain)\nExchanges the raw text string between two selected text boxes without carrying over any formatting.",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class SwapTextPlainFeature
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

                string t1 = s1.TextFrame.TextRange.Text;
                string t2 = s2.TextFrame.TextRange.Text;

                s1.TextFrame.TextRange.Text = t2;
                s2.TextFrame.TextRange.Text = t1;

                return true;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SwapTextPlain Error: {ex}");
                return false;
            }
        }
    }
}
