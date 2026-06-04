using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Clear text content from all selected shapes while preserving formatting.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnDeleteText",
        Name = "Delete Text",
        Tooltip = "Delete text",
        IconData = "M19,4H15.5L14.5,3H9.5L8.5,4H5V6H19M6,19A2,2 0 0,0 8,21H16A2,2 0 0,0 18,19V7H6V19Z",
        Color = "#EF4444",
        Description = "Clears all text from selected shapes while preserving their formatting and structure.",
        DetailedHelpText = "### Delete Text\nDeletes all text from the selected shapes, or when used with a slide selection, clears all text from all shapes on the selected slides simultaneously.",
        Keywords = "clear text, empty shapes, remove text, clear shapes, erase content",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class DeleteTextFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;

            try
            {
                for (int i = 1; i <= shapes.Count; i++)
                {
                    if (shapes[i].HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        shapes[i].TextFrame.TextRange.Text = "";
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DeleteTextFeature.Execute");
                return false;
            }
        }
    }
}
