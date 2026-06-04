using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert text at the current cursor position or at the end of the selected shape's text.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnInsertText",
        Name = "Insert Text",
        Tooltip = "Insert text at cursor",
        IconData = "M19,13H13V19H11V13H5V11H11V5H13V11H19V13Z",
        Color = "#FBBF24",
        Description = "Inserts text at the current cursor position in the active text box.",
        DetailedHelpText = "### Insert Text\nInserts a predefined text snippet at the current cursor position inside a text box.",
        Keywords = "type text, paste text, insert snippet, cursor position, add text",
        MinSelection = 0)]
    public static class InsertTextAtCursorFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - inserts default text (requires manual switch for custom text).
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            // Default execution - actual text is set via manual switch with dialog
            return false;
        }

        public static bool Execute(PowerPointManager manager, string text)
        {
            if (manager == null || string.IsNullOrEmpty(text)) return false;

            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;

                var selection = app.ActiveWindow.Selection;

                if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    selection.TextRange.InsertAfter(text);
                    return true;
                }
                else if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
                {
                    var shape = selection.ShapeRange[1];
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        shape.TextFrame.TextRange.InsertAfter(text);
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertTextAtCursorFeature.Execute");
                return false;
            }
        }
    }
}
