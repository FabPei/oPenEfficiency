using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Services;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert or update a Checkbox (tick/cross/empty).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnCheckbox",
        Name = "Checkbox",
        Tooltip = "Checkbox",
        IconData = "M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M19,19H5V5H19V19M10,17L5,12L6.41,10.59L10,14.17L17.59,6.58L19,8L10,17Z",
        Color = "#F43F5E",
        Description = "Inserts a vector-based checkmark or cross inside a box. Default is Tick state.",
        DetailedHelpText = "### Checkbox\nInserts a styled checkbox shape (checked or unchecked) onto the slide. Clicking the button toggles the visual state between checked and unchecked.",
        Keywords = "checkbox, tick, cross, toggle, checkmark",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class CheckboxFeature
    {
        private const string CheckboxTag = "oPE_Checkbox";

        public enum CheckboxState
        {
            Tick,
            Cross,
            Empty
        }

        /// <summary>
        /// Wrapper for auto-discovery - inserts checkbox with default Tick state.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, CheckboxState.Tick, hasFrame: true, fontSize: 24, colorRgb: 0x000000);
        }

        public static bool Execute(PowerPointManager manager, CheckboxState state, bool hasFrame, float fontSize, int colorRgb)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;

                // Update existing
                var selection = manager.GetSelectedShapes();
                if (selection != null && selection.Count == 1)
                {
                    var selected = selection[1];
                    bool isEE4P = ShapeMetadataService.GetTag(selected, "EE4P_SMART_ELEMENT") == "Check";
                    if ((selected.AlternativeText != null && selected.AlternativeText.StartsWith(CheckboxTag)) || isEE4P)
                    {
                        UpdateCheckbox(selected, state, hasFrame, fontSize, colorRgb);
                        return true;
                    }
                }

                // Insert new
                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                float size = 50f; 
                float left = (app.ActivePresentation.PageSetup.SlideWidth / 2) - (size / 2);
                float top = (app.ActivePresentation.PageSetup.SlideHeight / 2) - (size / 2);

                // Use a Textbox for symbol
                var shape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, size, size);
                
                // Configure essentials
                shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                shape.TextFrame.MarginLeft = 0;
                shape.TextFrame.MarginRight = 0;
                shape.TextFrame.MarginTop = 0;
                shape.TextFrame.MarginBottom = 0;
                
                UpdateCheckbox(shape, state, hasFrame, fontSize, colorRgb);
                
                // Tag it
                shape.AlternativeText = $"{CheckboxTag}|{(int)state}|{hasFrame}|{fontSize}|{colorRgb}";
                shape.Name = "Checkbox_" + Guid.NewGuid().ToString().Substring(0, 8);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CheckboxFeature.Execute");
                return false;
            }
        }

        private static void UpdateCheckbox(Shape shape, CheckboxState state, bool hasFrame, float fontSize, int colorRgb)
        {
            // Update Text
            string symbol = "";
            string fontName = "Wingdings"; // Common for symbols

            switch (state)
            {
                case CheckboxState.Tick:
                    symbol = "ü"; // Wingdings tick
                    break;
                case CheckboxState.Cross:
                    symbol = "û"; // Wingdings cross
                    break;
                case CheckboxState.Empty:
                    symbol = " ";
                    break;
            }

            shape.TextFrame.TextRange.Font.Name = fontName;
            shape.TextFrame.TextRange.Text = symbol;
            
            // Update Font Size
            shape.TextFrame.TextRange.Font.Size = fontSize;

            // Update Color
            if (colorRgb != -1)
            {
                shape.TextFrame.TextRange.Font.Color.RGB = colorRgb;
                if (hasFrame) shape.Line.ForeColor.RGB = colorRgb;
            }

            // Update Frame
            if (hasFrame)
            {
                shape.Line.Visible = Office.MsoTriState.msoTrue;
                shape.Line.Weight = 2f;
            }
            else
            {
                shape.Line.Visible = Office.MsoTriState.msoFalse;
            }

            // Update Tag
            shape.AlternativeText = $"{CheckboxTag}|{(int)state}|{hasFrame}|{fontSize}|{colorRgb}";
        }
    }
}
