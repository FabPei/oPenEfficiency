using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Models;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert a grouped illustrative sticker (text bracketed by two solid lines).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnIllustrativeSticker",
        Name = "Illustrative Sticker",
        Tooltip = "Illustrative sticker",
        IconData = "M17,17H7V14L3,18L7,22V19H19V13H17M7,7H17V10L21,6L17,2V5H5V11H7V7Z",
        Color = "#F43F5E",
        Description = "Inserts a text bracketed by two solid lines (configurable color, font, size).",
        DetailedHelpText = "### Illustrative Sticker\nInserts a pre-styled consulting sticker (ribbon) onto the slide for status indication.\n**Right-Click Options:**\n* Choose from **Preset Labels**: 'Work in progress', 'Updated', 'For discussion', 'Backup', 'Confidential', or 'Strictly confidential'.",
        Keywords = "sticker, label, illustrative, status, ribbon",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class IllustrativeStickerFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - inserts sticker with default text.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, "Illustrative");
        }

        public static bool Execute(PowerPointManager manager, string text = "Illustrative")
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                var config = oPenEfficiency.JsonService.Load<SidebarConfig>("sidebar_config.json");
                
                int accentColor = 0x0000FF; // Default Blue BGR
                float fontSize = 14f;
                string fontName = "";
                bool fontBold = true;
                bool fontItalic = false;
                float lineSize = 3f;

                if (config != null)
                {
                    try
                    {
                        var hex = config.IllustrativeStickerColor.Replace("#", "");
                        if (hex.Length == 6)
                        {
                            int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                            int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                            int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                            accentColor = r | (g << 8) | (b << 16);
                        }
                    } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"IllustrativeStickerFeature color parse error: {ex.Message}"); }

                    fontSize = (float)config.IllustrativeStickerFontSize;
                    fontName = config.IllustrativeStickerFont;
                    fontBold = config.IllustrativeStickerFontBold;
                    fontItalic = config.IllustrativeStickerFontItalic;
                    lineSize = (float)config.IllustrativeStickerLineSize;
                }

                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;

                // 1. Create Text Box
                var textShape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 100, 30);
                var tr = textShape.TextFrame.TextRange;
                tr.Text = text;
                tr.Font.Bold = fontBold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                tr.Font.Italic = fontItalic ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                tr.Font.Size = fontSize > 0 ? fontSize : 14f;
                tr.Font.Color.RGB = accentColor;
                if (!string.IsNullOrEmpty(fontName) && fontName != "Theme Font")
                {
                    tr.Font.Name = fontName;
                }
                tr.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                
                textShape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                textShape.TextFrame.MarginLeft = 0;
                textShape.TextFrame.MarginRight = 0;
                textShape.TextFrame.MarginTop = 0;
                textShape.TextFrame.MarginBottom = 0;
                
                textShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                float finalWidth = textShape.Width;
                float finalHeight = textShape.Height;
                textShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeNone;
                textShape.Width = finalWidth;
                textShape.Height = finalHeight;

                float margin = 10f;
                float lineOffset = 8f;
                float lineHeight = lineSize > 0 ? lineSize : 3f;
                float lineWidth = textShape.Width;

                textShape.Left = slideWidth - textShape.Width - margin;
                textShape.Top = margin + lineOffset + lineHeight;

                // 3. Create Top Line
                var topLine = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 
                    textShape.Left,
                    textShape.Top - lineOffset - lineHeight,
                    lineWidth, lineHeight);
                topLine.Fill.ForeColor.RGB = accentColor;
                topLine.Line.Visible = Office.MsoTriState.msoFalse;

                // 4. Create Bottom Line
                var bottomLine = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                    textShape.Left,
                    textShape.Top + textShape.Height + lineOffset,
                    lineWidth, lineHeight);
                bottomLine.Fill.ForeColor.RGB = accentColor;
                bottomLine.Line.Visible = Office.MsoTriState.msoFalse;

                string[] shapeNames = { textShape.Name, topLine.Name, bottomLine.Name };
                var group = slide.Shapes.Range(shapeNames).Group();
                group.Name = "StickyNote_Illustrative_" + Guid.NewGuid().ToString().Substring(0, 8);
                group.Select();

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IllustrativeStickerFeature.Execute");
                return false;
            }
        }
    }
}
