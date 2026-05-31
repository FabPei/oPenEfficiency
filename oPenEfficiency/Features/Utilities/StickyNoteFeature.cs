using System;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert a sticky note on the slide.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStickyNote",
        Name = "Sticky Note",
        Tooltip = "Sticky note",
        IconData = "M19,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M19,19H5V5H19V19M14,17H7V15H14V17M17,13H7V11H17V13M17,9H7V7H17V9Z",
        Color = "#FBBF24",
        Description = "Inserts a sticky note at the top-right of the slide. Style is configurable in Settings.",
        DetailedHelpText = "### Sticky Note\nInserts a styled sticky-note annotation shape onto the slide for review comments or reminders directly on the canvas.",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class StickyNoteFeature
    {
        public static bool Execute(PowerPointManager manager, string text = null)
        {
            if (manager == null) return false;

            string style = "Theme";
            try
            {
                var config = oPenEfficiency.JsonService.Load<SidebarConfig>("sidebar_config.json");
                if (config != null && !string.IsNullOrEmpty(config.StickyNoteStyle))
                {
                    style = config.StickyNoteStyle;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StickyNoteFeature.Execute.ConfigLoad");
                style = "Theme";
            }

            switch (style)
            {
                case "Consulting":
                    return AddConsultingStickyNote(manager, text);
                case "GreenB":
                    return AddGreenBSticker(manager, text);
                case "Theme":
                default:
                    return InsertRealisticStickyNote(manager, text);
            }
        }

        public static bool AddGreenBSticker(PowerPointManager manager, string text = null, Slide targetSlide = null)
        {
            try
            {
                var app = manager.GetApplication();
                var slide = targetSlide ?? app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                float cmToPt = 28.346f;
                float width = 4.0f * cmToPt;
                float height = 1.0f * cmToPt;
                
                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float left = slideWidth - width;
                float top = 0f;

                var config = oPenEfficiency.JsonService.Load<SidebarConfig>("sidebar_config.json");
                string bgColorHex = "#FFFF00"; // GreenB sticky note always uses yellow
                string fontColorHex = config?.StickyNoteFontColor ?? "#575757";

                var shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
                shape.Name = "StickyNote_GreenB_" + Guid.NewGuid().ToString().Substring(0, 8);
                shape.Fill.ForeColor.RGB = HexToRgb(bgColorHex);
                shape.Fill.Solid();

                shape.Line.Visible = Office.MsoTriState.msoTrue;
                shape.Line.ForeColor.RGB = HexToRgb(fontColorHex); // Match font color for border
                shape.Line.Weight = 1.0f;

                shape.Shadow.Type = Office.MsoShadowType.msoShadow25; 
                shape.Shadow.Visible = Office.MsoTriState.msoTrue;
                shape.Shadow.ForeColor.RGB = 0x000000;
                shape.Shadow.Transparency = 0.0f;
                shape.Shadow.Blur = 0.0f;
                shape.Shadow.OffsetX = 4.0f;
                shape.Shadow.OffsetY = 4.0f;

                string defaultText = text ?? config?.StickyNoteDefaultText ?? "";
                string preferredFont = config?.StickyNoteFont ?? "";
                shape.TextFrame.TextRange.Text = defaultText;
                shape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                if (!string.IsNullOrEmpty(preferredFont)) {
                    shape.TextFrame.TextRange.Font.Name = preferredFont;
                } else {
                    shape.TextFrame.TextRange.Font.Name = "Segoe UI Variable Display";
                }
                shape.TextFrame.TextRange.Font.Size = 12;
                shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoFalse;
                shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                shape.TextFrame.TextRange.Font.Color.RGB = HexToRgb(fontColorHex);
                shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;

                if (string.IsNullOrEmpty(text)) shape.TextFrame.TextRange.Select();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StickyNoteFeature.AddGreenBSticker");
                return false;
            }
        }

        private static string GetThemeBodyFont(PowerPointManager manager)
        {
            try
            {
                var app = manager.GetApplication();
                if (app != null && app.ActivePresentation != null)
                {
                    return app.ActivePresentation.SlideMaster.Theme.ThemeFontScheme.MinorFont.Item(Office.MsoFontLanguageIndex.msoThemeLatin).Name;
                }
            }
            catch { }
            return "Segoe UI";
        }

        public static bool InsertRealisticStickyNote(PowerPointManager manager, string text = null)
        {
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                float cmToPt = 28.346f;
                float width = 6.5f * cmToPt;
                float height = 2.2f * cmToPt;
                float marginCm = 0.2f * cmToPt;

                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float left = slideWidth - width;
                float top = 0f;

                var shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeFoldedCorner, left, top, width, height);
                shape.Name = "StickyNote_Theme_" + Guid.NewGuid().ToString().Substring(0, 8);

                if (shape.Adjustments.Count > 0) shape.Adjustments[1] = 0.12f;

                shape.TextFrame.MarginLeft = marginCm;
                shape.TextFrame.MarginRight = marginCm;
                shape.TextFrame.MarginTop = marginCm;
                shape.TextFrame.MarginBottom = marginCm;

                var config = oPenEfficiency.JsonService.Load<SidebarConfig>("sidebar_config.json");
                string bgColorHex = "#FFF494"; // Standard sticky note uses light yellow
                string fontColorHex = config?.StickyNoteFontColor ?? "#323232";

                shape.Fill.ForeColor.RGB = HexToRgb(bgColorHex);
                shape.Fill.Solid();
                shape.Line.Visible = Office.MsoTriState.msoFalse;

                shape.Shadow.Type = Office.MsoShadowType.msoShadow21;
                shape.Shadow.Visible = Office.MsoTriState.msoTrue;
                shape.Shadow.ForeColor.RGB = 0x000000;
                shape.Shadow.Transparency = 0.6f;
                shape.Shadow.Blur = 12f;
                shape.Shadow.OffsetX = 0f;
                shape.Shadow.OffsetY = 5f;
                
                string defaultText = text ?? config?.StickyNoteDefaultText ?? "";
                string preferredFont = config?.StickyNoteFont ?? "";
                var textRange = shape.TextFrame.TextRange;
                textRange.Text = defaultText;
                shape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                if (!string.IsNullOrEmpty(preferredFont)) {
                    textRange.Font.Name = preferredFont;
                } else {
                    textRange.Font.Name = GetThemeBodyFont(manager);
                }
                textRange.Font.Size = 10.5f;
                textRange.Font.Bold = Office.MsoTriState.msoFalse;
                textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                textRange.Font.Color.RGB = HexToRgb(fontColorHex);
                
                shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                if (string.IsNullOrEmpty(text)) textRange.Select();

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StickyNoteFeature.InsertRealisticStickyNote");
                return false;
            }
        }

        private static int HexToRgb(string hex)
        {
            try
            {
                if (string.IsNullOrEmpty(hex)) return 0;
                hex = hex.Replace("#", "");
                if (hex.Length == 3)
                    hex = $"{hex[0]}{hex[0]}{hex[1]}{hex[1]}{hex[2]}{hex[2]}";

                int r = int.Parse(hex.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
                int g = int.Parse(hex.Substring(2, 2), System.Globalization.NumberStyles.HexNumber);
                int b = int.Parse(hex.Substring(4, 2), System.Globalization.NumberStyles.HexNumber);

                // PowerPoint RGB is 0xBBGGRR (actually it's a COLORREF which is 0x00BBGGRR)
                return r + (g << 8) + (b << 16);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StickyNoteFeature.HexToRgb");
                return 0;
            }
        }

        public static bool AddConsultingStickyNote(PowerPointManager manager, string text = null, Slide targetSlide = null)
        {
            try
            {
                var app = manager.GetApplication();
                var slide = targetSlide ?? app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                float cmToPt = 28.346f;
                float width = 6.0f * cmToPt;
                float height = 2.5f * cmToPt;
                float marginCm = 0.2f * cmToPt;

                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float left = slideWidth - width;
                float top = 0f;

                var shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeFoldedCorner, left, top, width, height);
                shape.Name = "StickyNote_Consulting_" + Guid.NewGuid().ToString().Substring(0, 8);

                var config = oPenEfficiency.JsonService.Load<SidebarConfig>("sidebar_config.json");
                string bgColorHex = "#FFFF99"; // Consulting sticky note uses light yellow
                string fontColorHex = config?.StickyNoteFontColor ?? "#333333";

                shape.Fill.ForeColor.RGB = HexToRgb(bgColorHex);
                shape.Fill.Solid();

                shape.Line.Visible = Office.MsoTriState.msoTrue;
                shape.Line.ForeColor.RGB = 0x808080;
                shape.Line.Weight = 1.25f;

                if (shape.Adjustments.Count > 0) shape.Adjustments[1] = 0.18f; 

                shape.TextFrame.MarginLeft = marginCm;
                shape.TextFrame.MarginRight = marginCm;
                shape.TextFrame.MarginTop = marginCm;
                shape.TextFrame.MarginBottom = marginCm;

                string defaultText = text ?? config?.StickyNoteDefaultText ?? "";
                string preferredFont = config?.StickyNoteFont ?? "";
                var textRange = shape.TextFrame.TextRange;
                textRange.Text = defaultText;
                shape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                if (!string.IsNullOrEmpty(preferredFont)) {
                    textRange.Font.Name = preferredFont;
                } else {
                    textRange.Font.Name = GetThemeBodyFont(manager);
                }
                textRange.Font.Size = 10;
                textRange.Font.Bold = Office.MsoTriState.msoFalse;
                textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                textRange.Font.Color.RGB = HexToRgb(fontColorHex);

                shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;
                if (string.IsNullOrEmpty(text)) textRange.Select();

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StickyNoteFeature.AddConsultingStickyNote");
                return false;
            }
        }
    }
}
