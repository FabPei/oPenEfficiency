using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Add a sticker-like object to the top-right corner of the active slide.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAddSticker",
        Name = "Add Corner Sticker",
        SearchAliases = new[] { "Add Sticker" },
        Tooltip = "Add a compact status sticker to the slide corner",
        IconData = "M19,3H5C3.89,3 3,3.89 3,5V19C3,20.11 3.89,21 5,21H19C20.11,21 21,20.11 21,19V5C21,3.89 20.11,3 19,3M19,19H5V5H19V19M7,17H9V7H7V17M11,17H13V7H11V17M15,17H17V7H15V17Z",
        Color = "#F43F5E",
        Description = "Adds a small colored sticker (rounded rectangle) to the top-right corner of the slide.",
        DetailedHelpText = "### Add Corner Sticker\nInserts a compact rounded status sticker in the top-right corner of the current slide. The primary action uses the configured default text and font. Right-click inserts ready-made Draft, Review, Approved, or Confidential stickers.",
        MinSelection = 0)]
    public static class AddStickerFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var config = oPenEfficiency.JsonService.Load<oPenEfficiency.Models.SidebarConfig>("sidebar_config.json");
                return CreateSticker(manager, config?.StickyNoteDefaultText ?? "", config?.StickyNoteFont ?? "", 0x50AF4C);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AddStickerFeature.Execute");
                return false;
            }
        }

        public static bool ExecutePreset(PowerPointManager manager, string text, int fillColorBgr)
        {
            try
            {
                var config = oPenEfficiency.JsonService.Load<oPenEfficiency.Models.SidebarConfig>("sidebar_config.json");
                return CreateSticker(manager, text ?? "", config?.StickyNoteFont ?? "", fillColorBgr);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AddStickerFeature.ExecutePreset");
                return false;
            }
        }

        private static bool CreateSticker(PowerPointManager manager, string text, string preferredFont, int fillColorBgr)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;

                float stickerWidth = 100f;
                float stickerHeight = 30f;
                float margin = 10f;
                float left = slideWidth - stickerWidth - margin;
                float top = margin;

                try { app.StartNewUndoEntry(); } catch { /* Sticker insertion remains valid when undo grouping is unavailable. */ }

                var shape = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                    left, top, stickerWidth, stickerHeight);
                shape.Name = "StickyNote_Simple_" + Guid.NewGuid().ToString().Substring(0, 8);

                shape.Fill.ForeColor.RGB = fillColorBgr;
                shape.Line.Visible = Office.MsoTriState.msoFalse;

                shape.TextFrame.TextRange.Text = text;
                shape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                if (!string.IsNullOrEmpty(preferredFont)) {
                    shape.TextFrame.TextRange.Font.Name = preferredFont;
                }
                shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF; // White
                shape.TextFrame.TextRange.Font.Size = 10;
                shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                shape.TextFrame.MarginLeft = 2;
                shape.TextFrame.MarginRight = 2;
                shape.TextFrame.MarginTop = 2;
                shape.TextFrame.MarginBottom = 2;

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AddStickerFeature.CreateSticker");
                return false;
            }
        }
    }
}
