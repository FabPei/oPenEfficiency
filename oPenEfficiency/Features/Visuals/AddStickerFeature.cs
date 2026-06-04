using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Add a sticker-like object to the top-right corner of the active slide.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAddSticker",
        Name = "Add Sticker",
        Tooltip = "Add sticker",
        IconData = "M19,3H5C3.89,3 3,3.89 3,5V19C3,20.11 3.89,21 5,21H19C20.11,21 21,20.11 21,19V5C21,3.89 20.11,3 19,3M19,19H5V5H19V19M7,17H9V7H7V17M11,17H13V7H11V17M15,17H17V7H15V17Z",
        Color = "#F43F5E",
        Description = "Adds a small colored sticker (rounded rectangle) to the top-right corner of the slide.",
        DetailedHelpText = "### Add Sticker\nInserts one of the curated icon sticker assets from the library onto the current slide at a default size and position.",
        Keywords = "sticker, label, tag, note, post-it",
        MinSelection = 0)]
    public static class AddStickerFeature
    {
        public static bool Execute(PowerPointManager manager)
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

                var shape = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                    left, top, stickerWidth, stickerHeight);
                shape.Name = "StickyNote_Simple_" + Guid.NewGuid().ToString().Substring(0, 8);

                shape.Fill.ForeColor.RGB = 0x50AF4C; // Green (#4CAF50 in BGR)
                shape.Line.Visible = Office.MsoTriState.msoFalse;

                var config = oPenEfficiency.JsonService.Load<oPenEfficiency.Models.SidebarConfig>("sidebar_config.json");
                string defaultText = config?.StickyNoteDefaultText ?? "";
                string preferredFont = config?.StickyNoteFont ?? "";
                shape.TextFrame.TextRange.Text = defaultText;
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
                System.Diagnostics.Debug.WriteLine($"AddSticker Error: {ex}");
                return false;
            }
        }
    }
}
