using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Models;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    public enum StatusStickerScope
    {
        CurrentSlide,
        SelectedSlides,
        EntirePresentation
    }

    public enum StatusStickerPosition
    {
        TopLeft,
        TopRight,
        BottomLeft,
        BottomRight
    }

    /// <summary>
    /// Inserts and manages consulting status stickers. New stickers carry oPE_ tags so they
    /// can be updated consistently across slides without relying only on a shape name.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnIllustrativeSticker",
        Name = "Status Sticker",
        Tooltip = "Insert or manage slide status stickers",
        SearchAliases = new[] { "Status Stickers", "Illustrative Sticker", "Illustrative Stickers", "Draft", "Confidential", "Client-ready" },
        IconData = "M17,17H7V14L3,18L7,22V19H19V13H17M7,7H17V10L21,6L17,2V5H5V11H7V7Z",
        Color = "#F43F5E",
        Description = "Adds and manages consulting status stickers, including workflow, confidentiality and audience labels across one or many slides.",
        DetailedHelpText = "### Status Sticker\nClick to insert the default status label. Right-click to choose presets or a custom label, combine workflow, confidentiality and audience states, update an existing sticker, apply it to selected slides or the entire presentation, change its corner position, use the presentation theme color, add date or author metadata, remove stickers, or manage presets in Settings. Existing tagged stickers are updated instead of duplicated by bulk actions.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class IllustrativeStickerFeature
    {
        private const string StickerTag = "oPE_StatusSticker";
        private const string TextTag = "oPE_StatusText";
        private const string WorkTag = "oPE_StatusWork";
        private const string ConfidentialityTag = "oPE_StatusConfidentiality";
        private const string AudienceTag = "oPE_StatusAudience";
        private const string PositionTag = "oPE_StatusPosition";
        private const string ThemeColorTag = "oPE_StatusThemeColor";
        private const string DateTag = "oPE_StatusIncludeDate";
        private const string AuthorTag = "oPE_StatusIncludeAuthor";
        private const string ColorTag = "oPE_StatusColor";

        private sealed class StickerState
        {
            public string Text { get; set; }
            public string WorkStatus { get; set; }
            public string Confidentiality { get; set; }
            public string Audience { get; set; }
            public StatusStickerPosition Position { get; set; } = StatusStickerPosition.TopRight;
            public bool UseThemeColor { get; set; }
            public bool IncludeDate { get; set; }
            public bool IncludeAuthor { get; set; }
            public string ColorHex { get; set; }
        }

        public static bool Execute(PowerPointManager manager)
        {
            var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
            string label = StickerPreviewModel.GetFirstStatusLabel(config?.IllustrativeStickerTexts);
            return Execute(manager, label);
        }

        public static bool Execute(PowerPointManager manager, string text)
        {
            if (manager == null || string.IsNullOrWhiteSpace(text)) return false;
            var state = NewState(manager, text.Trim());
            return ApplyState(manager, state, StatusStickerScope.CurrentSlide, true);
        }

        public static bool ApplyPreset(PowerPointManager manager, string text, StatusStickerScope scope)
        {
            if (manager == null || string.IsNullOrWhiteSpace(text)) return false;
            return ApplyState(manager, NewState(manager, text.Trim()), scope, true);
        }

        public static bool ApplyDimension(
            PowerPointManager manager,
            string workStatus = null,
            string confidentiality = null,
            string audience = null)
        {
            if (manager == null) return false;
            var slide = manager.GetCurrentSlide();
            if (slide == null) return false;
            var state = ReadState(FindSticker(slide)) ?? NewState(manager, null);
            if (workStatus != null) state.WorkStatus = workStatus;
            if (confidentiality != null) state.Confidentiality = confidentiality;
            if (audience != null) state.Audience = audience;
            state.Text = null;
            state.UseThemeColor = false;
            state.ColorHex = GetSemanticColor(state);
            return ApplyState(manager, state, StatusStickerScope.CurrentSlide, true);
        }

        public static bool RefreshCurrent(PowerPointManager manager)
        {
            if (manager == null) return false;
            var state = ReadState(FindSticker(manager.GetCurrentSlide()));
            return state != null && ApplyState(manager, state, StatusStickerScope.CurrentSlide, true);
        }

        public static bool CopyCurrentToScope(PowerPointManager manager, StatusStickerScope scope)
        {
            if (manager == null) return false;
            var state = ReadState(FindSticker(manager.GetCurrentSlide()));
            return state != null && ApplyState(manager, state, scope, true);
        }

        public static bool SetPosition(PowerPointManager manager, StatusStickerPosition position)
        {
            if (manager == null) return false;
            var state = ReadState(FindSticker(manager.GetCurrentSlide()));
            if (state == null) return false;
            state.Position = position;
            return ApplyState(manager, state, StatusStickerScope.CurrentSlide, true);
        }

        public static bool UseThemeColor(PowerPointManager manager)
        {
            if (manager == null) return false;
            var state = ReadState(FindSticker(manager.GetCurrentSlide()));
            if (state == null) return false;
            state.UseThemeColor = true;
            return ApplyState(manager, state, StatusStickerScope.CurrentSlide, true);
        }

        public static bool ToggleDate(PowerPointManager manager)
        {
            if (manager == null) return false;
            var state = ReadState(FindSticker(manager.GetCurrentSlide()));
            if (state == null) return false;
            state.IncludeDate = !state.IncludeDate;
            return ApplyState(manager, state, StatusStickerScope.CurrentSlide, true);
        }

        public static bool ToggleAuthor(PowerPointManager manager)
        {
            if (manager == null) return false;
            var state = ReadState(FindSticker(manager.GetCurrentSlide()));
            if (state == null) return false;
            state.IncludeAuthor = !state.IncludeAuthor;
            return ApplyState(manager, state, StatusStickerScope.CurrentSlide, true);
        }

        public static bool Remove(PowerPointManager manager, StatusStickerScope scope)
        {
            if (manager == null) return false;
            try
            {
                int removed = 0;
                foreach (var slide in GetSlides(manager, scope))
                {
                    for (int i = slide.Shapes.Count; i >= 1; i--)
                    {
                        var shape = slide.Shapes[i];
                        if (IsSticker(shape))
                        {
                            shape.Delete();
                            removed++;
                        }
                    }
                }
                return removed > 0;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IllustrativeStickerFeature.Remove");
                return false;
            }
        }

        public static IReadOnlyList<string> GetPresetLabels()
        {
            var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
            var labels = SplitLabels(config?.IllustrativeStickerTexts).ToList();
            if (labels.Count == 0)
                labels.AddRange(SplitLabels("Work in progress, Updated, For discussion, Backup, Confidential, Strictly confidential"));
            return labels;
        }

        private static StickerState NewState(PowerPointManager manager, string text)
        {
            var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
            return new StickerState
            {
                Text = text,
                Position = ParsePosition(config?.IllustrativeStickerPosition),
                UseThemeColor = config?.IllustrativeStickerUseThemeColor ?? false,
                IncludeDate = config?.IllustrativeStickerIncludeDate ?? false,
                IncludeAuthor = config?.IllustrativeStickerIncludeAuthor ?? false,
                ColorHex = config?.IllustrativeStickerColor ?? "#FF0000"
            };
        }

        private static bool ApplyState(PowerPointManager manager, StickerState state, StatusStickerScope scope, bool updateExisting)
        {
            try
            {
                var slides = GetSlides(manager, scope);
                if (slides.Count == 0) return false;
                Shape last = null;
                foreach (var slide in slides)
                    last = CreateOrReplace(manager, slide, state, updateExisting);

                if (last != null && scope == StatusStickerScope.CurrentSlide)
                    last.Select();
                return last != null;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IllustrativeStickerFeature.ApplyState");
                return false;
            }
        }

        private static Shape CreateOrReplace(PowerPointManager manager, Slide slide, StickerState state, bool updateExisting)
        {
            Shape existing = updateExisting ? FindSticker(slide) : null;
            if (existing != null)
            {
                var old = ReadState(existing);
                if (old != null && string.IsNullOrEmpty(state.Text) && string.IsNullOrEmpty(state.WorkStatus)) state.Text = old.Text;
                existing.Delete();
            }

            var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
            int color = ResolveColor(manager, state, config);
            float fontSize = (float)(config?.IllustrativeStickerFontSize ?? 14d);
            float lineHeight = (float)(config?.IllustrativeStickerLineSize ?? 3d);
            string displayText = BuildDisplayText(state);
            if (string.IsNullOrWhiteSpace(displayText)) displayText = "Status";

            var textShape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 100, 30);
            var tr = textShape.TextFrame.TextRange;
            tr.Text = displayText;
            tr.Font.Bold = (config?.IllustrativeStickerFontBold ?? true) ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            tr.Font.Italic = (config?.IllustrativeStickerFontItalic ?? false) ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            tr.Font.Size = fontSize > 0 ? fontSize : 14f;
            tr.Font.Color.RGB = color;
            if (!string.IsNullOrWhiteSpace(config?.IllustrativeStickerFont) && config.IllustrativeStickerFont != "Theme Font")
                tr.Font.Name = config.IllustrativeStickerFont;
            tr.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;

            textShape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
            textShape.TextFrame.MarginLeft = 0;
            textShape.TextFrame.MarginRight = 0;
            textShape.TextFrame.MarginTop = 0;
            textShape.TextFrame.MarginBottom = 0;
            textShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            float width = textShape.Width;
            float height = textShape.Height;
            textShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeNone;
            textShape.Width = width;
            textShape.Height = height;

            float lineOffset = 8f;
            lineHeight = lineHeight > 0 ? lineHeight : 3f;
            var topLine = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, width, lineHeight);
            var bottomLine = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, height + (2 * lineOffset) + lineHeight, width, lineHeight);
            topLine.Fill.ForeColor.RGB = color;
            bottomLine.Fill.ForeColor.RGB = color;
            topLine.Line.Visible = Office.MsoTriState.msoFalse;
            bottomLine.Line.Visible = Office.MsoTriState.msoFalse;
            textShape.Left = 0;
            textShape.Top = lineOffset + lineHeight;

            string[] names = { textShape.Name, topLine.Name, bottomLine.Name };
            var group = slide.Shapes.Range(names).Group();
            group.Name = "oPE_StatusSticker_" + Guid.NewGuid().ToString("N").Substring(0, 8);
            WriteState(group, state);
            PositionWithCollisionAvoidance(manager, slide, group, state.Position);
            return group;
        }

        private static void PositionWithCollisionAvoidance(PowerPointManager manager, Slide slide, Shape group, StatusStickerPosition position)
        {
            var pres = manager.GetCurrentPresentation();
            float slideWidth = pres.PageSetup.SlideWidth;
            float slideHeight = pres.PageSetup.SlideHeight;
            const float margin = 10f;
            const float gap = 8f;
            bool right = position == StatusStickerPosition.TopRight || position == StatusStickerPosition.BottomRight;
            bool bottom = position == StatusStickerPosition.BottomLeft || position == StatusStickerPosition.BottomRight;
            group.Left = right ? slideWidth - group.Width - margin : margin;
            group.Top = bottom ? slideHeight - group.Height - margin : margin;

            for (int attempt = 0; attempt < 20; attempt++)
            {
                bool collision = false;
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var other = slide.Shapes[i];
                    if (other.Name == group.Name || !IsSticker(other)) continue;
                    if (Intersects(group, other)) { collision = true; break; }
                }
                if (!collision) break;
                group.Top += bottom ? -(group.Height + gap) : group.Height + gap;
            }
        }

        private static bool Intersects(Shape a, Shape b)
        {
            return a.Left < b.Left + b.Width && a.Left + a.Width > b.Left &&
                   a.Top < b.Top + b.Height && a.Top + a.Height > b.Top;
        }

        private static List<Slide> GetSlides(PowerPointManager manager, StatusStickerScope scope)
        {
            var result = new List<Slide>();
            if (scope == StatusStickerScope.CurrentSlide)
            {
                var slide = manager.GetCurrentSlide();
                if (slide != null) result.Add(slide);
            }
            else if (scope == StatusStickerScope.SelectedSlides)
            {
                result.AddRange(manager.GetSelectedSlidesForOperation());
                if (result.Count == 0 && manager.GetCurrentSlide() != null) result.Add(manager.GetCurrentSlide());
            }
            else
            {
                var presentation = manager.GetCurrentPresentation();
                if (presentation != null)
                    foreach (Slide slide in presentation.Slides) result.Add(slide);
            }
            return result.GroupBy(s => s.SlideID).Select(g => g.First()).ToList();
        }

        private static Shape FindSticker(Slide slide)
        {
            if (slide == null) return null;
            for (int i = 1; i <= slide.Shapes.Count; i++)
            {
                var shape = slide.Shapes[i];
                if (IsSticker(shape)) return shape;
            }
            return null;
        }

        private static bool IsSticker(Shape shape)
        {
            if (shape == null) return false;
            try
            {
                return string.Equals(shape.Tags[StickerTag], "1", StringComparison.OrdinalIgnoreCase) ||
                       shape.Name.StartsWith("oPE_StatusSticker_", StringComparison.OrdinalIgnoreCase) ||
                       shape.Name.StartsWith("StickyNote_Illustrative_", StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        private static StickerState ReadState(Shape shape)
        {
            if (!IsSticker(shape)) return null;
            var state = new StickerState
            {
                Text = GetTag(shape, TextTag),
                WorkStatus = GetTag(shape, WorkTag),
                Confidentiality = GetTag(shape, ConfidentialityTag),
                Audience = GetTag(shape, AudienceTag),
                Position = ParsePosition(GetTag(shape, PositionTag)),
                UseThemeColor = GetTag(shape, ThemeColorTag) == "1",
                IncludeDate = GetTag(shape, DateTag) == "1",
                IncludeAuthor = GetTag(shape, AuthorTag) == "1",
                ColorHex = GetTag(shape, ColorTag)
            };
            if (string.IsNullOrEmpty(state.Text) && shape.Type == Office.MsoShapeType.msoGroup)
            {
                for (int i = 1; i <= shape.GroupItems.Count; i++)
                {
                    var item = shape.GroupItems[i];
                    if (item.HasTextFrame == Office.MsoTriState.msoTrue && item.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        state.Text = item.TextFrame.TextRange.Text;
                        break;
                    }
                }
            }
            return state;
        }

        private static void WriteState(Shape shape, StickerState state)
        {
            SetTag(shape, StickerTag, "1");
            SetTag(shape, TextTag, state.Text);
            SetTag(shape, WorkTag, state.WorkStatus);
            SetTag(shape, ConfidentialityTag, state.Confidentiality);
            SetTag(shape, AudienceTag, state.Audience);
            SetTag(shape, PositionTag, state.Position.ToString());
            SetTag(shape, ThemeColorTag, state.UseThemeColor ? "1" : "0");
            SetTag(shape, DateTag, state.IncludeDate ? "1" : "0");
            SetTag(shape, AuthorTag, state.IncludeAuthor ? "1" : "0");
            SetTag(shape, ColorTag, state.ColorHex);
        }

        private static string BuildDisplayText(StickerState state)
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(state.Text)) parts.Add(state.Text.Trim());
            else
            {
                if (!string.IsNullOrWhiteSpace(state.WorkStatus)) parts.Add(state.WorkStatus);
                if (!string.IsNullOrWhiteSpace(state.Confidentiality)) parts.Add(state.Confidentiality);
                if (!string.IsNullOrWhiteSpace(state.Audience)) parts.Add(state.Audience);
            }
            if (state.IncludeDate) parts.Add(DateTime.Today.ToString("yyyy-MM-dd"));
            if (state.IncludeAuthor) parts.Add(Environment.UserName);
            return string.Join(" · ", parts);
        }

        private static string GetSemanticColor(StickerState state)
        {
            if (string.Equals(state.Confidentiality, "Strictly confidential", StringComparison.OrdinalIgnoreCase)) return "#991B1B";
            if (string.Equals(state.Confidentiality, "Confidential", StringComparison.OrdinalIgnoreCase)) return "#DC2626";
            if (string.Equals(state.WorkStatus, "Client-ready", StringComparison.OrdinalIgnoreCase)) return "#15803D";
            if (string.Equals(state.WorkStatus, "For decision", StringComparison.OrdinalIgnoreCase)) return "#D97706";
            if (string.Equals(state.WorkStatus, "Draft", StringComparison.OrdinalIgnoreCase)) return "#2563EB";
            if (string.Equals(state.Audience, "Internal", StringComparison.OrdinalIgnoreCase)) return "#4B5563";
            var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
            return string.IsNullOrWhiteSpace(config?.IllustrativeStickerColor) ? "#FF0000" : config.IllustrativeStickerColor;
        }

        private static int ResolveColor(PowerPointManager manager, StickerState state, SidebarConfig config)
        {
            if (state.UseThemeColor)
            {
                try
                {
                    return manager.GetCurrentPresentation().SlideMaster.Theme.ThemeColorScheme
                        .Colors(Office.MsoThemeColorSchemeIndex.msoThemeAccent1).RGB;
                }
                catch (Exception ex) { ExceptionLogger.Log(ex, "IllustrativeStickerFeature.ResolveThemeColor"); }
            }
            return ParseHexColor(string.IsNullOrWhiteSpace(state.ColorHex) ? config?.IllustrativeStickerColor : state.ColorHex);
        }

        private static int ParseHexColor(string hex)
        {
            try
            {
                string value = (hex ?? "#FF0000").Replace("#", "");
                int r = Convert.ToInt32(value.Substring(0, 2), 16);
                int g = Convert.ToInt32(value.Substring(2, 2), 16);
                int b = Convert.ToInt32(value.Substring(4, 2), 16);
                return r | (g << 8) | (b << 16);
            }
            catch { return 0x0000FF; }
        }

        private static StatusStickerPosition ParsePosition(string value)
        {
            return Enum.TryParse(value, true, out StatusStickerPosition position) ? position : StatusStickerPosition.TopRight;
        }

        private static IEnumerable<string> SplitLabels(string labels)
        {
            return (labels ?? "").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(label => label.Trim()).Where(label => label.Length > 0);
        }

        private static string GetTag(Shape shape, string name)
        {
            try { return shape.Tags[name] ?? ""; }
            catch { return ""; }
        }

        private static void SetTag(Shape shape, string name, string value)
        {
            shape.Tags.Add(name, value ?? "");
        }
    }
}
