using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.Services
{
    public partial class HtmlPresentationImportService
    {
        private void RenderGptDeckSlide(Slide slide, HtmlSlideModel model, SlideScale scale, bool useSourceBackground, bool useSourceStyling, string fontName)
        {
            string html = model.RawHtml ?? string.Empty;
            string gptFont = !string.IsNullOrWhiteSpace(fontName) ? "Aptos" : null;

            if (useSourceBackground)
            {
                RenderGptDeckBackground(slide, model, scale);
            }

            if (model.Classes.Any(c => string.Equals(c, "hero", StringComparison.OrdinalIgnoreCase)))
            {
                RenderGptHero(slide, html, scale, gptFont);
                return;
            }

            HtmlBlock timeline = FindBlocksByClass(html, "div", "timeline").FirstOrDefault();
            if (timeline != null)
            {
                RenderGptHeader(slide, html, scale, model.IsDark, gptFont);
                RenderGptTimeline(slide, timeline.InnerHtml, scale, gptFont);
                RenderGptSource(slide, html, scale, model.IsDark, gptFont);
                return;
            }

            if (html.IndexOf("<table", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                RenderGptHeader(slide, html, scale, model.IsDark, gptFont);
                float tableBottom = RenderGptTable(slide, html, scale, model.IsDark, gptFont);
                RenderGptLooseCardGrid(slide, html, new LayoutRect(MapX(54f, scale), tableBottom + ScaleY(22f, scale), ScaleX(792f, scale), ScaleY(145f, scale)), 2, model.IsDark, scale, gptFont);
                RenderGptSource(slide, html, scale, model.IsDark, gptFont);
                return;
            }

            HtmlBlock columns = FindBlocksByClass(html, "div", "cols").FirstOrDefault();
            if (columns != null)
            {
                // FindBlocks only tracks nesting of the searched tag. A column h2 can
                // therefore look like a direct slide h2; prefer the column's own header.
                bool columnProvidesHeader = FindDirectChildBlocks(columns.InnerHtml, "div")
                    .Any(column => !ClassContains(column, "photo") &&
                        (!string.IsNullOrWhiteSpace(ExtractFirstTagInnerHtml(column.InnerHtml, "h2")) ||
                         !string.IsNullOrWhiteSpace(ExtractBlockByClass(column.InnerHtml, "div", "kicker"))));
                bool hasDirectHeader = !columnProvidesHeader && FindDirectChildTagBlocks(html, "h2").Count > 0;
                if (hasDirectHeader) RenderGptHeader(slide, html, scale, model.IsDark, gptFont);
                RenderGptColumns(slide, columns.InnerHtml, scale, model.IsDark, hasDirectHeader, gptFont);
                RenderGptSource(slide, html, scale, model.IsDark, gptFont);
                return;
            }

            RenderGptHeader(slide, html, scale, model.IsDark, gptFont);
            RenderGptLooseCardGrid(slide, html, new LayoutRect(MapX(54f, scale), MapY(132f, scale), ScaleX(792f, scale), ScaleY(325f, scale)), 3, model.IsDark, scale, gptFont);
            RenderGptSource(slide, html, scale, model.IsDark, gptFont);
        }

        private void RenderGptDeckBackground(Slide slide, HtmlSlideModel model, SlideScale scale)
        {
            bool isHero = model.Classes.Any(c => string.Equals(c, "hero", StringComparison.OrdinalIgnoreCase));
            if (isHero)
            {
                Match heroRule = Regex.Match(model.StyleSheet ?? string.Empty, "\\.hero\\s*\\{(?<css>.*?)\\}", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                Match imageUrl = Regex.Match(heroRule.Groups["css"].Value, "url\\(\\s*['\"]?(?<url>data:image/.*?)['\"]?\\s*\\)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                string imagePath = ResolveMediaSource(imageUrl.Success ? imageUrl.Groups["url"].Value : null);
                if (!string.IsNullOrWhiteSpace(imagePath))
                {
                    Shape image = slide.Shapes.AddPicture(imagePath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 0f, 0f, ScaleX(HtmlSlideWidth, scale), ScaleY(HtmlSlideHeight, scale));
                    image.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                }

                Shape overlay = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0f, 0f, ScaleX(HtmlSlideWidth, scale), ScaleY(HtmlSlideHeight, scale));
                overlay.Line.Visible = Office.MsoTriState.msoFalse;
                overlay.Fill.ForeColor.RGB = ToRgb("#06111d");
                overlay.Fill.Transparency = 0.28f;
                return;
            }

            Shape background = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0f, 0f, ScaleX(HtmlSlideWidth, scale), ScaleY(HtmlSlideHeight, scale));
            background.Line.Visible = Office.MsoTriState.msoFalse;
            background.Fill.Visible = Office.MsoTriState.msoTrue;
            background.Fill.Transparency = 0f;
            if (model.IsDark)
            {
                background.Fill.ForeColor.RGB = ToRgb("#1f5e78");
                background.Fill.BackColor.RGB = ToRgb("#03070c");
                background.Fill.TwoColorGradient(Office.MsoGradientStyle.msoGradientFromCenter, 1);
                SetOpaqueGradientStops(background, "#1f5e78", "#03070c");
            }
            else
            {
                background.Fill.ForeColor.RGB = ToRgb("#f8fdff");
                background.Fill.BackColor.RGB = ToRgb("#d7edf6");
                background.Fill.TwoColorGradient(Office.MsoGradientStyle.msoGradientDiagonalUp, 1);
                SetOpaqueGradientStops(background, "#f8fdff", "#d7edf6");
            }
            background.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
        }

        private void SetOpaqueGradientStops(Shape shape, string startColor, string endColor)
        {
            try
            {
                int count = shape.Fill.GradientStops.Count;
                for (int index = 1; index <= count; index++)
                {
                    shape.Fill.GradientStops[index].Transparency = 0f;
                }
                if (count >= 1)
                {
                    shape.Fill.GradientStops[1].Color.RGB = ToRgb(startColor);
                    shape.Fill.GradientStops[1].Position = 0f;
                }
                if (count >= 2)
                {
                    shape.Fill.GradientStops[count].Color.RGB = ToRgb(endColor);
                    shape.Fill.GradientStops[count].Position = 1f;
                }
            }
            catch
            {
                // The native two-color fill remains available on older PowerPoint versions.
            }
        }

        private void RenderGptHero(Slide slide, string html, SlideScale scale, string fontName)
        {
            string kickerHtml = ExtractBlockByClass(html, "div", "kicker");
            string titleHtml = ExtractFirstTagInnerHtml(html, "h1");
            HtmlBlock subtitleBlock = FindBlocksByClass(html, "p", "subtitle").FirstOrDefault();
            string badgeHtml = ExtractBlockByClass(html, "div", "badge");
            float left = MapX(55f, scale);

            var kicker = AddTextBox(slide, left, MapY(132f, scale), ScaleX(700f, scale), ScaleY(24f, scale), CleanInlineText(kickerHtml).ToUpperInvariant(), 15f * scale.Average, "#f3b63f", true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
            var title = AddTextBox(slide, left, MapY(163f, scale), ScaleX(760f, scale), ScaleY(95f, scale), CleanInlineText(titleHtml), 64f * scale.Average, "#ffffff", true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
            var subtitle = AddTextBox(slide, left, MapY(265f, scale), ScaleX(650f, scale), ScaleY(75f, scale), CleanInlineText(subtitleBlock != null ? subtitleBlock.InnerHtml : string.Empty), 28f * scale.Average, "#dff3ff", false, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
            ApplyHtmlHyperlinks(kicker, kickerHtml);
            ApplyHtmlHyperlinks(title, titleHtml);
            ApplyHtmlHyperlinks(subtitle, subtitleBlock != null ? subtitleBlock.InnerHtml : null);

            if (!string.IsNullOrWhiteSpace(badgeHtml))
            {
                Shape badge = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left, MapY(365f, scale), ScaleX(315f, scale), ScaleY(42f, scale));
                badge.Line.Visible = Office.MsoTriState.msoFalse;
                badge.Fill.ForeColor.RGB = ToRgb("#ffffff");
                badge.Adjustments[1] = 0.45f;
                SetShapeText(badge, CleanInlineText(badgeHtml), 15f * scale.Average, "#0d5c83", true, PpParagraphAlignment.ppAlignCenter, Office.MsoVerticalAnchor.msoAnchorMiddle, fontName);
                ApplyHtmlHyperlinks(badge, badgeHtml);
            }
        }

        private float RenderGptHeader(Slide slide, string html, SlideScale scale, bool isDark, string fontName)
        {
            string kickerHtml = ExtractBlockByClass(html, "div", "kicker");
            string titleHtml = ExtractFirstTagInnerHtml(html, "h2");
            string kickerColor = isDark ? "#f3b63f" : "#0d5c83";
            string titleColor = isDark ? "#ffffff" : "#102033";
            AddTextBox(slide, MapX(54f, scale), MapY(28f, scale), ScaleX(790f, scale), ScaleY(22f, scale), CleanInlineText(kickerHtml).ToUpperInvariant(), 15f * scale.Average, kickerColor, true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
            var title = AddTextBox(slide, MapX(54f, scale), MapY(52f, scale), ScaleX(790f, scale), ScaleY(64f, scale), CleanInlineText(titleHtml), 41f * scale.Average, titleColor, true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
            ApplyHtmlHyperlinks(title, titleHtml);
            return MapY(120f, scale);
        }

        private void RenderGptSource(Slide slide, string html, SlideScale scale, bool isDark, string fontName)
        {
            string sourceHtml = ExtractBlockByClass(html, "div", "source");
            if (string.IsNullOrWhiteSpace(sourceHtml)) return;
            var source = AddTextBox(slide, MapX(54f, scale), MapY(530f, scale), ScaleX(790f, scale), ScaleY(18f, scale), CleanInlineText(sourceHtml), 12f * scale.Average, isDark ? "#c7d8e2" : "#718096", false, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorMiddle, fontName);
            ApplyHtmlHyperlinks(source, sourceHtml);
        }

        private void RenderGptColumns(Slide slide, string columnsHtml, SlideScale scale, bool isDark, bool hasDirectHeader, string fontName)
        {
            List<HtmlBlock> columns = FindDirectChildBlocks(columnsHtml, "div");
            if (columns.Count < 2) return;

            float left = MapX(54f, scale);
            float top = hasDirectHeader ? MapY(126f, scale) : MapY(38f, scale);
            float totalWidth = ScaleX(792f, scale);
            float gap = ScaleX(28f, scale);
            float firstWidth = (totalWidth - gap) * 1.05f / 2f;
            float secondWidth = totalWidth - gap - firstWidth;
            float height = hasDirectHeader ? ScaleY(375f, scale) : ScaleY(465f, scale);

            RenderGptColumn(slide, columns[0], new LayoutRect(left, top, firstWidth, height), isDark, scale, fontName);
            RenderGptColumn(slide, columns[1], new LayoutRect(left + firstWidth + gap, top, secondWidth, height), isDark, scale, fontName);
        }

        private void RenderGptColumn(Slide slide, HtmlBlock column, LayoutRect rect, bool isDark, SlideScale scale, string fontName)
        {
            if (ClassContains(column, "photo"))
            {
                RenderGptPhoto(slide, column.InnerHtml, rect);
                return;
            }

            // A .cols layout may use a card as its second column. Render the card as a
            // single component so headings, metrics, styling, and body text stay together.
            if (ClassContains(column, "card"))
            {
                RenderGptCard(slide, column, rect, isDark, scale, fontName);
                return;
            }

            float currentY = rect.Top;
            string kickerHtml = ExtractBlockByClass(column.InnerHtml, "div", "kicker");
            if (!string.IsNullOrWhiteSpace(kickerHtml))
            {
                AddTextBox(slide, rect.Left, currentY, rect.Width, ScaleY(22f, scale), CleanInlineText(kickerHtml).ToUpperInvariant(), 15f * scale.Average, isDark ? "#f3b63f" : "#0d5c83", true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                currentY += ScaleY(25f, scale);
            }

            string titleHtml = ExtractFirstTagInnerHtml(column.InnerHtml, "h2");
            if (!string.IsNullOrWhiteSpace(titleHtml))
            {
                var title = AddTextBox(slide, rect.Left, currentY, rect.Width, ScaleY(92f, scale), CleanInlineText(titleHtml), 38f * scale.Average, isDark ? "#ffffff" : "#102033", true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                title.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                ApplyHtmlHyperlinks(title, titleHtml);
                currentY += title.Height + ScaleY(14f, scale);
            }

            // FindBlocks only tracks nesting of matching tags. A paragraph inside a card can
            // therefore be mistaken for a column paragraph and rendered twice above the cards.
            foreach (HtmlDomNode paragraph in ParseDomFragment(column.InnerHtml).Children.Where(node => string.Equals(node.TagName, "p", StringComparison.OrdinalIgnoreCase)))
            {
                string sizeText = ReadStyle(paragraph.GetAttribute("style"), "font-size");
                float size = CssPixelsToPoints(20f);
                float parsed;
                if (TryParseCssPixels(sizeText, out parsed)) size = CssPixelsToPoints(parsed);
                var text = AddTextBox(slide, rect.Left, currentY, rect.Width, ScaleY(105f, scale), CleanInlineText(paragraph.InnerHtml), size * scale.Average, isDark ? "#c7d8e2" : "#102033", false, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                text.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                ApplyHtmlBoldFormatting(text, paragraph.InnerHtml);
                ApplyHtmlHyperlinks(text, paragraph.InnerHtml);
                currentY += text.Height + ScaleY(18f, scale);
            }

            HtmlBlock iconRow = FindBlocksByClass(column.InnerHtml, "div", "iconrow").FirstOrDefault();
            if (iconRow != null)
            {
                RenderGptLooseCardGrid(slide, iconRow.InnerHtml, new LayoutRect(rect.Left, currentY, rect.Width, Math.Max(0f, rect.Top + rect.Height - currentY)), 3, isDark, scale, fontName);
                return;
            }

            HtmlBlock nestedGrid = FindBlocksByClass(column.InnerHtml, "div", "grid").FirstOrDefault();
            if (nestedGrid != null)
            {
                RenderGptLooseCardGrid(slide, nestedGrid.InnerHtml, new LayoutRect(rect.Left, currentY, rect.Width, Math.Max(0f, rect.Top + rect.Height - currentY)), 1, isDark, scale, fontName);
                return;
            }

            List<HtmlBlock> directCards = FindDirectChildBlocks(column.InnerHtml, "div").Where(block => ClassContains(block, "card")).ToList();
            if (directCards.Count > 0)
            {
                RenderGptCardList(slide, directCards, new LayoutRect(rect.Left, currentY, rect.Width, Math.Max(0f, rect.Top + rect.Height - currentY)), 1, isDark, scale, fontName);
                return;
            }

            if (ContainsMedia(column.InnerHtml)) RenderGptPhoto(slide, column.InnerHtml, new LayoutRect(rect.Left, currentY, rect.Width, Math.Max(0f, rect.Top + rect.Height - currentY)));
        }

        private void RenderGptPhoto(Slide slide, string html, LayoutRect rect)
        {
            Match imageMatch = Regex.Match(html ?? string.Empty, "<img\\b[^>]*>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!imageMatch.Success) return;
            string path = ResolveMediaSource(ReadAttribute(imageMatch.Value, "src"));
            if (string.IsNullOrWhiteSpace(path)) return;
            try
            {
                Shape picture = slide.Shapes.AddPicture(path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, rect.Left, rect.Top, -1f, -1f);
                ApplyCoverCrop(picture, rect);
                ApplyRoundedPictureMask(picture);
                picture.AlternativeText = WebUtility.HtmlDecode(ReadAttribute(imageMatch.Value, "alt") ?? string.Empty);
                ApplyShapeHyperlink(picture, FindWrappingAnchorHref(html, imageMatch.Value));
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "HtmlPresentationImportService.RenderGptPhoto");
            }
        }

        private static void ApplyCoverCrop(Shape picture, LayoutRect target)
        {
            if (picture == null || target.Width <= 0f || target.Height <= 0f || picture.Width <= 0f || picture.Height <= 0f)
            {
                return;
            }

            float sourceWidth = picture.Width;
            float sourceHeight = picture.Height;
            float scale = Math.Max(target.Width / sourceWidth, target.Height / sourceHeight);
            float scaledWidth = sourceWidth * scale;
            float scaledHeight = sourceHeight * scale;

            try
            {
                Office.Crop crop = picture.PictureFormat.Crop;
                crop.ShapeLeft = target.Left;
                crop.ShapeTop = target.Top;
                crop.ShapeWidth = target.Width;
                crop.ShapeHeight = target.Height;
                crop.PictureWidth = scaledWidth;
                crop.PictureHeight = scaledHeight;
                crop.PictureOffsetX = 0f;
                crop.PictureOffsetY = 0f;
            }
            catch
            {
                // Older Office builds fall back to a proportional centered image.
                picture.LockAspectRatio = Office.MsoTriState.msoTrue;
                picture.Width = scaledWidth;
                picture.Height = scaledHeight;
                picture.Left = target.Left - (scaledWidth - target.Width) / 2f;
                picture.Top = target.Top - (scaledHeight - target.Height) / 2f;
            }
        }

        private void ApplyRoundedPictureMask(Shape picture)
        {
            if (picture == null) return;
            try
            {
                picture.AutoShapeType = Office.MsoAutoShapeType.msoShapeRoundedRectangle;
                picture.Adjustments[1] = 0.08f;
            }
            catch
            {
                // Crop-to-shape is unavailable on some older PowerPoint builds.
            }

            try
            {
                picture.Shadow.Visible = Office.MsoTriState.msoTrue;
                picture.Shadow.ForeColor.RGB = ToRgb("#06111d");
                picture.Shadow.Transparency = 0.76f;
                picture.Shadow.Blur = 18f;
                picture.Shadow.OffsetX = 0f;
                picture.Shadow.OffsetY = 8f;
            }
            catch
            {
                // The image remains usable if the host cannot apply a native shadow.
            }
        }

        private void RenderGptLooseCardGrid(Slide slide, string html, LayoutRect rect, int columns, bool isDark, SlideScale scale, string fontName)
        {
            List<HtmlBlock> cards = FindDirectChildBlocks(html, "div").Where(block => ClassContains(block, "card")).ToList();
            if (cards.Count == 0)
            {
                foreach (HtmlBlock grid in FindBlocksByClass(html, "div", "grid"))
                {
                    cards = FindDirectChildBlocks(grid.InnerHtml, "div").Where(block => ClassContains(block, "card")).ToList();
                    if (cards.Count > 0) break;
                }
            }
            RenderGptCardList(slide, cards, rect, columns, isDark, scale, fontName);
        }

        private void RenderGptCardList(Slide slide, List<HtmlBlock> cards, LayoutRect rect, int columns, bool isDark, SlideScale scale, string fontName)
        {
            if (cards == null || cards.Count == 0) return;
            columns = Math.Max(1, Math.Min(columns, cards.Count));
            int rows = (int)Math.Ceiling(cards.Count / (double)columns);
            float gapX = ScaleX(columns > 1 ? 18f : 0f, scale);
            float gapY = ScaleY(16f, scale);
            float cardWidth = (rect.Width - gapX * (columns - 1)) / columns;
            float cardHeight = Math.Max(GetGptCardMinimumHeight(cards, scale), (rect.Height - gapY * (rows - 1)) / rows);

            // A single card following narrative text is content-sized in HTML. Do not stretch
            // it to fill the rest of its column; doing so made callout cards such as
            // "Management-Insight" much taller than their HTML counterpart.
            if (columns == 1 && cards.Count == 1 && cards[0].Style.IndexOf("margin-top", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                cardHeight = Math.Min(cardHeight, ScaleY(112f, scale));
            }

            for (int index = 0; index < cards.Count; index++)
            {
                int row = index / columns;
                int column = index % columns;
                RenderGptCard(slide, cards[index], new LayoutRect(rect.Left + column * (cardWidth + gapX), rect.Top + row * (cardHeight + gapY), cardWidth, cardHeight), isDark, scale, fontName);
            }
        }

        private float GetGptCardMinimumHeight(IEnumerable<HtmlBlock> cards, SlideScale scale)
        {
            // Metric cards contain an icon, a frequently wrapped large value, and a caption.
            // Give their background enough room instead of sizing it only from the remaining column height.
            bool hasMetricCards = cards.Any(card =>
                FindTagBlocks(card.InnerHtml, "svg").Count > 0 &&
                !string.IsNullOrWhiteSpace(ExtractBlockByClass(card.InnerHtml, "div", "metric")));

            return ScaleY(hasMetricCards ? 250f : 82f, scale);
        }

        private void RenderGptCard(Slide slide, HtmlBlock card, LayoutRect rect, bool isDark, SlideScale scale, string fontName)
        {
            bool translucentDark = isDark || card.Style.IndexOf("rgba(255", StringComparison.OrdinalIgnoreCase) >= 0;
            Shape body = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, rect.Left, rect.Top, rect.Width, rect.Height);
            body.Line.Visible = Office.MsoTriState.msoTrue;
            body.Line.ForeColor.RGB = ToRgb(translucentDark ? "#476474" : "#d8e7ee");
            body.Line.Transparency = translucentDark ? 0.45f : 0.15f;
            body.Fill.ForeColor.RGB = ToRgb(translucentDark ? "#173341" : "#ffffff");
            body.Fill.Transparency = translucentDark ? 0.08f : 0.04f;
            body.Adjustments[1] = 0.12f;

            float left = rect.Left + ScaleX(20f, scale);
            float width = rect.Width - ScaleX(40f, scale);
            float currentY = rect.Top + ScaleY(18f, scale);
            HtmlBlock svg = FindTagBlocks(card.InnerHtml, "svg").FirstOrDefault();
            if (svg != null)
            {
                RenderGptInlineSvg(slide, svg, left, currentY, ScaleX(42f, scale), ScaleY(42f, scale));
                currentY += ScaleY(54f, scale);
            }

            string metricHtml = ExtractBlockByClass(card.InnerHtml, "div", "metric");
            if (!string.IsNullOrWhiteSpace(metricHtml))
            {
                bool yellow = card.InnerHtml.IndexOf("yellow", StringComparison.OrdinalIgnoreCase) >= 0;
                var metric = AddTextBox(slide, left, currentY, width, ScaleY(58f, scale), CleanInlineText(metricHtml), 34f * scale.Average, yellow ? "#f3b63f" : "#0d5c83", true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                metric.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                currentY += metric.Height + ScaleY(5f, scale);
            }

            List<HtmlBlock> headings = FindTagBlocks(card.InnerHtml, "h3");
            List<HtmlBlock> paragraphs = FindTagBlocks(card.InnerHtml, "p");
            if (headings.Count > 0)
            {
                var heading = AddTextBox(slide, left, currentY, width, ScaleY(32f, scale), CleanInlineText(headings[0].InnerHtml), 20f * scale.Average, translucentDark ? "#ffffff" : "#102033", true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                ApplyHtmlHyperlinks(heading, headings[0].InnerHtml);
                currentY += heading.Height + ScaleY(6f, scale);
            }

            HtmlBlock bar = FindBlocksByClass(card.InnerHtml, "div", "bar").FirstOrDefault();
            if (bar != null)
            {
                RenderGptBar(slide, bar, left, currentY, width, scale);
                currentY += ScaleY(27f, scale);
            }

            if (paragraphs.Count > 0)
            {
                var paragraph = AddTextBox(slide, left, currentY, width, Math.Max(ScaleY(38f, scale), rect.Top + rect.Height - ScaleY(16f, scale) - currentY), CleanInlineText(paragraphs[0].InnerHtml), 16f * scale.Average, translucentDark ? "#c7d8e2" : "#102033", false, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                paragraph.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                ApplyHtmlBoldFormatting(paragraph, paragraphs[0].InnerHtml);
                ApplyHtmlHyperlinks(paragraph, paragraphs[0].InnerHtml);
                currentY += paragraph.Height + ScaleY(10f, scale);
            }

            if (headings.Count > 1 && currentY < rect.Top + rect.Height - ScaleY(54f, scale))
            {
                var heading = AddTextBox(slide, left, currentY, width, ScaleY(30f, scale), CleanInlineText(headings[1].InnerHtml), 18f * scale.Average, translucentDark ? "#ffffff" : "#102033", true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                currentY += heading.Height + ScaleY(4f, scale);
                if (paragraphs.Count > 1)
                {
                    var paragraph = AddTextBox(slide, left, currentY, width, Math.Max(ScaleY(32f, scale), rect.Top + rect.Height - ScaleY(12f, scale) - currentY), CleanInlineText(paragraphs[1].InnerHtml), 15f * scale.Average, translucentDark ? "#c7d8e2" : "#102033", false, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                    ApplyHtmlHyperlinks(paragraph, paragraphs[1].InnerHtml);
                }
            }
        }

        private void RenderGptInlineSvg(Slide slide, HtmlBlock svg, float left, float top, float width, float height)
        {
            try
            {
                string svgHtml = Regex.Replace(svg.FullHtml, "<svg\\b", "<svg fill=\"none\" stroke=\"#0d5c83\" stroke-width=\"2.2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"", RegexOptions.IgnoreCase);
                string path = CreateTemporaryFile(".svg", Encoding.UTF8.GetBytes(svgHtml));
                slide.Shapes.AddPicture(path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, left, top, width, height);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "HtmlPresentationImportService.RenderGptInlineSvg");
            }
        }

        private void RenderGptBar(Slide slide, HtmlBlock bar, float left, float top, float width, SlideScale scale)
        {
            float height = ScaleY(14f, scale);
            Shape track = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left, top, width, height);
            track.Line.Visible = Office.MsoTriState.msoFalse;
            track.Fill.ForeColor.RGB = ToRgb("#d6e9f1");
            track.Adjustments[1] = 0.45f;

            HtmlBlock span = FindTagBlocks(bar.InnerHtml, "span").FirstOrDefault();
            float percentage = ReadCssDimension(span != null ? ReadStyle(span.Style, "width") : null, 1f, 1f);
            if (percentage <= 0f || percentage > 1f) percentage = 0.5f;
            Shape fill = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left, top, width * percentage, height);
            fill.Line.Visible = Office.MsoTriState.msoFalse;
            fill.Fill.ForeColor.RGB = ToRgb("#0d5c83");
            fill.Fill.BackColor.RGB = ToRgb("#f3b63f");
            fill.Fill.TwoColorGradient(Office.MsoGradientStyle.msoGradientHorizontal, 1);
            fill.Adjustments[1] = 0.45f;
        }

        private void RenderGptTimeline(Slide slide, string timelineHtml, SlideScale scale, string fontName)
        {
            List<HtmlBlock> steps = FindDirectChildBlocks(timelineHtml, "div").Where(block => ClassContains(block, "step")).ToList();
            if (steps.Count == 0) return;
            float left = MapX(54f, scale);
            float top = MapY(132f, scale);
            float totalWidth = ScaleX(792f, scale);
            float gap = ScaleX(16f, scale);
            float width = (totalWidth - gap * (steps.Count - 1)) / steps.Count;
            float height = ScaleY(300f, scale);

            for (int index = 0; index < steps.Count; index++)
            {
                float cardLeft = left + index * (width + gap);
                Shape body = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, cardLeft, top, width, height);
                body.Line.Visible = Office.MsoTriState.msoFalse;
                body.Fill.ForeColor.RGB = ToRgb("#ffffff");
                body.Adjustments[1] = 0.14f;

                float textLeft = cardLeft + ScaleX(15f, scale);
                float textWidth = width - ScaleX(30f, scale);
                string monthHtml = ExtractBlockByClass(steps[index].InnerHtml, "div", "month");
                string headingHtml = ExtractFirstTagInnerHtml(steps[index].InnerHtml, "h3");
                string paragraphHtml = ExtractFirstTagInnerHtml(steps[index].InnerHtml, "p");
                AddTextBox(slide, textLeft, top + ScaleY(18f, scale), textWidth, ScaleY(28f, scale), CleanInlineText(monthHtml), 16f * scale.Average, "#f3b63f", true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                AddTextBox(slide, textLeft, top + ScaleY(55f, scale), textWidth, ScaleY(52f, scale), CleanInlineText(headingHtml), 20f * scale.Average, "#102033", true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                var paragraph = AddTextBox(slide, textLeft, top + ScaleY(112f, scale), textWidth, height - ScaleY(130f, scale), CleanInlineText(paragraphHtml), 15f * scale.Average, "#102033", false, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                paragraph.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                ApplyHtmlHyperlinks(paragraph, paragraphHtml);
            }
        }

        private float RenderGptTable(Slide slide, string html, SlideScale scale, bool isDark, string fontName)
        {
            HtmlBlock tableBlock = FindTagBlocks(html, "table").FirstOrDefault();
            if (tableBlock == null) return MapY(120f, scale);
            List<HtmlBlock> rows = FindTagBlocks(tableBlock.InnerHtml, "tr");
            if (rows.Count == 0) return MapY(120f, scale);
            float left = MapX(54f, scale);
            float top = MapY(122f, scale);
            float width = ScaleX(792f, scale);
            float height = RenderStructuredTable(slide, rows, left, top, width, ScaleY(50f, scale), scale, false, fontName, "#06111d", "#ffffff");
            return top + height;
        }

        private void SetTableCellBorders(Cell cell, string colorHex)
        {
            int color = ToRgb(colorHex);
            PpBorderType[] borderTypes =
            {
                PpBorderType.ppBorderTop,
                PpBorderType.ppBorderBottom,
                PpBorderType.ppBorderLeft,
                PpBorderType.ppBorderRight
            };

            foreach (PpBorderType borderType in borderTypes)
            {
                var border = cell.Borders[borderType];
                border.Visible = Office.MsoTriState.msoTrue;
                border.ForeColor.RGB = color;
                border.Weight = 0.75f;
            }
        }

    }
}
