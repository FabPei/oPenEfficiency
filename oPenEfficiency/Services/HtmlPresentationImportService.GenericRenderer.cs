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
        private void RenderGenericDomSlide(Slide slide, HtmlSlideModel model, SlideScale scale, bool useSourceBackground, bool useSourceStyling, string fontName)
        {
            HtmlDomNode parsedRoot = ParseDomFragment(model.RawHtml);
            var slideNode = new HtmlDomNode
            {
                TagName = "section",
                OpenTag = model.SectionOpenTag,
                Attributes = ParseHtmlAttributes(model.SectionOpenTag),
                Classes = model.Classes ?? new List<string>(),
                InnerHtml = model.RawHtml,
                FullHtml = model.RawHtml
            };
            foreach (HtmlDomNode child in parsedRoot.Children)
            {
                child.Parent = slideNode;
                slideNode.Children.Add(child);
            }

            var resolver = new CssStyleResolver(model.StyleSheet);
            Dictionary<string, string> slideStyle = resolver.GetStyle(slideNode);
            if (useSourceBackground) RenderGenericCssBackground(slide, slideStyle, scale);

            float[] padding = ReadGenericBox(slideStyle, "padding", scale, ScaleX(48f, scale), ScaleY(34f, scale));
            var content = new LayoutRect(
                padding[3],
                padding[0],
                Math.Max(1f, ScaleX(HtmlSlideWidth, scale) - padding[1] - padding[3]),
                Math.Max(1f, ScaleY(HtmlSlideHeight, scale) - padding[0] - padding[2]));
            string resolvedFont = _presentationTheme != null
                ? _presentationTheme.BodyFont
                : (!string.IsNullOrWhiteSpace(fontName) ? GetGenericFontName(GetStyleValue(slideStyle, "font-family")) : null);
            RenderGenericChildren(slide, slideNode.Children, content, resolver, scale, resolvedFont, useSourceStyling);
        }

        private void RenderGenericCssBackground(Slide slide, Dictionary<string, string> style, SlideScale scale)
        {
            string background = GetStyleValue(style, "background") ?? GetStyleValue(style, "background-color") ?? "#ffffff";
            Match urlMatch = Regex.Match(background, "url\\(\\s*['\"]?(?<url>.*?)['\"]?\\s*\\)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (urlMatch.Success)
            {
                string path = ResolveMediaSource(urlMatch.Groups["url"].Value);
                if (!string.IsNullOrWhiteSpace(path))
                {
                    Shape image = slide.Shapes.AddPicture(path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 0f, 0f, -1f, -1f);
                    ApplyCoverCrop(image, new LayoutRect(0f, 0f, ScaleX(HtmlSlideWidth, scale), ScaleY(HtmlSlideHeight, scale)));
                    image.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                }
            }

            Shape backgroundShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0f, 0f, ScaleX(HtmlSlideWidth, scale), ScaleY(HtmlSlideHeight, scale));
            backgroundShape.Line.Visible = Office.MsoTriState.msoFalse;
            ApplyCssFill(backgroundShape, background, GetCssOpacity(style), "#ffffff");
            backgroundShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
        }

        private float RenderGenericChildren(Slide slide, IEnumerable<HtmlDomNode> nodes, LayoutRect area, CssStyleResolver resolver, SlideScale scale, string fontName, bool useSourceStyling)
        {
            float cursorY = area.Top;
            float maxBottom = area.Top;
            foreach (HtmlDomNode node in nodes.Where(node => node.TagName != "#text"))
            {
                Dictionary<string, string> style = resolver.GetStyle(node);
                if (string.Equals(GetStyleValue(style, "display"), "none", StringComparison.OrdinalIgnoreCase)) continue;
                bool absolute = string.Equals(GetStyleValue(style, "position"), "absolute", StringComparison.OrdinalIgnoreCase) || string.Equals(GetStyleValue(style, "position"), "fixed", StringComparison.OrdinalIgnoreCase);
                float[] margin = ReadGenericBox(style, "margin", scale, 0f, 0f);
                float left = area.Left + margin[3];
                float top = cursorY + margin[0];
                float width = Math.Max(1f, area.Width - margin[1] - margin[3]);
                string leftValue = GetPositionValue(style, "left", 3);
                string topValue = GetPositionValue(style, "top", 0);
                string rightValue = GetPositionValue(style, "right", 1);
                string bottomValue = GetPositionValue(style, "bottom", 2);
                bool hasLeft = !string.IsNullOrWhiteSpace(leftValue) && !string.Equals(leftValue, "auto", StringComparison.OrdinalIgnoreCase);
                bool hasTop = !string.IsNullOrWhiteSpace(topValue) && !string.Equals(topValue, "auto", StringComparison.OrdinalIgnoreCase);
                bool hasRight = !string.IsNullOrWhiteSpace(rightValue) && !string.Equals(rightValue, "auto", StringComparison.OrdinalIgnoreCase);
                bool hasBottom = !string.IsNullOrWhiteSpace(bottomValue) && !string.Equals(bottomValue, "auto", StringComparison.OrdinalIgnoreCase);
                string widthValue = GetStyleValue(style, "width");
                float explicitWidth = ReadGenericLength(widthValue, area.Width, scale, true, 16f);
                if (explicitWidth > 0f) width = Math.Min(area.Width, explicitWidth);
                float maxWidth = ReadGenericLength(GetStyleValue(style, "max-width"), area.Width, scale, true, 16f);
                if (maxWidth > 0f) width = Math.Min(width, maxWidth);
                string heightValue = GetStyleValue(style, "height") ?? GetStyleValue(style, "min-height");
                float explicitHeight = ReadGenericLength(heightValue, area.Height, scale, false, 16f);
                float maxHeight = ReadGenericLength(GetStyleValue(style, "max-height"), area.Height, scale, false, 16f);
                if (maxHeight > 0f && explicitHeight > 0f) explicitHeight = Math.Min(explicitHeight, maxHeight);

                if (absolute)
                {
                    float cssLeft = ReadGenericLength(leftValue, area.Width, scale, true, 16f);
                    float cssTop = ReadGenericLength(topValue, area.Height, scale, false, 16f);
                    float cssRight = ReadGenericLength(rightValue, area.Width, scale, true, 16f);
                    float cssBottom = ReadGenericLength(bottomValue, area.Height, scale, false, 16f);
                    if (hasLeft) left = area.Left + cssLeft;
                    if (hasTop) top = area.Top + cssTop;
                    if (hasLeft && hasRight && explicitWidth <= 0f) width = Math.Max(1f, area.Width - cssLeft - cssRight);
                    if (hasTop && hasBottom && explicitHeight <= 0f) explicitHeight = Math.Max(1f, area.Height - cssTop - cssBottom);
                    if (!hasLeft && hasRight && explicitWidth > 0f) left = area.Left + area.Width - cssRight - width;
                    if (!hasTop && hasBottom && explicitHeight > 0f) top = area.Top + area.Height - cssBottom - explicitHeight;
                }
                else if (explicitWidth > 0f && IsAutoHorizontalMargin(style))
                {
                    left = area.Left + (area.Width - width) / 2f;
                }

                float availableHeight = explicitHeight > 0f ? explicitHeight : Math.Max(ScaleY(30f, scale), area.Top + area.Height - top - margin[2]);
                int firstShapeIndex = slide.Shapes.Count + 1;
                float used = RenderGenericNode(slide, node, new LayoutRect(left, top, width, availableHeight), resolver, scale, fontName, useSourceStyling);
                float renderedHeight = Math.Max(used, explicitHeight);
                float shiftX = 0f;
                float shiftY = 0f;
                if (absolute && !hasLeft && hasRight) shiftX = area.Left + area.Width - ReadGenericLength(rightValue, area.Width, scale, true, 16f) - width - left;
                if (absolute && !hasTop && hasBottom) shiftY = area.Top + area.Height - ReadGenericLength(bottomValue, area.Height, scale, false, 16f) - renderedHeight - top;
                ApplyGenericShapeLayout(slide, firstShapeIndex, style, shiftX, shiftY, width, renderedHeight, scale);

                float renderedBottom = top + shiftY + renderedHeight + margin[2];
                maxBottom = Math.Max(maxBottom, renderedBottom);
                if (!absolute) cursorY = renderedBottom;
            }
            return Math.Max(0f, maxBottom - area.Top);
        }

        private float RenderGenericNode(Slide slide, HtmlDomNode node, LayoutRect area, CssStyleResolver resolver, SlideScale scale, string fontName, bool useSourceStyling)
        {
            Dictionary<string, string> style = resolver.GetStyle(node);
            switch (_renderRouter.GetNodeRenderer(node.TagName, GetStyleValue(style, "display")))
            {
                case HtmlNodeRendererKind.Text:
                    return RenderGenericTextNode(slide, node, area, style, scale, fontName);
                case HtmlNodeRendererKind.List:
                    return RenderGenericList(slide, node, area, style, scale, fontName);
                case HtmlNodeRendererKind.Image:
                    return RenderGenericImage(slide, node, area, style, scale);
                case HtmlNodeRendererKind.Svg:
                    return RenderGenericSvg(slide, node, area, style, scale);
                case HtmlNodeRendererKind.Media:
                    return RenderGenericMediaFallback(slide, node, area, style, scale, fontName);
                case HtmlNodeRendererKind.Table:
                    return RenderHtmlTables(slide, node.FullHtml, area.Left, area.Top, area.Width, scale, false, useSourceStyling, fontName);
                case HtmlNodeRendererKind.Rule:
                    Shape line = slide.Shapes.AddLine(area.Left, area.Top, area.Left + area.Width, area.Top);
                    line.Line.ForeColor.RGB = ToRgb(CssColorToHex(GetStyleValue(style, "border-color"), "#d7dde5"));
                    return ScaleY(8f, scale);
                case HtmlNodeRendererKind.Grid:
                case HtmlNodeRendererKind.Flex:
                    return RenderGenericContainer(slide, node, area, style, resolver, scale, fontName, useSourceStyling);
                default:
                    return RenderGenericContainer(slide, node, area, style, resolver, scale, fontName, useSourceStyling);
            }
        }

        private float RenderGenericTextNode(Slide slide, HtmlDomNode node, LayoutRect area, Dictionary<string, string> style, SlideScale scale, string fontName)
        {
            float defaultSize = node.TagName == "h1" ? 44f : node.TagName == "h2" ? 34f : node.TagName == "h3" ? 26f : node.TagName == "h4" ? 20f : 16f;
            float fontSize = ReadGenericFontSize(GetStyleValue(style, "font-size"), defaultSize);
            bool bold = node.TagName.StartsWith("h", StringComparison.OrdinalIgnoreCase) || IsCssBold(GetStyleValue(style, "font-weight"));
            string text = CleanInlineText(node.InnerHtml);
            string transform = GetStyleValue(style, "text-transform");
            if (string.Equals(transform, "uppercase", StringComparison.OrdinalIgnoreCase)) text = text.ToUpperInvariant();
            PpParagraphAlignment alignment = GetParagraphAlignment(GetStyleValue(style, "text-align"));
            string color = CssColorToHex(GetStyleValue(style, "color"), "#2c3e50");
            var shape = AddTextBox(slide, area.Left, area.Top, area.Width, Math.Min(area.Height, ScaleY(120f, scale)), text, fontSize * scale.Average, color, bold, alignment, Office.MsoVerticalAnchor.msoAnchorTop, GetGenericFontName(GetStyleValue(style, "font-family")) ?? fontName);
            shape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            ApplyCssTextAppearance(shape, style);
            if (string.Equals(GetStyleValue(style, "font-style"), "italic", StringComparison.OrdinalIgnoreCase)) shape.TextFrame.TextRange.Font.Italic = Office.MsoTriState.msoTrue;
            ApplyHtmlBoldFormatting(shape, node.InnerHtml);
            ApplyHtmlItalicFormatting(shape, node.InnerHtml);
            ApplyHtmlHyperlinks(shape, node.InnerHtml);
            return shape.Height;
        }

        private float RenderGenericList(Slide slide, HtmlDomNode node, LayoutRect area, Dictionary<string, string> style, SlideScale scale, string fontName)
        {
            float currentY = area.Top;
            int index = 1;
            foreach (HtmlDomNode item in node.Children.Where(child => child.TagName == "li"))
            {
                Dictionary<string, string> itemStyle = style;
                string marker = node.TagName == "ol" ? index.ToString() + "." : "•";
                float fontSize = ReadGenericFontSize(GetStyleValue(itemStyle, "font-size"), 16f);
                AddTextBox(slide, area.Left, currentY, ScaleX(24f, scale), ScaleY(28f, scale), marker, fontSize * scale.Average, CssColorToHex(GetStyleValue(itemStyle, "color"), "#2c3e50"), true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                var text = AddTextBox(slide, area.Left + ScaleX(28f, scale), currentY, area.Width - ScaleX(28f, scale), ScaleY(60f, scale), CleanInlineText(item.InnerHtml), fontSize * scale.Average, CssColorToHex(GetStyleValue(itemStyle, "color"), "#2c3e50"), false, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                text.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                ApplyHtmlBoldFormatting(text, item.InnerHtml);
                ApplyHtmlItalicFormatting(text, item.InnerHtml);
                ApplyHtmlHyperlinks(text, item.InnerHtml);
                currentY += text.Height + ScaleY(8f, scale);
                index++;
            }
            return currentY - area.Top;
        }

        private float RenderGenericImage(Slide slide, HtmlDomNode node, LayoutRect area, Dictionary<string, string> style, SlideScale scale)
        {
            string path = ResolveMediaSource(GetGenericImageSource(node));
            if (string.IsNullOrWhiteSpace(path)) return 0f;
            float height = ReadGenericLength(GetStyleValue(style, "height") ?? node.GetAttribute("height"), area.Height, scale, false, 16f);
            if (height <= 0f) height = Math.Min(area.Height, ScaleY(220f, scale));
            float width = ReadGenericLength(GetStyleValue(style, "width") ?? node.GetAttribute("width"), area.Width, scale, true, 16f);
            if (width <= 0f) width = area.Width;
            try
            {
                Shape picture = slide.Shapes.AddPicture(path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, area.Left, area.Top, -1f, -1f);
                var target = new LayoutRect(area.Left, area.Top, width, height);
                if (string.Equals(GetStyleValue(style, "object-fit"), "cover", StringComparison.OrdinalIgnoreCase)) ApplyCoverCrop(picture, target);
                else
                {
                    picture.LockAspectRatio = Office.MsoTriState.msoTrue;
                    float factor = Math.Min(width / picture.Width, height / picture.Height);
                    picture.Width *= factor;
                    picture.Height *= factor;
                }
                if (ReadGenericLength(GetStyleValue(style, "border-radius"), width, scale, true, 16f) > 0f) ApplyRoundedPictureMask(picture);
                picture.AlternativeText = node.GetAttribute("alt") ?? string.Empty;
                ApplyShapeHyperlink(picture, node.Parent != null && node.Parent.TagName == "a" ? node.Parent.GetAttribute("href") : null);
                return height;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "HtmlPresentationImportService.RenderGenericImage");
                return 0f;
            }
        }

        private static string GetGenericImageSource(HtmlDomNode node)
        {
            string source = node.GetAttribute("src") ?? node.GetAttribute("data-src");
            if (!string.IsNullOrWhiteSpace(source)) return source;
            string srcSet = node.GetAttribute("srcset") ?? node.GetAttribute("data-srcset");
            if (string.IsNullOrWhiteSpace(srcSet)) return null;
            string candidate = srcSet.Split(',').LastOrDefault();
            return string.IsNullOrWhiteSpace(candidate) ? null : Regex.Split(candidate.Trim(), "\\s+")[0];
        }

        private float RenderGenericSvg(Slide slide, HtmlDomNode node, LayoutRect area, Dictionary<string, string> style, SlideScale scale)
        {
            float width = ReadGenericLength(GetStyleValue(style, "width"), area.Width, scale, true, 16f);
            float height = ReadGenericLength(GetStyleValue(style, "height"), area.Height, scale, false, 16f);
            if (width <= 0f) width = Math.Min(area.Width, ScaleX(64f, scale));
            if (height <= 0f) height = ScaleY(64f, scale);
            try
            {
                string path = CreateTemporaryFile(".svg", Encoding.UTF8.GetBytes(node.FullHtml));
                slide.Shapes.AddPicture(path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, area.Left, area.Top, width, height);
                return height;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "HtmlPresentationImportService.RenderGenericSvg");
                return 0f;
            }
        }

        private float RenderGenericMediaFallback(Slide slide, HtmlDomNode node, LayoutRect area, Dictionary<string, string> style, SlideScale scale, string fontName)
        {
            string source = node.GetAttribute("src") ?? node.GetAttribute("data-src") ?? node.GetAttribute("href");
            string poster = node.GetAttribute("poster") ?? node.GetAttribute("data-poster") ?? node.GetAttribute("data-fallback") ?? node.GetAttribute("data-image");
            float height = ReadGenericLength(GetStyleValue(style, "height"), area.Height, scale, false, 16f);
            if (height <= 0f) height = Math.Min(area.Height, ScaleY(node.TagName == "audio" ? 56f : 180f, scale));
            float width = ReadGenericLength(GetStyleValue(style, "width"), area.Width, scale, true, 16f);
            if (width <= 0f) width = area.Width;
            var rect = new LayoutRect(area.Left, area.Top, width, height);

            string posterPath = ResolveMediaSource(poster);
            if (!string.IsNullOrWhiteSpace(posterPath))
            {
                try
                {
                    Shape picture = slide.Shapes.AddPicture(posterPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, rect.Left, rect.Top, -1f, -1f);
                    ApplyCoverCrop(picture, rect);
                    ApplyShapeHyperlink(picture, source);
                    return height;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "HtmlPresentationImportService.RenderGenericMediaFallback.Poster");
                }
            }

            Shape placeholder = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, rect.Left, rect.Top, rect.Width, rect.Height);
            placeholder.Fill.ForeColor.RGB = ToRgb(node.TagName == "audio" ? "#eef3f8" : "#1f3846");
            placeholder.Line.ForeColor.RGB = ToRgb("#94a3b8");
            placeholder.Adjustments[1] = 0.05f;
            string icon = node.TagName == "video" ? "▶" : node.TagName == "audio" ? "♪" : node.TagName == "canvas" ? "◫" : "↗";
            string label = node.TagName == "iframe" ? "Embedded content" : char.ToUpperInvariant(node.TagName[0]) + node.TagName.Substring(1);
            SetShapeText(placeholder, icon + "  " + label, 16f * scale.Average, node.TagName == "audio" ? "#2c3e50" : "#ffffff", true, PpParagraphAlignment.ppAlignCenter, Office.MsoVerticalAnchor.msoAnchorMiddle, fontName);
            ApplyShapeHyperlink(placeholder, source);
            AddWarning(label + " was imported as a clickable static placeholder.", null);
            return height;
        }

        private float RenderGenericContainer(Slide slide, HtmlDomNode node, LayoutRect area, Dictionary<string, string> style, CssStyleResolver resolver, SlideScale scale, string fontName, bool useSourceStyling)
        {
            string display = GetStyleValue(style, "display") ?? "block";
            float[] padding = ReadGenericBox(style, "padding", scale, 0f, 0f);
            float explicitHeight = ReadGenericLength(GetStyleValue(style, "height") ?? GetStyleValue(style, "min-height"), area.Height, scale, false, 16f);
            float containerHeight = explicitHeight > 0f ? explicitHeight : area.Height;
            Shape containerShape = RenderGenericContainerBackground(slide, area.Left, area.Top, area.Width, containerHeight, style, scale);
            var inner = new LayoutRect(area.Left + padding[3], area.Top + padding[0], Math.Max(1f, area.Width - padding[1] - padding[3]), Math.Max(1f, containerHeight - padding[0] - padding[2]));
            float used;
            if (display.IndexOf("grid", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                used = RenderGenericGrid(slide, node, inner, style, resolver, scale, fontName, useSourceStyling);
            }
            else if (string.Equals(display, "flex", StringComparison.OrdinalIgnoreCase) || string.Equals(display, "inline-flex", StringComparison.OrdinalIgnoreCase))
            {
                used = RenderGenericFlex(slide, node, inner, style, resolver, scale, fontName, useSourceStyling);
            }
            else
            {
                used = RenderGenericChildren(slide, node.Children, inner, resolver, scale, fontName, useSourceStyling);
                if (used <= 0f && !string.IsNullOrWhiteSpace(node.GetText())) used = RenderGenericTextNode(slide, node, inner, style, scale, fontName);
            }
            float finalHeight = explicitHeight > 0f ? explicitHeight : Math.Max(used + padding[0] + padding[2], ScaleY(1f, scale));
            if (containerShape != null) containerShape.Height = finalHeight;
            return finalHeight;
        }

        private float RenderGenericGrid(Slide slide, HtmlDomNode node, LayoutRect area, Dictionary<string, string> style, CssStyleResolver resolver, SlideScale scale, string fontName, bool useSourceStyling)
        {
            List<HtmlDomNode> children = node.Children.Where(child => child.TagName != "#text").ToList();
            if (children.Count == 0) return 0f;
            int columns = GetGenericColumnCount(style, children.Count, area.Width, scale);
            float gap = ReadGenericLength(GetStyleValue(style, "gap"), area.Width, scale, true, 16f);
            if (gap <= 0f) gap = ScaleX(12f, scale);
            var cells = new HtmlLayoutEngine().Grid(
                new HtmlLayoutRect(area.Left, area.Top, area.Width, area.Height), children.Count, columns, gap, gap);
            float maxBottom = area.Top;
            for (int index = 0; index < children.Count; index++)
            {
                HtmlLayoutRect cell = cells[index];
                float cellHeight = Math.Max(ScaleY(50f, scale), cell.Height);
                float used = RenderGenericNode(slide, children[index], new LayoutRect(cell.Left, cell.Top, Math.Max(1f, cell.Width), cellHeight), resolver, scale, fontName, useSourceStyling);
                maxBottom = Math.Max(maxBottom, cell.Top + Math.Max(used, cellHeight));
            }
            return maxBottom - area.Top;
        }

        private Shape RenderGenericContainerBackground(Slide slide, float left, float top, float width, float height, Dictionary<string, string> style, SlideScale scale)
        {
            string background = GetStyleValue(style, "background-color") ?? GetStyleValue(style, "background");
            string border = GetStyleValue(style, "border-color") ?? GetStyleValue(style, "border");
            string backgroundImage = GetStyleValue(style, "background-image");
            if (string.IsNullOrWhiteSpace(background) && string.IsNullOrWhiteSpace(border) && string.IsNullOrWhiteSpace(backgroundImage) && string.IsNullOrWhiteSpace(GetStyleValue(style, "box-shadow"))) return null;
            Shape shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left, top, width, Math.Max(1f, height));
            shape.Fill.Visible = string.IsNullOrWhiteSpace(background) || string.Equals(background, "transparent", StringComparison.OrdinalIgnoreCase) ? Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;
            if (shape.Fill.Visible == Office.MsoTriState.msoTrue) ApplyCssFill(shape, background, GetCssOpacity(style), "#ffffff");
            if (!string.IsNullOrWhiteSpace(backgroundImage) && backgroundImage.IndexOf("url(", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Match imageMatch = Regex.Match(backgroundImage, "url\\(\\s*['\"]?(?<url>.*?)['\"]?\\s*\\)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                string imagePath = imageMatch.Success ? ResolveMediaSource(imageMatch.Groups["url"].Value) : null;
                if (!string.IsNullOrWhiteSpace(imagePath))
                {
                    try { shape.Fill.UserPicture(imagePath); } catch { AddWarning("The background image could not be applied as a shape fill.", null); }
                }
            }
            shape.Line.Visible = string.IsNullOrWhiteSpace(border) || string.Equals(border, "none", StringComparison.OrdinalIgnoreCase) ? Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;
            if (shape.Line.Visible == Office.MsoTriState.msoTrue) ApplyCssBorder(shape, border, style, scale);
            float radius = ReadGenericLength(GetStyleValue(style, "border-radius"), width, scale, true, 16f);
            shape.Adjustments[1] = radius > 0f ? Math.Min(0.3f, radius / Math.Max(1f, Math.Min(width, height))) : 0.01f;
            ApplyCssShadow(shape, GetStyleValue(style, "box-shadow"));
            shape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
            return shape;
        }

        private void ApplyCssFill(Shape shape, string cssValue, float opacity, string fallback)
        {
            List<string> colors = ExtractCssColors(cssValue);
            string start = colors.Count > 0 ? colors[0] : CssColorToHex(cssValue, fallback);
            string end = colors.Count > 1 ? colors[colors.Count - 1] : start;
            shape.Fill.Visible = Office.MsoTriState.msoTrue;
            shape.Fill.ForeColor.RGB = ToRgb(start);
            shape.Fill.Transparency = Math.Max(0f, Math.Min(100f, 100f * (1f - opacity * GetCssColorOpacity(cssValue))));
            if (colors.Count <= 1 || cssValue.IndexOf("gradient", StringComparison.OrdinalIgnoreCase) < 0)
            {
                shape.Fill.Solid();
                return;
            }

            shape.Fill.BackColor.RGB = ToRgb(end);
            shape.Fill.TwoColorGradient(GetCssGradientDirection(cssValue), 1);
            SetOpaqueGradientStops(shape, start, end);
            try
            {
                dynamic stops = shape.Fill.GradientStops;
                for (int index = 1; index < colors.Count - 1; index++)
                {
                    dynamic stop = stops.Insert(ToRgb(colors[index]), index / (float)(colors.Count - 1));
                    stop.Transparency = 0f;
                }
            }
            catch
            {
                // Native PowerPoint gradients remain usable if a version cannot add extra stops.
            }
        }

        private static List<string> ExtractCssColors(string value)
        {
            return Regex.Matches(value ?? string.Empty, "#[0-9a-fA-F]{3,8}|rgba?\\([^)]*\\)")
                .Cast<Match>().Select(match => CssColorToHex(match.Value, "#ffffff")).ToList();
        }

        private static Office.MsoGradientStyle GetCssGradientDirection(string cssValue)
        {
            Match angleMatch = Regex.Match(cssValue ?? string.Empty, "(?:linear-gradient\\(\\s*)?(?<angle>-?[0-9.]+)deg", RegexOptions.IgnoreCase);
            if (angleMatch.Success)
            {
                float angle;
                if (float.TryParse(angleMatch.Groups["angle"].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out angle))
                {
                    angle = (angle % 360f + 360f) % 360f;
                    if (angle >= 45f && angle < 135f) return Office.MsoGradientStyle.msoGradientHorizontal;
                    if (angle >= 135f && angle < 225f) return Office.MsoGradientStyle.msoGradientDiagonalDown;
                    if (angle >= 225f && angle < 315f) return Office.MsoGradientStyle.msoGradientVertical;
                    return Office.MsoGradientStyle.msoGradientDiagonalUp;
                }
            }
            if ((cssValue ?? string.Empty).IndexOf("to right", StringComparison.OrdinalIgnoreCase) >= 0) return Office.MsoGradientStyle.msoGradientHorizontal;
            if ((cssValue ?? string.Empty).IndexOf("to bottom", StringComparison.OrdinalIgnoreCase) >= 0) return Office.MsoGradientStyle.msoGradientVertical;
            return Office.MsoGradientStyle.msoGradientDiagonalUp;
        }

        private static float GetCssOpacity(Dictionary<string, string> style)
        {
            return Math.Max(0f, Math.Min(1f, ReadCssFloat(GetStyleValue(style, "opacity"), 1f)));
        }

        private static float GetCssColorOpacity(string cssValue)
        {
            Match rgba = Regex.Match(cssValue ?? string.Empty, "rgba\\([^)]*,\\s*(?<opacity>[0-9.]+)\\s*\\)", RegexOptions.IgnoreCase);
            return rgba.Success ? Math.Max(0f, Math.Min(1f, ReadCssFloat(rgba.Groups["opacity"].Value, 1f))) : 1f;
        }

        private void ApplyCssBorder(Shape shape, string border, Dictionary<string, string> style, SlideScale scale)
        {
            shape.Line.ForeColor.RGB = ToRgb(CssColorToHex(border, "#d7dde5"));
            Match width = Regex.Match(border ?? string.Empty, "(?<width>[0-9.]+)(px|pt)", RegexOptions.IgnoreCase);
            if (width.Success) shape.Line.Weight = Math.Max(0.25f, ReadGenericLength(width.Groups["width"].Value, HtmlSlideWidth, scale, true, 16f));
            string styleValue = GetStyleValue(style, "border-style") ?? border;
            if (styleValue.IndexOf("dashed", StringComparison.OrdinalIgnoreCase) >= 0) shape.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;
            else if (styleValue.IndexOf("dotted", StringComparison.OrdinalIgnoreCase) >= 0) shape.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
        }

        private void ApplyCssShadow(Shape shape, string cssShadow)
        {
            if (string.IsNullOrWhiteSpace(cssShadow) || string.Equals(cssShadow, "none", StringComparison.OrdinalIgnoreCase)) return;
            Match shadow = Regex.Match(cssShadow, "(?<x>-?[0-9.]+)px\\s+(?<y>-?[0-9.]+)px(?:\\s+(?<blur>[0-9.]+)px)?(?:\\s+(?<color>#[0-9a-fA-F]{3,8}|rgba?\\([^)]*\\)))?", RegexOptions.IgnoreCase);
            if (!shadow.Success) return;
            shape.Shadow.Visible = Office.MsoTriState.msoTrue;
            shape.Shadow.OffsetX = ReadCssFloat(shadow.Groups["x"].Value, 0f);
            shape.Shadow.OffsetY = ReadCssFloat(shadow.Groups["y"].Value, 0f);
            shape.Shadow.Blur = ReadCssFloat(shadow.Groups["blur"].Value, 4f);
            string color = shadow.Groups["color"].Value;
            shape.Shadow.ForeColor.RGB = ToRgb(CssColorToHex(color, "#000000"));
            shape.Shadow.Transparency = 100f * (1f - GetCssColorOpacity(color));
        }

        private void ApplyCssTextAppearance(Shape shape, Dictionary<string, string> style)
        {
            string background = GetStyleValue(style, "background") ?? GetStyleValue(style, "background-image");
            if (!string.IsNullOrWhiteSpace(background) && background.IndexOf("gradient", StringComparison.OrdinalIgnoreCase) >= 0 &&
                (string.Equals(GetStyleValue(style, "-webkit-text-fill-color"), "transparent", StringComparison.OrdinalIgnoreCase) || GetStyleValue(style, "background-clip") == "text"))
            {
                try
                {
                    dynamic fill = shape.TextFrame2.TextRange.Font.Fill;
                    List<string> colors = ExtractCssColors(background);
                    if (colors.Count >= 2)
                    {
                        fill.Visible = Office.MsoTriState.msoTrue;
                        fill.TwoColorGradient(GetCssGradientDirection(background), 1);
                        fill.GradientStops[1].Color.RGB = ToRgb(colors[0]);
                        fill.GradientStops[1].Transparency = 0f;
                        fill.GradientStops[fill.GradientStops.Count].Color.RGB = ToRgb(colors[colors.Count - 1]);
                        fill.GradientStops[fill.GradientStops.Count].Transparency = 0f;
                    }
                }
                catch (Exception ex) { ReportOptionalImportFailure(ex, "HtmlPresentationImportService.ApplyTextRunFormatting"); }
            }
        }

        private static int GetGenericColumnCount(Dictionary<string, string> style, int childCount, float availableWidth, SlideScale scale)
        {
            string template = GetStyleValue(style, "grid-template-columns");
            Match repeat = Regex.Match(template ?? string.Empty, "repeat\\(\\s*(\\d+)", RegexOptions.IgnoreCase);
            int parsed;
            if (repeat.Success && int.TryParse(repeat.Groups[1].Value, out parsed)) return Math.Max(1, parsed);
            Match minMax = Regex.Match(template ?? string.Empty, "repeat\\(\\s*(?:auto-fit|auto-fill)\\s*,\\s*minmax\\(\\s*(?<minimum>[^,]+)", RegexOptions.IgnoreCase);
            if (minMax.Success)
            {
                float minimum = ReadGenericLength(minMax.Groups["minimum"].Value, availableWidth, scale, true, 16f);
                if (minimum > 0f) return Math.Max(1, Math.Min(childCount, (int)Math.Floor(availableWidth / minimum)));
            }
            if (!string.IsNullOrWhiteSpace(template)) return Math.Max(1, Regex.Split(template.Trim(), "\\s+").Length);
            return Math.Max(1, Math.Min(childCount, 3));
        }

        private float RenderGenericFlex(Slide slide, HtmlDomNode node, LayoutRect area, Dictionary<string, string> style, CssStyleResolver resolver, SlideScale scale, string fontName, bool useSourceStyling)
        {
            List<HtmlDomNode> children = node.Children.Where(child => child.TagName != "#text").OrderBy(child => GetCssOrder(resolver.GetStyle(child))).ToList();
            if (children.Count == 0) return 0f;

            bool column = string.Equals(GetStyleValue(style, "flex-direction"), "column", StringComparison.OrdinalIgnoreCase) || string.Equals(GetStyleValue(style, "flex-direction"), "column-reverse", StringComparison.OrdinalIgnoreCase);
            bool wrap = GetStyleValue(style, "flex-wrap") != null && !string.Equals(GetStyleValue(style, "flex-wrap"), "nowrap", StringComparison.OrdinalIgnoreCase);
            float gap = ReadGenericLength(GetStyleValue(style, "gap") ?? GetStyleValue(style, column ? "row-gap" : "column-gap"), column ? area.Height : area.Width, scale, !column, 16f);
            string alignment = GetStyleValue(style, "align-items") ?? "stretch";
            string justify = GetStyleValue(style, "justify-content") ?? "flex-start";
            float mainSize = column ? area.Height : area.Width;
            float crossSize = column ? area.Width : area.Height;
            var lines = new List<List<FlexItem>>();
            var currentLine = new List<FlexItem>();
            float usedMain = 0f;

            foreach (HtmlDomNode child in children)
            {
                Dictionary<string, string> childStyle = resolver.GetStyle(child);
                float requestedMain = ReadGenericLength(GetStyleValue(childStyle, column ? "height" : "width") ?? GetStyleValue(childStyle, "flex-basis"), mainSize, scale, !column, 16f);
                if (requestedMain <= 0f) requestedMain = wrap ? Math.Max(ScaleX(120f, scale), mainSize / Math.Min(children.Count, 3f)) : 0f;
                float requestedCross = ReadGenericLength(GetStyleValue(childStyle, column ? "width" : "height"), crossSize, scale, column, 16f);
                var item = new FlexItem { Node = child, Style = childStyle, MainSize = requestedMain, CrossSize = requestedCross, Grow = ReadCssFloat(GetStyleValue(childStyle, "flex-grow"), 0f) };
                if (wrap && currentLine.Count > 0 && usedMain + gap + item.MainSize > mainSize)
                {
                    lines.Add(currentLine);
                    currentLine = new List<FlexItem>();
                    usedMain = 0f;
                }
                if (currentLine.Count > 0) usedMain += gap;
                currentLine.Add(item);
                usedMain += item.MainSize;
            }
            if (currentLine.Count > 0) lines.Add(currentLine);

            float crossCursor = column ? area.Left : area.Top;
            float maxBottom = area.Top;
            foreach (List<FlexItem> line in lines)
            {
                float fixedMain = line.Sum(item => item.MainSize) + gap * Math.Max(0, line.Count - 1);
                float grow = line.Sum(item => item.Grow);
                float remaining = Math.Max(0f, mainSize - fixedMain);
                foreach (FlexItem item in line)
                {
                    if (grow > 0f) item.MainSize += remaining * item.Grow / grow;
                }
                float distributedGap = gap;
                float mainCursor = column ? area.Top : area.Left;
                if (justify == "center") mainCursor += Math.Max(0f, mainSize - (line.Sum(item => item.MainSize) + gap * Math.Max(0, line.Count - 1))) / 2f;
                else if (justify == "flex-end" || justify == "end") mainCursor += Math.Max(0f, mainSize - (line.Sum(item => item.MainSize) + gap * Math.Max(0, line.Count - 1)));
                else if (justify == "space-between" && line.Count > 1) distributedGap = (mainSize - line.Sum(item => item.MainSize)) / (line.Count - 1);
                else if (justify == "space-around" && line.Count > 0)
                {
                    distributedGap = (mainSize - line.Sum(item => item.MainSize)) / line.Count;
                    mainCursor += distributedGap / 2f;
                }

                float lineCross = line.Max(item => item.CrossSize > 0f ? item.CrossSize : crossSize);
                foreach (FlexItem item in line)
                {
                    float itemCross = item.CrossSize > 0f ? item.CrossSize : (alignment == "stretch" ? lineCross : Math.Min(lineCross, ScaleY(120f, scale)));
                    float crossOffset = alignment == "center" ? (lineCross - itemCross) / 2f : alignment == "flex-end" || alignment == "end" ? lineCross - itemCross : 0f;
                    var itemArea = column
                        ? new LayoutRect(crossCursor + crossOffset, mainCursor, itemCross, item.MainSize)
                        : new LayoutRect(mainCursor, crossCursor + crossOffset, item.MainSize, itemCross);
                    float rendered = RenderGenericNode(slide, item.Node, itemArea, resolver, scale, fontName, useSourceStyling);
                    if (column) mainCursor += Math.Max(item.MainSize, rendered) + distributedGap;
                    else mainCursor += item.MainSize + distributedGap;
                    maxBottom = Math.Max(maxBottom, itemArea.Top + Math.Max(itemArea.Height, rendered));
                }
                crossCursor += lineCross + gap;
            }
            return Math.Max(0f, maxBottom - area.Top);
        }

        private static int GetCssOrder(Dictionary<string, string> style)
        {
            int value;
            return int.TryParse(GetStyleValue(style, "order"), out value) ? value : 0;
        }

        private static float ReadCssFloat(string value, float fallback)
        {
            float parsed;
            return float.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out parsed) ? parsed : fallback;
        }

        private static string GetPositionValue(Dictionary<string, string> style, string property, int shorthandIndex)
        {
            string direct = GetStyleValue(style, property);
            if (!string.IsNullOrWhiteSpace(direct)) return direct;
            string inset = GetStyleValue(style, "inset");
            if (string.IsNullOrWhiteSpace(inset)) return null;
            string[] values = Regex.Split(inset.Trim(), "\\s+");
            if (values.Length == 1) return values[0];
            if (values.Length == 2) return shorthandIndex == 0 || shorthandIndex == 2 ? values[0] : values[1];
            if (values.Length == 3) return shorthandIndex == 0 ? values[0] : shorthandIndex == 2 ? values[2] : values[1];
            return values[Math.Min(shorthandIndex, values.Length - 1)];
        }

        private static bool IsAutoHorizontalMargin(Dictionary<string, string> style)
        {
            string shorthand = GetStyleValue(style, "margin") ?? string.Empty;
            string left = GetStyleValue(style, "margin-left");
            string right = GetStyleValue(style, "margin-right");
            if (string.Equals(left, "auto", StringComparison.OrdinalIgnoreCase) && string.Equals(right, "auto", StringComparison.OrdinalIgnoreCase)) return true;
            string[] values = Regex.Split(shorthand.Trim(), "\\s+");
            return values.Length >= 2 && string.Equals(values[1], "auto", StringComparison.OrdinalIgnoreCase) &&
                (values.Length < 4 || string.Equals(values[3], "auto", StringComparison.OrdinalIgnoreCase));
        }

        private static void ApplyGenericShapeLayout(Slide slide, int firstShapeIndex, Dictionary<string, string> style, float shiftX, float shiftY, float width, float height, SlideScale scale)
        {
            string transform = GetStyleValue(style, "transform") ?? string.Empty;
            float translateX = 0f;
            float translateY = 0f;
            Match translate = Regex.Match(transform, "translate(?:3d)?\\(\\s*([^,]+),\\s*([^,)]+)", RegexOptions.IgnoreCase);
            if (translate.Success)
            {
                translateX += ReadGenericLength(translate.Groups[1].Value, width, scale, true, 16f);
                translateY += ReadGenericLength(translate.Groups[2].Value, height, scale, false, 16f);
            }
            Match translateXMatch = Regex.Match(transform, "translateX\\(\\s*([^)]+)\\)", RegexOptions.IgnoreCase);
            Match translateYMatch = Regex.Match(transform, "translateY\\(\\s*([^)]+)\\)", RegexOptions.IgnoreCase);
            if (translateXMatch.Success) translateX += ReadGenericLength(translateXMatch.Groups[1].Value, width, scale, true, 16f);
            if (translateYMatch.Success) translateY += ReadGenericLength(translateYMatch.Groups[1].Value, height, scale, false, 16f);

            float rotation = 0f;
            Match rotate = Regex.Match(transform, "rotate\\(\\s*(-?[0-9.]+)(deg|rad|turn)?\\s*\\)", RegexOptions.IgnoreCase);
            if (rotate.Success)
            {
                float parsed;
                if (float.TryParse(rotate.Groups[1].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out parsed))
                {
                    string unit = rotate.Groups[2].Value.ToLowerInvariant();
                    rotation = unit == "rad" ? parsed * 180f / (float)Math.PI : unit == "turn" ? parsed * 360f : parsed;
                }
            }

            int zIndex;
            int.TryParse(GetStyleValue(style, "z-index"), out zIndex);
            var shapes = new List<Shape>();
            for (int index = firstShapeIndex; index <= slide.Shapes.Count; index++) shapes.Add(slide.Shapes[index]);
            foreach (Shape shape in shapes)
            {
                shape.Left += shiftX + translateX;
                shape.Top += shiftY + translateY;
                if (Math.Abs(rotation) > 0.01f) shape.Rotation += rotation;
                if (zIndex > 0) shape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                else if (zIndex < 0) shape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
            }
        }

        private static float[] ReadGenericBox(Dictionary<string, string> style, string property, SlideScale scale, float defaultHorizontal, float defaultVertical)
        {
            string shorthand = GetStyleValue(style, property);
            string[] values = string.IsNullOrWhiteSpace(shorthand) ? new string[0] : Regex.Split(shorthand.Trim(), "\\s+");
            string top = values.Length > 0 ? values[0] : null;
            string right = values.Length > 1 ? values[1] : top;
            string bottom = values.Length > 2 ? values[2] : top;
            string left = values.Length > 3 ? values[3] : right;
            top = GetStyleValue(style, property + "-top") ?? top;
            right = GetStyleValue(style, property + "-right") ?? right;
            bottom = GetStyleValue(style, property + "-bottom") ?? bottom;
            left = GetStyleValue(style, property + "-left") ?? left;
            return new[]
            {
                string.IsNullOrWhiteSpace(top) ? defaultVertical : ReadGenericLength(top, HtmlSlideHeight, scale, false, 16f),
                string.IsNullOrWhiteSpace(right) ? defaultHorizontal : ReadGenericLength(right, HtmlSlideWidth, scale, true, 16f),
                string.IsNullOrWhiteSpace(bottom) ? defaultVertical : ReadGenericLength(bottom, HtmlSlideHeight, scale, false, 16f),
                string.IsNullOrWhiteSpace(left) ? defaultHorizontal : ReadGenericLength(left, HtmlSlideWidth, scale, true, 16f)
            };
        }

        private static float ReadGenericLength(string value, float reference, SlideScale scale, bool horizontal, float fontSize)
        {
            if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)) return 0f;
            string cleaned = value.Trim();
            Match calc = Regex.Match(cleaned, "^calc\\((?<expression>.*)\\)$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (calc.Success)
            {
                float total = 0f;
                bool found = false;
                foreach (Match term in Regex.Matches(calc.Groups["expression"].Value, "(?<sign>[+-]?)\\s*(?<value>[0-9.]+(?:%|vw|vh|rem|em|px|pt)?)", RegexOptions.IgnoreCase))
                {
                    float amount = ReadGenericLength(term.Groups["value"].Value, reference, scale, horizontal, fontSize);
                    total += term.Groups["sign"].Value == "-" ? -amount : amount;
                    found = true;
                }
                if (found) return total;
            }
            Match clamp = Regex.Match(cleaned, "clamp\\([^,]+,([^,]+),[^)]+\\)", RegexOptions.IgnoreCase);
            if (clamp.Success) cleaned = clamp.Groups[1].Value.Trim();
            float parsed;
            if (cleaned.EndsWith("%", StringComparison.Ordinal) && float.TryParse(cleaned.TrimEnd('%'), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out parsed)) return reference * parsed / 100f;
            if (cleaned.EndsWith("vw", StringComparison.OrdinalIgnoreCase) && float.TryParse(cleaned.Substring(0, cleaned.Length - 2), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out parsed)) return ScaleX(HtmlSlideWidth * parsed / 100f, scale);
            if (cleaned.EndsWith("vh", StringComparison.OrdinalIgnoreCase) && float.TryParse(cleaned.Substring(0, cleaned.Length - 2), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out parsed)) return ScaleY(HtmlSlideHeight * parsed / 100f, scale);
            if (cleaned.EndsWith("rem", StringComparison.OrdinalIgnoreCase) && float.TryParse(cleaned.Substring(0, cleaned.Length - 3), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out parsed)) return parsed * fontSize * scale.Average;
            if (cleaned.EndsWith("em", StringComparison.OrdinalIgnoreCase) && float.TryParse(cleaned.Substring(0, cleaned.Length - 2), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out parsed)) return parsed * fontSize * scale.Average;
            cleaned = Regex.Replace(cleaned, "(px|pt)$", string.Empty, RegexOptions.IgnoreCase).Trim();
            return float.TryParse(cleaned, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out parsed) ? parsed * (horizontal ? scale.X : scale.Y) : 0f;
        }

        private static float ReadGenericFontSize(string value, float fallback)
        {
            if (string.IsNullOrWhiteSpace(value)) return CssPixelsToPoints(fallback);
            Match clamp = Regex.Match(value, "clamp\\([^,]+,([^,]+),[^)]+\\)", RegexOptions.IgnoreCase);
            string cleaned = clamp.Success ? clamp.Groups[1].Value.Trim() : value.Trim();
            float parsed;
            if (cleaned.EndsWith("vw", StringComparison.OrdinalIgnoreCase) && float.TryParse(cleaned.Substring(0, cleaned.Length - 2), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out parsed)) return CssPixelsToPoints(HtmlSlideWidth * parsed / 100f);
            bool isPointValue = cleaned.EndsWith("pt", StringComparison.OrdinalIgnoreCase);
            cleaned = Regex.Replace(cleaned, "(px|pt|rem|em)$", string.Empty, RegexOptions.IgnoreCase).Trim();
            if (!float.TryParse(cleaned, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out parsed)) return CssPixelsToPoints(fallback);
            return isPointValue ? parsed : CssPixelsToPoints(parsed);
        }

        private static string GetStyleValue(Dictionary<string, string> style, string property)
        {
            string value;
            return style != null && style.TryGetValue(property, out value) ? value : null;
        }

        private static bool IsCssBold(string value)
        {
            int weight;
            return string.Equals(value, "bold", StringComparison.OrdinalIgnoreCase) || string.Equals(value, "bolder", StringComparison.OrdinalIgnoreCase) || (int.TryParse(value, out weight) && weight >= 600);
        }

        private static PpParagraphAlignment GetParagraphAlignment(string value)
        {
            if (string.Equals(value, "center", StringComparison.OrdinalIgnoreCase)) return PpParagraphAlignment.ppAlignCenter;
            if (string.Equals(value, "right", StringComparison.OrdinalIgnoreCase) || string.Equals(value, "end", StringComparison.OrdinalIgnoreCase)) return PpParagraphAlignment.ppAlignRight;
            if (string.Equals(value, "justify", StringComparison.OrdinalIgnoreCase)) return PpParagraphAlignment.ppAlignJustify;
            return PpParagraphAlignment.ppAlignLeft;
        }

        private string GetGenericFontName(string cssFontFamily)
        {
            if (string.IsNullOrWhiteSpace(cssFontFamily)) return DefaultFont;
            string[] candidates = cssFontFamily.Split(',').Select(value => value.Trim().Trim('\'', '"')).Where(value => !string.IsNullOrWhiteSpace(value)).ToArray();
            foreach (string candidate in candidates)
            {
                if (IsGenericFontFamily(candidate)) return GetPowerPointGenericFont(candidate);
                try
                {
                    if (FontFamily.Families.Any(family => string.Equals(family.Name, candidate, StringComparison.OrdinalIgnoreCase))) return candidate;
                }
                catch
                {
                    return candidate;
                }
            }

            string missing = candidates.FirstOrDefault() ?? DefaultFont;
            if (_reportedMissingFonts.Add(missing)) AddWarning("Font '" + missing + "' is not installed; using " + DefaultFont + ".", null);
            return DefaultFont;
        }

        private static bool IsGenericFontFamily(string value)
        {
            return string.Equals(value, "serif", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "sans-serif", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "monospace", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "cursive", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "fantasy", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "system-ui", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetPowerPointGenericFont(string value)
        {
            if (string.Equals(value, "serif", StringComparison.OrdinalIgnoreCase)) return "Georgia";
            if (string.Equals(value, "monospace", StringComparison.OrdinalIgnoreCase)) return "Consolas";
            if (string.Equals(value, "cursive", StringComparison.OrdinalIgnoreCase)) return "Segoe Script";
            return DefaultFont;
        }

        private static string CssColorToHex(string cssColor, string fallback)
        {
            if (string.IsNullOrWhiteSpace(cssColor)) return fallback;
            Match hex = Regex.Match(cssColor, "#[0-9a-fA-F]{3,8}");
            if (hex.Success) return hex.Value.Length > 7 ? hex.Value.Substring(0, 7) : hex.Value;
            Match rgb = Regex.Match(cssColor, "rgba?\\(\\s*(\\d+)\\s*,\\s*(\\d+)\\s*,\\s*(\\d+)", RegexOptions.IgnoreCase);
            if (rgb.Success) return string.Format("#{0:X2}{1:X2}{2:X2}", Math.Min(255, int.Parse(rgb.Groups[1].Value)), Math.Min(255, int.Parse(rgb.Groups[2].Value)), Math.Min(255, int.Parse(rgb.Groups[3].Value)));
            string named = cssColor.Split(new[] { ' ', ';' }, StringSplitOptions.RemoveEmptyEntries).LastOrDefault();
            try { return "#" + ColorTranslator.FromHtml(named).ToArgb().ToString("X8").Substring(2); } catch { return fallback; }
        }

    }
}
