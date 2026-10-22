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
    /// <summary>
    /// Imports a constrained, AI-generated HTML presentation format and rebuilds it
    /// as editable PowerPoint slides. This MVP intentionally targets section-based
    /// HTML decks instead of arbitrary web pages.
    /// </summary>
    public partial class HtmlPresentationImportService
    {
        private const float HtmlSlideWidth = 900f;
        private const float HtmlSlideHeight = 560f;

        private const string DefaultFont = "Segoe UI";
        private const string LightTextHex = "#2c3e50";
        private const string DarkTextHex = "#ffffff";
        private const string AccentHex = "#ff9f1c";
        private const string HeaderHex = "#1a3a5f";
        private const string MutedHex = "#718096";
        private const string FooterHex = "#a0aec0";
        private const string DarkFooterHex = "#94a3b8";
        private const string LightBackgroundHex = "#ffffff";
        private const string DarkBackgroundStartHex = "#0f2027";
        private const string DarkBackgroundHex = "#203a43";
        private const string DarkBackgroundAltHex = "#2c5364";
        private const string TitleGradientEndHex = "#ffe17d";
        private const string CardBackgroundHex = "#f8fafc";
        private const string CalloutBackgroundHex = "#f3f7fb";
        private const string DangerTextHex = "#c53030";
        private const string BodySecondaryHex = "#4a5568";
        private const string DarkBodySecondaryHex = "#e2e8f0";
        private const string SourceBodyFont = "Segoe UI";
        private const float SlideCardCornerAdjustment = 0.006f;
        private readonly List<string> _temporaryImportFiles = new List<string>();
        private readonly HashSet<string> _reportedMissingFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _reportedBlockedRemoteResources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly bool _allowRemoteResources;
        private readonly HtmlRenderRouter _renderRouter = new HtmlRenderRouter();
        private int _mediaRequestCount;
        private int _combinedMediaBytes;
        private string _htmlBaseDirectory;
        private PresentationThemePalette _presentationTheme;

        public HtmlImportReport LastReport { get; private set; }

        public HtmlPresentationImportService(bool allowRemoteResources = false)
        {
            _allowRemoteResources = allowRemoteResources;
        }

        public HtmlImportPreview GetImportPreview(string htmlFilePath)
        {
            if (string.IsNullOrWhiteSpace(htmlFilePath)) throw new ArgumentException("HTML file path is required.", "htmlFilePath");
            if (!File.Exists(htmlFilePath)) throw new FileNotFoundException("HTML file was not found.", htmlFilePath);

            _htmlBaseDirectory = Path.GetDirectoryName(Path.GetFullPath(htmlFilePath));
            _reportedMissingFonts.Clear();
            ResetMediaBudget();
            LastReport = new HtmlImportReport();
            try
            {
                List<HtmlSlideModel> slides = LoadSlides(htmlFilePath);
                var preview = new HtmlImportPreview();
                preview.Report = LastReport;

                foreach (var slide in slides)
                {
                    preview.Slides.Add(new HtmlImportSlidePreview
                    {
                        Number = slide.Number,
                        Title = !string.IsNullOrWhiteSpace(slide.Title) ? slide.Title : slide.FooterRight
                    });
                }

                return preview;
            }
            finally
            {
                _htmlBaseDirectory = null;
            }
        }

        public int ImportFromFile(PowerPointManager manager, string htmlFilePath, HtmlImportOptions options)
        {
            if (manager == null) throw new ArgumentNullException("manager");
            if (string.IsNullOrWhiteSpace(htmlFilePath)) throw new ArgumentException("HTML file path is required.", "htmlFilePath");
            if (!File.Exists(htmlFilePath)) throw new FileNotFoundException("HTML file was not found.", htmlFilePath);
            if (options == null) options = new HtmlImportOptions();

            var presentation = manager.GetApplication().ActivePresentation;
            if (presentation == null)
            {
                throw new InvalidOperationException("No active presentation is available.");
            }

            _htmlBaseDirectory = Path.GetDirectoryName(Path.GetFullPath(htmlFilePath));
            _reportedMissingFonts.Clear();
            ResetMediaBudget();
            LastReport = new HtmlImportReport();
            _presentationTheme = options.UsePresentationTheme ? PresentationThemePalette.TryCreate(presentation) : null;
            if (options.UsePresentationTheme && _presentationTheme == null)
            {
                AddWarning("The presentation theme could not be read; the HTML import uses default colors.", null);
            }
            try
            {
                List<HtmlSlideModel> slides = LoadSlides(htmlFilePath);
                if (slides.Count == 0)
                {
                    return 0;
                }

                int imported = 0;
                int originalSlideCount = presentation.Slides.Count;
                try
                {
                    foreach (var slideModel in slides)
                    {
                        if (options.SelectedSlideNumbers.Count > 0 && !options.SelectedSlideNumbers.Contains(slideModel.Number))
                        {
                            continue;
                        }

                        RenderSlide(presentation, slideModel, options);
                        imported++;
                    }
                }
                catch
                {
                    for (int index = presentation.Slides.Count; index > originalSlideCount; index--)
                    {
                        try { presentation.Slides[index].Delete(); }
                        catch (Exception ex) { ReportOptionalImportFailure(ex, "HtmlPresentationImportService.Import.RollbackSlide"); }
                    }
                    throw;
                }

                return imported;
            }
            finally
            {
                CleanupTemporaryImportFiles();
                _htmlBaseDirectory = null;
                _presentationTheme = null;
            }
        }

        private void RenderSlide(Presentation presentation, HtmlSlideModel model, HtmlImportOptions options)
        {
            Slide slide = presentation.Slides.Add(presentation.Slides.Count + 1, PpSlideLayout.ppLayoutBlank);
            try
            {
                if (!string.IsNullOrWhiteSpace(model.Id))
                {
                    try { slide.Name = model.Id; }
                    catch (Exception ex) { ReportOptionalImportFailure(ex, "HtmlPresentationImportService.RenderSlide.Name"); }
                }
                float slideWidth = presentation.PageSetup.SlideWidth;
                float slideHeight = presentation.PageSetup.SlideHeight;
                bool useSourceBackground = options != null && options.PreserveSourceBackground;
                bool useSourceStyling = (options == null || options.PreserveSourceStyling) && _presentationTheme == null;
                bool useDarkTheme = model.IsDark && useSourceBackground;
                string fontName = GetFontName(options);
                LayoutRect canvasRect = useSourceBackground
                    ? GetHtmlCanvasRect(slideWidth, slideHeight)
                    : new LayoutRect(0f, 0f, slideWidth, slideHeight);
                var scale = new SlideScale(canvasRect.Width / HtmlSlideWidth, canvasRect.Height / HtmlSlideHeight, canvasRect.Left, canvasRect.Top);

                if (model.IsGptDeck)
                {
                    RenderGptDeckSlide(slide, model, scale, useSourceBackground, useSourceStyling, fontName);
                    return;
                }

                if (model.IsGenericDom)
                {
                    RenderGenericDomSlide(slide, model, scale, useSourceBackground, useSourceStyling, fontName);
                    return;
                }

                if (useSourceBackground)
                {
                    RenderBackground(slide, model.IsDark, slideWidth, slideHeight, canvasRect);
                }

                if (!string.IsNullOrWhiteSpace(model.TitleSlideHtml))
                {
                    RenderTitleSlide(slide, model, scale, useDarkTheme, useSourceStyling, fontName);
                    if (ContainsMedia(model.TitleSlideHtml))
                    {
                        RenderEmbeddedMedia(slide, model.TitleSlideHtml, new LayoutRect(MapX(300f, scale), MapY(72f, scale), ScaleX(300f, scale), ScaleY(72f, scale)), scale);
                    }
                    return;
                }

                RenderHeader(slide, model, scale, useDarkTheme, useSourceStyling, fontName);
                RenderFooter(slide, model, scale, useDarkTheme, useSourceStyling, fontName);

                string contentHtml = model.ContentHtml ?? string.Empty;
                if (contentHtml.IndexOf("process-flow", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    RenderProcessFlow(slide, contentHtml, scale, useSourceStyling, fontName);
                    return;
                }

                string columnsHtml = ExtractBlockByClass(contentHtml, "div", "content-columns");
                if (!string.IsNullOrWhiteSpace(columnsHtml))
                {
                    RenderContentColumns(slide, columnsHtml, scale, useDarkTheme, useSourceStyling, fontName);
                    return;
                }

                RenderStandaloneRichContent(slide, contentHtml, scale, useDarkTheme, useSourceStyling, fontName);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "HtmlPresentationImportService.RenderSlide." + model.Number);
                try { slide.Delete(); }
                catch (Exception cleanupEx) { ReportOptionalImportFailure(cleanupEx, "HtmlPresentationImportService.RenderSlide.Rollback"); }
                throw;
            }
        }

        private void RenderBackground(Slide slide, bool isDark, float slideWidth, float slideHeight, LayoutRect canvasRect)
        {
            if (isDark)
            {
                RenderDarkGradientBackground(slide, slideWidth, slideHeight);
            }
            else
            {
                var background = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, slideWidth, slideHeight);
                background.Name = "oPE_HtmlImport_Background_" + Guid.NewGuid().ToString("N");
                background.Line.Visible = Office.MsoTriState.msoFalse;
                background.Fill.ForeColor.RGB = ToRgb(LightBackgroundHex);
                background.Fill.Solid();
                background.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
            }
        }

        private static LayoutRect GetHtmlCanvasRect(float slideWidth, float slideHeight)
        {
            return new LayoutRect(0f, 0f, slideWidth, slideHeight);
        }

        private void RenderTitleSlide(Slide slide, HtmlSlideModel model, SlideScale scale, bool useDarkTheme, bool useSourceStyling, string fontName)
        {
            string html = model.TitleSlideHtml ?? string.Empty;
            string titleHtml = ExtractFirstTagInnerHtml(html, "h1");
            string subtitleHtml = ExtractFirstTagInnerHtml(html, "p");
            string badgeHtml = ExtractBlockByClass(html, "div", "accent-badge");
            string title = CleanInlineText(titleHtml);
            string subtitle = CleanInlineText(subtitleHtml);
            string badge = CleanInlineText(badgeHtml);
            float slideCanvasWidth = ScaleX(HtmlSlideWidth, scale);

            var titleShape = AddTextBox(slide, MapX(90f, scale), MapY(155f, scale), ScaleX(720f, scale), ScaleY(92f, scale), title, 46f * scale.Average, GetTitleColor(useSourceStyling), true, PpParagraphAlignment.ppAlignCenter, Office.MsoVerticalAnchor.msoAnchorMiddle, fontName);
            var subtitleShape = AddTextBox(slide, MapX(120f, scale), MapY(245f, scale), ScaleX(660f, scale), ScaleY(62f, scale), subtitle, 20f * scale.Average, GetSubtitleColor(useDarkTheme, useSourceStyling), false, PpParagraphAlignment.ppAlignCenter, Office.MsoVerticalAnchor.msoAnchorMiddle, fontName);
            ApplyHtmlHyperlinks(titleShape, titleHtml);
            ApplyHtmlHyperlinks(subtitleShape, subtitleHtml);

            if (useDarkTheme && useSourceStyling)
            {
                ApplyTitleTextGradient(titleShape);
            }

            if (!string.IsNullOrWhiteSpace(badge))
            {
                float badgeWidth = ScaleX(300f, scale);
                float badgeHeight = ScaleY(42f, scale);
                float badgeLeft = scale.OffsetX + ((slideCanvasWidth - badgeWidth) / 2f);
                float badgeTop = MapY(350f, scale);

                var badgeShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, badgeLeft, badgeTop, badgeWidth, badgeHeight);
                badgeShape.Line.ForeColor.RGB = ToRgb("#d89a37");
                badgeShape.Line.Transparency = 0.42f;
                badgeShape.Fill.ForeColor.RGB = ToRgb(GetBadgeBackgroundColor(useDarkTheme, useSourceStyling));
                badgeShape.Fill.Transparency = useDarkTheme && useSourceStyling ? 0.18f : 0f;
                badgeShape.Adjustments[1] = 0.2f;

                SetShapeText(badgeShape, badge, 15f * scale.Average, GetAccentColor(useSourceStyling), true, PpParagraphAlignment.ppAlignCenter, Office.MsoVerticalAnchor.msoAnchorMiddle, fontName);
                ApplyHtmlHyperlinks(badgeShape, badgeHtml);
            }

            RenderFooter(slide, model, scale, useDarkTheme, useSourceStyling, fontName);
        }

        private void RenderHeader(Slide slide, HtmlSlideModel model, SlideScale scale, bool useDarkTheme, bool useSourceStyling, string fontName)
        {
            string title = model.Title;
            if (string.IsNullOrWhiteSpace(title))
            {
                title = CleanInlineText(ExtractFirstTagInnerHtml(model.HeaderHtml, "h2"));
            }

            string titleHtml = ExtractFirstTagInnerHtml(model.HeaderHtml, "h2");
            var titleShape = AddTextBox(slide, MapX(50f, scale), MapY(36f, scale), ScaleX(760f, scale), ScaleY(34f, scale), title, 28f * scale.Average, GetHeaderColor(useDarkTheme, useSourceStyling), true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorMiddle, fontName);
            ApplyHtmlHyperlinks(titleShape, titleHtml);

            var accent = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, MapX(50f, scale), MapY(74f, scale), ScaleX(60f, scale), Math.Max(2f, ScaleY(4f, scale)));
            accent.Fill.ForeColor.RGB = ToRgb(GetAccentColor(useSourceStyling));
            accent.Line.Visible = Office.MsoTriState.msoFalse;
        }

        private void RenderFooter(Slide slide, HtmlSlideModel model, SlideScale scale, bool isDark, bool useSourceStyling, string fontName)
        {
            float top = MapY(500f, scale);
            float left = MapX(55f, scale);
            float width = ScaleX(790f, scale);

            var divider = slide.Shapes.AddLine(left, top, left + width, top);
            divider.Line.ForeColor.RGB = ToRgb(GetDividerColor(isDark, useSourceStyling));
            divider.Line.Transparency = isDark ? 0.35f : 0.15f;
            divider.Line.Weight = Math.Max(0.75f, scale.Average);

            AddTextBox(slide, left, top + ScaleY(8f, scale), ScaleX(520f, scale), ScaleY(18f, scale), model.FooterLeft, 10.5f * scale.Average, GetFooterColor(isDark, useSourceStyling), false, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorMiddle, fontName);
            AddTextBox(slide, left + ScaleX(590f, scale), top + ScaleY(8f, scale), ScaleX(200f, scale), ScaleY(18f, scale), model.FooterRight, 10.5f * scale.Average, GetFooterColor(isDark, useSourceStyling), false, PpParagraphAlignment.ppAlignRight, Office.MsoVerticalAnchor.msoAnchorMiddle, fontName);
        }

        private void RenderContentColumns(Slide slide, string columnsHtml, SlideScale scale, bool isDark, bool useSourceStyling, string fontName)
        {
            List<HtmlBlock> cols = FindBlocksByClass(columnsHtml, "div", "col");
            if (cols.Count == 0)
            {
                return;
            }

            float gap = ScaleX(40f, scale);
            float totalWidth = ScaleX(800f, scale);
            float left = MapX(50f, scale);
            float top = MapY(110f, scale);
            float height = ScaleY(360f, scale);
            var plannedColumns = new HtmlLayoutEngine().Columns(
                new HtmlLayoutRect(left, top, totalWidth, height), cols.Count, gap);

            for (int i = 0; i < cols.Count; i++)
            {
                HtmlLayoutRect planned = plannedColumns[i];
                var rect = new LayoutRect(planned.Left, planned.Top, planned.Width, planned.Height);
                int shapeCountBeforeColumn = slide.Shapes.Count;
                RenderColumn(slide, cols[i].InnerHtml, rect, scale, isDark, useSourceStyling, fontName);
                CenterColumnContent(slide, shapeCountBeforeColumn, rect);
            }
        }

        private static void CenterColumnContent(Slide slide, int shapeCountBeforeColumn, LayoutRect rect)
        {
            float contentTop = float.MaxValue;
            float contentBottom = float.MinValue;
            int shapeCountAfterColumn = slide.Shapes.Count;

            for (int index = shapeCountBeforeColumn + 1; index <= shapeCountAfterColumn; index++)
            {
                Shape shape = slide.Shapes[index];
                contentTop = Math.Min(contentTop, shape.Top);
                contentBottom = Math.Max(contentBottom, shape.Top + shape.Height);
            }

            if (contentTop == float.MaxValue || contentBottom <= contentTop)
            {
                return;
            }

            // HTML uses align-items: center on .content-columns. Each column must therefore
            // be vertically centered according to its own rendered height, not its neighbor's.
            float targetTop = rect.Top + (rect.Height - (contentBottom - contentTop)) / 2f;
            float offset = targetTop - contentTop;

            for (int index = shapeCountBeforeColumn + 1; index <= shapeCountAfterColumn; index++)
            {
                slide.Shapes[index].IncrementTop(offset);
            }
        }

        private void RenderColumn(Slide slide, string columnHtml, LayoutRect rect, SlideScale scale, bool isDark, bool useSourceStyling, string fontName)
        {
            float currentY = rect.Top;
            string profileGridHtml = ExtractBlockByClass(columnHtml, "div", "profile-grid");
            if (!string.IsNullOrWhiteSpace(profileGridHtml))
            {
                float usedHeight = RenderProfileGrid(slide, profileGridHtml, rect.Left, currentY, rect.Width, scale, isDark, useSourceStyling, fontName);
                currentY += usedHeight + ScaleY(14f, scale);
            }

            string heading = CleanInlineText(ExtractFirstTagInnerHtml(columnHtml, "h3"));
            if (!string.IsNullOrWhiteSpace(heading))
            {
                AddTextBox(slide, rect.Left, currentY, rect.Width, ScaleY(24f, scale), heading, 18f * scale.Average, GetSmallHeadingColor(isDark, useSourceStyling), true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorMiddle, fontName);
                currentY += ScaleY(30f, scale);
            }

            List<HtmlBlock> directBlocks = FindDirectChildBlocks(columnHtml, "div");
            foreach (var block in directBlocks)
            {
                if (ClassContains(block, "profile-grid") || !ClassContains(block, "profile-item"))
                {
                    continue;
                }

                DrawProfileItem(slide, block, rect.Left, currentY, rect.Width, ScaleY(82f, scale), scale, isDark, useSourceStyling, fontName);
                currentY += ScaleY(96f, scale);
            }

            List<HtmlBlock> paragraphBlocks = FindDirectChildTagBlocks(columnHtml, "p");
            foreach (var paragraph in paragraphBlocks)
            {
                currentY += RenderParagraph(slide, paragraph, rect.Left, currentY, rect.Width, scale, isDark, useSourceStyling, fontName) + ScaleY(10f, scale);
            }

            foreach (var block in directBlocks)
            {
                if (ClassContains(block, "callout-box"))
                {
                    currentY += RenderCalloutBox(slide, block.InnerHtml, rect.Left, currentY, rect.Width, scale, isDark, useSourceStyling, fontName) + ScaleY(12f, scale);
                }
            }

            string unorderedList = ExtractFirstTagBlock(columnHtml, "ul");
            if (!string.IsNullOrWhiteSpace(unorderedList))
            {
                currentY += RenderBulletList(slide, unorderedList, rect.Left, currentY, rect.Width, scale, isDark, useSourceStyling, fontName);
            }

            if (ContainsTable(columnHtml))
            {
                currentY += RenderHtmlTables(slide, columnHtml, rect.Left, currentY, rect.Width, scale, isDark, useSourceStyling, fontName);
            }
            if (ContainsMedia(columnHtml))
            {
                RenderEmbeddedMedia(slide, columnHtml, new LayoutRect(rect.Left, currentY, rect.Width, Math.Max(0f, rect.Top + rect.Height - currentY)), scale);
            }
        }

        private float RenderProfileGrid(Slide slide, string gridHtml, float left, float top, float width, SlideScale scale, bool isDark, bool useSourceStyling, string fontName)
        {
            List<HtmlBlock> items = FindBlocksByClass(gridHtml, "div", "profile-item");
            if (items.Count == 0)
            {
                return 0f;
            }

            float gap = ScaleX(20f, scale);
            float cardHeight = ScaleY(112f, scale);
            float fullCardHeight = ScaleY(76f, scale);
            float halfWidth = (width - gap) / 2f;
            float currentY = top;
            int colIndex = 0;
            int rowIndex = 0;

            for (int i = 0; i < items.Count; i++)
            {
                bool spanTwo = items[i].Style.IndexOf("grid-column: span 2", StringComparison.OrdinalIgnoreCase) >= 0;
                float cardLeft;
                float cardTop;
                float cardWidth;
                float currentHeight = spanTwo ? fullCardHeight : cardHeight;

                if (spanTwo)
                {
                    if (colIndex != 0)
                    {
                        rowIndex++;
                        colIndex = 0;
                    }

                    cardLeft = left;
                    cardTop = top + rowIndex * (cardHeight + gap);
                    cardWidth = width;
                    rowIndex++;
                }
                else
                {
                    cardLeft = left + colIndex * (halfWidth + gap);
                    cardTop = top + rowIndex * (cardHeight + gap);
                    cardWidth = halfWidth;
                    colIndex++;

                    if (colIndex >= 2)
                    {
                        colIndex = 0;
                        rowIndex++;
                    }
                }

                DrawProfileItem(slide, items[i], cardLeft, cardTop, cardWidth, currentHeight, scale, isDark, useSourceStyling, fontName);
                currentY = Math.Max(currentY, cardTop + currentHeight);
            }

            return currentY - top;
        }

        private void DrawProfileItem(Slide slide, HtmlBlock item, float left, float top, float width, float height, SlideScale scale, bool isDark, bool useSourceStyling, string fontName)
        {
            string fillHex = ReadStyle(item.Style, "background");
            if (string.IsNullOrWhiteSpace(fillHex) || !useSourceStyling)
            {
                fillHex = CardBackgroundHex;
            }

            string accentHex = ReadStyle(item.Style, "border-left-color");
            if (string.IsNullOrWhiteSpace(accentHex) || !useSourceStyling)
            {
                accentHex = HeaderHex;
            }

            float borderWidth = ScaleX(4f, scale);
            var accent = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left, top, width, height);
            accent.Line.Visible = Office.MsoTriState.msoFalse;
            accent.Fill.ForeColor.RGB = ToRgb(accentHex);
            accent.Adjustments[1] = 0.08f;

            var body = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left + borderWidth, top, width - borderWidth, height);
            body.Line.Visible = Office.MsoTriState.msoFalse;
            body.Fill.ForeColor.RGB = ToRgb(fillHex);
            body.Adjustments[1] = 0.08f;

            HtmlBlock labelBlock = FindBlocksByClass(item.InnerHtml, "div", "profile-label").FirstOrDefault();
            HtmlBlock valueBlock = FindBlocksByClass(item.InnerHtml, "div", "profile-value").FirstOrDefault();
            string label = CleanInlineText(labelBlock != null ? labelBlock.InnerHtml : string.Empty).ToUpperInvariant();
            string value = CleanInlineText(valueBlock != null ? valueBlock.InnerHtml : string.Empty);

            string labelColor = useSourceStyling ? ReadStyle(labelBlock != null ? labelBlock.Style : string.Empty, "color") : MutedHex;
            if (string.IsNullOrWhiteSpace(labelColor)) labelColor = useSourceStyling ? "#718096" : MutedHex;

            string valueColor = useSourceStyling ? ReadStyle(valueBlock != null ? valueBlock.Style : string.Empty, "color") : LightTextHex;
            if (string.IsNullOrWhiteSpace(valueColor)) valueColor = useSourceStyling ? "#2d3748" : LightTextHex;

            float textLeft = left + borderWidth + ScaleX(20f, scale);
            float textWidth = width - borderWidth - ScaleX(40f, scale);
            var labelShape = AddTextBox(slide, textLeft, top + ScaleY(15f, scale), textWidth, ScaleY(20f, scale), label, 13f * scale.Average, labelColor, true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
            labelShape.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
            labelShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;

            float valueTop = labelShape.Top + labelShape.Height + ScaleY(4f, scale);
            var valueShape = AddTextBox(slide, textLeft, valueTop, textWidth, Math.Max(ScaleY(22f, scale), top + height - ScaleY(15f, scale) - valueTop), value, 18f * scale.Average, valueColor, true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
            valueShape.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
            valueShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            ApplyHtmlHyperlinks(labelShape, labelBlock != null ? labelBlock.InnerHtml : null);
            ApplyHtmlHyperlinks(valueShape, valueBlock != null ? valueBlock.InnerHtml : null);
        }

        private float RenderCalloutBox(Slide slide, string html, float left, float top, float width, SlideScale scale, bool isDark, bool useSourceStyling, string fontName)
        {
            string fillHex = useSourceStyling ? (isDark ? "#38505f" : CalloutBackgroundHex) : "#f6f8fb";
            string accentHex = useSourceStyling ? (isDark ? AccentHex : HeaderHex) : GetAccentColor(useSourceStyling);
            float height = ScaleY(76f, scale);

            var box = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left, top, width, height);
            box.Line.Visible = Office.MsoTriState.msoFalse;
            box.Fill.ForeColor.RGB = ToRgb(fillHex);
            box.Adjustments[1] = 0.08f;

            var accent = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, ScaleX(4f, scale), height);
            accent.Line.Visible = Office.MsoTriState.msoFalse;
            accent.Fill.ForeColor.RGB = ToRgb(accentHex);

            var textShape = AddTextBox(slide, left + ScaleX(20f, scale), top, width - ScaleX(40f, scale), height, CleanInlineText(html), 16f * scale.Average, GetBodyColor(isDark, useSourceStyling), false, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorMiddle, fontName);
            textShape.TextFrame.TextRange.Font.Italic = Office.MsoTriState.msoTrue;
            ApplyHtmlBoldFormatting(textShape, html);
            ApplyHtmlHyperlinks(textShape, html);

            // PowerPoint does not consistently honor VerticalAnchor for imported text boxes.
            // Fit to the rendered text first, then place that text block in the box's true center.
            textShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            textShape.Top = top + (height - textShape.Height) / 2f;

            return height;
        }

        private float RenderParagraph(Slide slide, HtmlBlock paragraph, float left, float top, float width, SlideScale scale, bool isDark, bool useSourceStyling, string fontName)
        {
            string sizeText = ReadStyle(paragraph.Style, "font-size");
            float fontSize = CssPixelsToPoints(16f);
            if (!string.IsNullOrWhiteSpace(sizeText))
            {
                float parsed;
                if (TryParseCssPixels(sizeText, out parsed))
                {
                    fontSize = CssPixelsToPoints(parsed);
                }
            }

            string paragraphColor = useSourceStyling ? ReadStyle(paragraph.Style, "color") : null;
            if (string.IsNullOrWhiteSpace(paragraphColor)) paragraphColor = GetBodyColor(isDark, useSourceStyling);

            var shape = AddTextBox(slide, left, top, width, ScaleY(72f, scale), CleanInlineText(paragraph.InnerHtml), fontSize * scale.Average, paragraphColor, false, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
            shape.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
            shape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;

            ApplyHtmlBoldFormatting(shape, paragraph.InnerHtml);
            ApplyHtmlHyperlinks(shape, paragraph.InnerHtml);

            return shape.Height;
        }

        private float RenderBulletList(Slide slide, string unorderedListHtml, float left, float top, float width, SlideScale scale, bool isDark, bool useSourceStyling, string fontName)
        {
            List<string> items = ExtractTagContents(unorderedListHtml, "li");
            float currentY = top;

            foreach (string itemHtml in items)
            {
                string text = CleanInlineText(itemHtml);

                AddTextBox(slide, left, currentY, ScaleX(18f, scale), ScaleY(26f, scale), "•", 24f * scale.Average, GetAccentColor(useSourceStyling), true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);

                var textShape = AddTextBox(slide, left + ScaleX(25f, scale), currentY, width - ScaleX(25f, scale), ScaleY(64f, scale), text, 18f * scale.Average, GetTextColor(isDark, useSourceStyling), false, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                textShape.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
                textShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;

                ApplyHtmlBoldFormatting(textShape, itemHtml);
                ApplyHtmlHyperlinks(textShape, itemHtml);

                currentY += textShape.Height + ScaleY(10f, scale);
            }

            return currentY - top;
        }

        private void RenderProcessFlow(Slide slide, string contentHtml, SlideScale scale, bool useSourceStyling, string fontName)
        {
            List<HtmlBlock> steps = FindBlocksByClass(contentHtml, "div", "process-step");
            float left = MapX(60f, scale);
            float top = MapY(120f, scale);
            float gap = ScaleY(18f, scale);
            float currentY = top;

            foreach (var step in steps)
            {
                string number = CleanInlineText(ExtractBlockTextByClass(step.InnerHtml, "div", "step-number"));
                string title = CleanInlineText(ExtractFirstTagInnerHtml(step.InnerHtml, "h4"));
                string text = CleanInlineText(ExtractFirstTagInnerHtml(step.InnerHtml, "p"));

                float circleSize = ScaleY(28f, scale);
                var circle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left, currentY + ScaleY(2f, scale), circleSize, circleSize);
                circle.Line.Visible = Office.MsoTriState.msoFalse;
                circle.Fill.ForeColor.RGB = ToRgb(GetAccentColor(useSourceStyling));
                SetShapeText(circle, number, 11.5f * scale.Average, "#ffffff", true, PpParagraphAlignment.ppAlignCenter, Office.MsoVerticalAnchor.msoAnchorMiddle, fontName);

                string titleHtml = ExtractFirstTagInnerHtml(step.InnerHtml, "h4");
                var titleShape = AddTextBox(slide, left + ScaleX(42f, scale), currentY, ScaleX(700f, scale), ScaleY(20f, scale), title, 14f * scale.Average, GetHeaderColor(false, useSourceStyling), true, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                ApplyHtmlHyperlinks(titleShape, titleHtml);

                var textShape = AddTextBox(slide, left + ScaleX(42f, scale), currentY + ScaleY(18f, scale), ScaleX(700f, scale), ScaleY(64f, scale), text, 12.5f * scale.Average, GetBodyColor(false, useSourceStyling), false, PpParagraphAlignment.ppAlignLeft, Office.MsoVerticalAnchor.msoAnchorTop, fontName);
                textShape.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
                textShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                ApplyHtmlBoldFormatting(textShape, ExtractFirstTagInnerHtml(step.InnerHtml, "p"));
                ApplyHtmlHyperlinks(textShape, ExtractFirstTagInnerHtml(step.InnerHtml, "p"));

                currentY += Math.Max(circle.Height, textShape.Height + ScaleY(18f, scale)) + gap;
            }
        }

        private void RenderStandaloneRichContent(Slide slide, string html, SlideScale scale, bool isDark, bool useSourceStyling, string fontName)
        {
            float left = MapX(50f, scale);
            float top = MapY(110f, scale);
            float width = ScaleX(800f, scale);
            float currentY = top;

            foreach (HtmlBlock paragraph in FindDirectChildTagBlocks(html, "p"))
            {
                currentY += RenderParagraph(slide, paragraph, left, currentY, width, scale, isDark, useSourceStyling, fontName) + ScaleY(10f, scale);
            }

            string listHtml = ExtractFirstTagBlock(html, "ul");
            if (!string.IsNullOrWhiteSpace(listHtml))
            {
                currentY += RenderBulletList(slide, listHtml, left, currentY, width, scale, isDark, useSourceStyling, fontName);
            }

            if (ContainsTable(html))
            {
                currentY += RenderHtmlTables(slide, html, left, currentY, width, scale, isDark, useSourceStyling, fontName);
            }
            if (ContainsMedia(html))
            {
                RenderEmbeddedMedia(slide, html, new LayoutRect(left, currentY, width, Math.Max(0f, MapY(480f, scale) - currentY)), scale);
            }
        }

        private float RenderHtmlTables(Slide slide, string html, float left, float top, float width, SlideScale scale, bool isDark, bool useSourceStyling, string fontName)
        {
            float currentY = top;
            foreach (HtmlBlock tableBlock in FindTagBlocks(html, "table"))
            {
                List<HtmlBlock> rows = FindTagBlocks(tableBlock.InnerHtml, "tr");
                float tableHeight = RenderStructuredTable(slide, rows, left, currentY, width, ScaleY(38f, scale), scale, isDark, fontName, null, null);
                if (tableHeight > 0f) currentY += tableHeight + ScaleY(16f, scale);
            }

            return currentY - top;
        }

        private float RenderStructuredTable(Slide slide, List<HtmlBlock> rows, float left, float top, float width, float defaultRowHeight, SlideScale scale, bool isDark, string fontName, string headerFillOverride, string headerTextOverride)
        {
            TableLayout layout = BuildTableLayout(rows);
            if (layout.ColumnCount == 0 || layout.RowCount == 0) return 0f;
            float[] rowHeights = Enumerable.Range(0, layout.RowCount).Select(index => defaultRowHeight).ToArray();
            foreach (TableCellPlacement placement in layout.Cells)
            {
                float requestedHeight = ReadCssDimension(ReadStyle(placement.Block.Style, "height") ?? ReadAttribute(placement.Block.OpenTag, "height"), defaultRowHeight / scale.Y, scale.Y);
                if (requestedHeight > 0f) rowHeights[placement.Row] = Math.Max(rowHeights[placement.Row], requestedHeight / placement.RowSpan);
            }
            float tableHeight = rowHeights.Sum();
            Shape tableShape = slide.Shapes.AddTable(layout.RowCount, layout.ColumnCount, left, top, width, tableHeight);
            Table table = tableShape.Table;

            for (int row = 0; row < layout.RowCount; row++)
            {
                try { table.Rows[row + 1].Height = rowHeights[row]; }
                catch (Exception ex) { ReportOptionalImportFailure(ex, "HtmlPresentationImportService.ApplyTableRowHeights"); }
            }
            ApplyTableColumnWidths(table, layout, width, scale);

            foreach (TableCellPlacement placement in layout.Cells.Where(value => value.RowSpan > 1 || value.ColumnSpan > 1))
            {
                try
                {
                    dynamic startCell = table.Cell(placement.Row + 1, placement.Column + 1);
                    dynamic endCell = table.Cell(placement.Row + placement.RowSpan, placement.Column + placement.ColumnSpan);
                    startCell.Merge(endCell);
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "HtmlPresentationImportService.MergeTableCell");
                    AddWarning("Merged table cells could not be fully combined.", null);
                }
            }

            foreach (TableCellPlacement placement in layout.Cells)
            {
                Cell cell = table.Cell(placement.Row + 1, placement.Column + 1);
                HtmlBlock cellBlock = placement.Block;
                bool isHeader = string.Equals(cellBlock.TagName, "th", StringComparison.OrdinalIgnoreCase);
                string cellHtml = cellBlock.InnerHtml ?? string.Empty;
                string fillColor = ReadStyle(cellBlock.Style, "background-color") ?? ReadStyle(cellBlock.Style, "background");
                string textColor = ReadStyle(cellBlock.Style, "color");
                string fontSizeText = ReadStyle(cellBlock.Style, "font-size");
                float fontSize = CssPixelsToPoints(14f);
                float parsedFontSize;
                if (TryParseCssPixels(fontSizeText, out parsedFontSize)) fontSize = CssPixelsToPoints(parsedFontSize);
                if (string.IsNullOrWhiteSpace(fillColor)) fillColor = isHeader ? (headerFillOverride ?? (isDark ? "#38505f" : "#eef3f8")) : (isDark ? "#294955" : "#ffffff");
                if (string.IsNullOrWhiteSpace(textColor)) textColor = isHeader && !string.IsNullOrWhiteSpace(headerTextOverride) ? headerTextOverride : (isDark ? DarkTextHex : LightTextHex);
                cell.Shape.Fill.Visible = Office.MsoTriState.msoTrue;
                cell.Shape.Fill.ForeColor.RGB = ToRgb(CssColorToHex(fillColor, "#ffffff"));
                SetTableCellBorders(cell, ReadStyle(cellBlock.Style, "border-color") ?? (isDark ? "#506672" : "#d7dde5"));
                cell.Shape.TextFrame.MarginLeft = ScaleX(8f, scale);
                cell.Shape.TextFrame.MarginRight = ScaleX(8f, scale);
                cell.Shape.TextFrame.MarginTop = ScaleY(5f, scale);
                cell.Shape.TextFrame.MarginBottom = ScaleY(5f, scale);
                SetShapeText(cell.Shape, CleanInlineText(cellHtml), fontSize * scale.Average, textColor, isHeader || IsCssBold(ReadStyle(cellBlock.Style, "font-weight")), GetParagraphAlignment(ReadStyle(cellBlock.Style, "text-align")), Office.MsoVerticalAnchor.msoAnchorMiddle, fontName);
                ApplyHtmlBoldFormatting(cell.Shape, cellHtml);
                ApplyHtmlItalicFormatting(cell.Shape, cellHtml);
                ApplyHtmlHyperlinks(cell.Shape, cellHtml);
            }
            return tableHeight;
        }

        private static TableLayout BuildTableLayout(List<HtmlBlock> rows)
        {
            var layout = new TableLayout { RowCount = rows == null ? 0 : rows.Count };
            var occupied = new HashSet<string>();
            for (int row = 0; row < layout.RowCount; row++)
            {
                int column = 0;
                foreach (HtmlBlock block in FindTableCells(rows[row].InnerHtml))
                {
                    while (occupied.Contains(row + ":" + column)) column++;
                    int columnSpan = Math.Max(1, ReadHtmlSpan(block.OpenTag, "colspan"));
                    int rowSpan = Math.Max(1, ReadHtmlSpan(block.OpenTag, "rowspan"));
                    layout.Cells.Add(new TableCellPlacement { Row = row, Column = column, RowSpan = rowSpan, ColumnSpan = columnSpan, Block = block });
                    for (int y = row; y < row + rowSpan; y++)
                    {
                        for (int x = column; x < column + columnSpan; x++) occupied.Add(y + ":" + x);
                    }
                    layout.ColumnCount = Math.Max(layout.ColumnCount, column + columnSpan);
                    column += columnSpan;
                }
            }
            return layout;
        }

        private static int ReadHtmlSpan(string openTag, string attribute)
        {
            int value;
            return int.TryParse(ReadAttribute(openTag, attribute), out value) ? value : 1;
        }

        private static void ApplyTableColumnWidths(Table table, TableLayout layout, float totalWidth, SlideScale scale)
        {
            var requested = new Dictionary<int, float>();
            foreach (TableCellPlacement placement in layout.Cells.Where(value => value.ColumnSpan == 1))
            {
                float width = ReadCssDimension(ReadStyle(placement.Block.Style, "width") ?? ReadAttribute(placement.Block.OpenTag, "width"), totalWidth / scale.X, scale.X);
                if (width > 0f && !requested.ContainsKey(placement.Column)) requested[placement.Column] = width;
            }
            if (requested.Count == 0) return;
            float requestedTotal = requested.Values.Sum();
            float remaining = Math.Max(1f, totalWidth - requestedTotal);
            int flexible = Math.Max(1, layout.ColumnCount - requested.Count);
            for (int column = 0; column < layout.ColumnCount; column++)
            {
                try { table.Columns[column + 1].Width = requested.ContainsKey(column) ? requested[column] : remaining / flexible; }
                catch (Exception ex) { ReportOptionalImportFailure(ex, "HtmlPresentationImportService.ApplyTableColumnWidths"); }
            }
        }

        private static List<HtmlBlock> FindTableCells(string rowHtml)
        {
            var cells = new List<HtmlBlock>();
            MatchCollection matches = Regex.Matches(rowHtml ?? string.Empty, "<(?<tag>th|td)\\b(?<attrs>[^>]*)>(?<content>.*?)</\\k<tag>\\s*>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            foreach (Match match in matches)
            {
                string openTag = "<" + match.Groups["tag"].Value + match.Groups["attrs"].Value + ">";
                cells.Add(new HtmlBlock
                {
                    TagName = match.Groups["tag"].Value,
                    OpenTag = openTag,
                    InnerHtml = match.Groups["content"].Value,
                    FullHtml = match.Value,
                    Classes = SplitClasses(ReadAttribute(openTag, "class")),
                    Style = ReadAttribute(openTag, "style") ?? string.Empty
                });
            }
            return cells;
        }

        private float RenderEmbeddedMedia(Slide slide, string html, LayoutRect area, SlideScale scale)
        {
            if (string.IsNullOrWhiteSpace(html) || area.Width <= 0f)
            {
                return 0f;
            }

            float currentY = area.Top;
            string mediaHtml = Regex.Replace(html, "<table\\b[^>]*>.*?</table\\s*>", string.Empty, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            MatchCollection imageMatches = Regex.Matches(mediaHtml, "<img\\b[^>]*>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            foreach (Match imageMatch in imageMatches)
            {
                string imageTag = imageMatch.Value;
                string source = ReadAttribute(imageTag, "src");
                string resolvedPath = ResolveMediaSource(source);
                if (string.IsNullOrWhiteSpace(resolvedPath))
                {
                    continue;
                }

                Shape picture = AddMediaPicture(slide, resolvedPath, imageTag, area.Left, currentY, area.Width, area.Height, scale);
                if (picture == null)
                {
                    continue;
                }

                picture.AlternativeText = WebUtility.HtmlDecode(ReadAttribute(imageTag, "alt") ?? string.Empty);
                ApplyShapeHyperlink(picture, FindWrappingAnchorHref(mediaHtml, imageTag));
                currentY = picture.Top + picture.Height + ScaleY(12f, scale);
            }

            foreach (HtmlBlock svgBlock in FindTagBlocks(mediaHtml, "svg"))
            {
                string svgPath = CreateTemporaryFile(".svg", Encoding.UTF8.GetBytes(svgBlock.FullHtml));
                Shape svgShape = AddMediaPicture(slide, svgPath, svgBlock.OpenTag, area.Left, currentY, area.Width, area.Height, scale);
                if (svgShape == null)
                {
                    continue;
                }

                svgShape.AlternativeText = CleanInlineText(ExtractFirstTagInnerHtml(svgBlock.InnerHtml, "title"));
                ApplyShapeHyperlink(svgShape, FindWrappingAnchorHref(mediaHtml, svgBlock.FullHtml));
                currentY = svgShape.Top + svgShape.Height + ScaleY(12f, scale);
            }

            foreach (Match mediaMatch in Regex.Matches(mediaHtml, "<(?<tag>video|audio|iframe|canvas)\\b(?<attrs>[^>]*)>", RegexOptions.IgnoreCase | RegexOptions.Singleline))
            {
                if (currentY >= area.Top + area.Height) break;
                string tagName = mediaMatch.Groups["tag"].Value.ToLowerInvariant();
                string openTag = "<" + tagName + mediaMatch.Groups["attrs"].Value + ">";
                string source = ReadAttribute(openTag, "src") ?? ReadAttribute(openTag, "data-src") ?? ReadAttribute(openTag, "href");
                string poster = ReadAttribute(openTag, "poster") ?? ReadAttribute(openTag, "data-poster") ?? ReadAttribute(openTag, "data-fallback") ?? ReadAttribute(openTag, "data-image");
                float height = tagName == "audio" ? ScaleY(48f, scale) : Math.Min(ScaleY(160f, scale), area.Top + area.Height - currentY);
                string posterPath = ResolveMediaSource(poster);
                Shape shape = null;
                if (!string.IsNullOrWhiteSpace(posterPath)) shape = AddMediaPicture(slide, posterPath, openTag, area.Left, currentY, area.Width, height, scale);
                if (shape == null)
                {
                    shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, area.Left, currentY, area.Width, height);
                    shape.Fill.ForeColor.RGB = ToRgb(tagName == "audio" ? "#eef3f8" : "#1f3846");
                    shape.Line.ForeColor.RGB = ToRgb("#94a3b8");
                    SetShapeText(shape, (tagName == "video" ? "▶" : tagName == "audio" ? "♪" : "↗") + "  " + tagName, 15f * scale.Average, tagName == "audio" ? "#2c3e50" : "#ffffff", true, PpParagraphAlignment.ppAlignCenter, Office.MsoVerticalAnchor.msoAnchorMiddle, null);
                }
                ApplyShapeHyperlink(shape, source);
                AddWarning(tagName + " was imported as a clickable static placeholder.", null);
                currentY = shape.Top + shape.Height + ScaleY(12f, scale);
            }

            return currentY - area.Top;
        }

        private Shape AddMediaPicture(Slide slide, string path, string sourceTag, float left, float top, float availableWidth, float availableHeight, SlideScale scale)
        {
            try
            {
                Shape picture = slide.Shapes.AddPicture(path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, left, top, -1f, -1f);
                picture.LockAspectRatio = Office.MsoTriState.msoTrue;

                float desiredWidth = ReadCssDimension(ReadStyle(ReadAttribute(sourceTag, "style"), "width") ?? ReadAttribute(sourceTag, "width"), availableWidth, scale.X);
                float desiredHeight = ReadCssDimension(ReadStyle(ReadAttribute(sourceTag, "style"), "height") ?? ReadAttribute(sourceTag, "height"), availableHeight, scale.Y);
                float maxWidth = desiredWidth > 0f ? Math.Min(desiredWidth, availableWidth) : availableWidth;
                float maxHeight = desiredHeight > 0f ? Math.Min(desiredHeight, availableHeight) : availableHeight;

                if (picture.Width > maxWidth) picture.Width = maxWidth;
                if (maxHeight > 0f && picture.Height > maxHeight) picture.Height = maxHeight;
                if (desiredWidth > 0f && picture.Width < desiredWidth) picture.Width = maxWidth;
                if (desiredHeight > 0f && picture.Height < desiredHeight) picture.Height = maxHeight;

                picture.Left = left + Math.Max(0f, (availableWidth - picture.Width) / 2f);
                picture.Top = top;
                return picture;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "HtmlPresentationImportService.AddMediaPicture");
                return null;
            }
        }

        private string ResolveMediaSource(string source)
        {
            if (string.IsNullOrWhiteSpace(source))
            {
                return null;
            }

            string decodedSource = WebUtility.HtmlDecode(source.Trim());
            try
            {
                if (decodedSource.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                {
                    Match dataMatch = Regex.Match(decodedSource, "^data:(?<mime>[^;,]+)?(?<base64>;base64)?,(?<data>.*)$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    if (!dataMatch.Success) return null;

                    string mime = dataMatch.Groups["mime"].Value;
                    string payload = dataMatch.Groups["data"].Value;
                    byte[] bytes = dataMatch.Groups["base64"].Success
                        ? Convert.FromBase64String(Regex.Replace(payload, "\\s+", string.Empty))
                        : Encoding.UTF8.GetBytes(Uri.UnescapeDataString(payload));
                    ConsumeMediaBudget(bytes);
                    return CreateTemporaryFile(GetExtensionForMimeType(mime), bytes);
                }

                Uri absoluteUri;
                if (Uri.TryCreate(decodedSource, UriKind.Absolute, out absoluteUri))
                {
                    if (absoluteUri.IsFile)
                    {
                        return File.Exists(absoluteUri.LocalPath) ? absoluteUri.LocalPath : null;
                    }

                    if (absoluteUri.Scheme == Uri.UriSchemeHttp || absoluteUri.Scheme == Uri.UriSchemeHttps)
                    {
                        if (!_allowRemoteResources)
                        {
                            ReportBlockedRemoteResource(absoluteUri.AbsoluteUri);
                            return null;
                        }
                        _mediaRequestCount++;
                        if (_mediaRequestCount > HtmlImportLimits.MaxMediaRequests)
                            throw new HtmlImportLimitException("The HTML import references more than 32 remote media resources.");
                        byte[] bytes = DownloadRemoteMedia(absoluteUri);
                        ConsumeMediaBudget(bytes);
                        return CreateTemporaryFile(Path.GetExtension(absoluteUri.AbsolutePath), bytes);
                    }
                }

                string localSource = decodedSource.Split(new[] { '?', '#' }, 2)[0].Replace('/', Path.DirectorySeparatorChar);
                string localPath = Path.IsPathRooted(localSource)
                    ? localSource
                    : Path.Combine(_htmlBaseDirectory ?? string.Empty, localSource);
                localPath = Path.GetFullPath(localPath);
                return File.Exists(localPath) ? localPath : null;
            }
            catch (HtmlImportLimitException)
            {
                throw;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "HtmlPresentationImportService.ResolveMediaSource");
                return null;
            }
        }

        private void ResetMediaBudget()
        {
            _mediaRequestCount = 0;
            _combinedMediaBytes = 0;
        }

        private void ConsumeMediaBudget(byte[] bytes)
        {
            int length = bytes?.Length ?? 0;
            if (length > HtmlImportLimits.MaxMediaBytes)
                throw new HtmlImportLimitException("A media resource exceeds the 10 MiB import limit.");
            if (_combinedMediaBytes > HtmlImportLimits.MaxCombinedMediaBytes - length)
                throw new HtmlImportLimitException("The combined media resources exceed the 50 MiB import limit.");
            _combinedMediaBytes += length;
        }

        private static byte[] DownloadRemoteMedia(Uri uri)
        {
            var request = HtmlImportResourceReader.CreateBoundedRequest(uri);
            try
            {
                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    if (uri.Scheme == Uri.UriSchemeHttps && response.ResponseUri.Scheme != Uri.UriSchemeHttps)
                        throw new HtmlImportLimitException("A remote media redirect attempted to downgrade HTTPS.");
                    if (response.ContentLength > HtmlImportLimits.MaxMediaBytes)
                        throw new HtmlImportLimitException("A media resource exceeds the 10 MiB import limit.");
                    using (var stream = response.GetResponseStream())
                    {
                        return HtmlImportResourceReader.ReadBytes(
                            stream,
                            HtmlImportLimits.MaxMediaBytes,
                            "A media resource exceeds the 10 MiB import limit.");
                    }
                }
            }
            catch (WebException ex) when (ex.Status == WebExceptionStatus.Timeout)
            {
                throw new HtmlImportLimitException("A remote media resource exceeded the 10-second download timeout.", ex);
            }
        }

        private void ReportBlockedRemoteResource(string resource)
        {
            if (_reportedBlockedRemoteResources.Add(resource ?? string.Empty))
            {
                AddWarning("A remote HTML resource was skipped because network access was not approved.", null);
            }
        }

        private string CreateTemporaryFile(string extension, byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
            {
                return null;
            }

            string normalizedExtension = string.IsNullOrWhiteSpace(extension) ? ".bin" : extension;
            if (!normalizedExtension.StartsWith(".", StringComparison.Ordinal)) normalizedExtension = "." + normalizedExtension;
            if (normalizedExtension.Length > 10 || normalizedExtension.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) normalizedExtension = ".bin";

            string path = Path.Combine(Path.GetTempPath(), "oPE_HtmlImport_" + Guid.NewGuid().ToString("N") + normalizedExtension);
            File.WriteAllBytes(path, bytes);
            _temporaryImportFiles.Add(path);
            return path;
        }

        private static string GetExtensionForMimeType(string mimeType)
        {
            switch ((mimeType ?? string.Empty).ToLowerInvariant())
            {
                case "image/svg+xml": return ".svg";
                case "image/jpeg": return ".jpg";
                case "image/gif": return ".gif";
                case "image/webp": return ".webp";
                case "image/bmp": return ".bmp";
                case "image/tiff": return ".tif";
                default: return ".png";
            }
        }

        private void CleanupTemporaryImportFiles()
        {
            foreach (string path in _temporaryImportFiles.ToList())
            {
                try
                {
                    if (File.Exists(path)) File.Delete(path);
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "HtmlPresentationImportService.CleanupTemporaryImportFiles");
                }
            }
            _temporaryImportFiles.Clear();
        }

        private static float ReadCssDimension(string value, float referenceSize, float scaleFactor)
        {
            if (string.IsNullOrWhiteSpace(value)) return 0f;
            string cleaned = value.Trim();
            float parsed;
            if (cleaned.EndsWith("%", StringComparison.Ordinal) && float.TryParse(cleaned.TrimEnd('%'), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out parsed))
            {
                return referenceSize * parsed / 100f;
            }

            cleaned = Regex.Replace(cleaned, "(px|pt)$", string.Empty, RegexOptions.IgnoreCase).Trim();
            return float.TryParse(cleaned, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out parsed)
                ? parsed * scaleFactor
                : 0f;
        }

        private static bool ContainsMedia(string html)
        {
            return !string.IsNullOrWhiteSpace(html) &&
                (html.IndexOf("<img", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 html.IndexOf("<svg", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 html.IndexOf("<video", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 html.IndexOf("<audio", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 html.IndexOf("<iframe", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 html.IndexOf("<canvas", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static bool ContainsTable(string html)
        {
            return !string.IsNullOrWhiteSpace(html) && html.IndexOf("<table", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void ApplyHtmlHyperlinks(Shape shape, string html)
        {
            if (shape == null || string.IsNullOrWhiteSpace(html) || html.IndexOf("<a", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return;
            }

            if (shape.HasTextFrame != Office.MsoTriState.msoTrue) return;

            string visibleText = shape.TextFrame.TextRange.Text ?? string.Empty;
            int searchStart = 0;
            MatchCollection anchors = Regex.Matches(html, "<a\\b(?<attrs>[^>]*)>(?<content>.*?)</a\\s*>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            foreach (Match anchor in anchors)
            {
                string href = ReadAttribute("<a" + anchor.Groups["attrs"].Value + ">", "href");
                string linkText = CleanInlineText(anchor.Groups["content"].Value);
                if (string.IsNullOrWhiteSpace(href) || string.IsNullOrWhiteSpace(linkText)) continue;

                int position = visibleText.IndexOf(linkText, searchStart, StringComparison.Ordinal);
                if (position < 0) continue;

                try
                {
                    TextRange range = shape.TextFrame.TextRange.Characters(position + 1, linkText.Length);
                    var actionSetting = range.ActionSettings[PpMouseActivation.ppMouseClick];
                    SetHyperlink(actionSetting, href);
                    range.Font.Underline = Office.MsoTriState.msoTrue;
                    range.Font.Color.RGB = ToRgb("#0563c1");
                    searchStart = position + linkText.Length;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "HtmlPresentationImportService.ApplyHtmlHyperlinks");
                }
            }
        }

        private void ApplyShapeHyperlink(Shape shape, string href)
        {
            if (shape == null || string.IsNullOrWhiteSpace(href)) return;
            try
            {
                SetHyperlink(shape.ActionSettings[PpMouseActivation.ppMouseClick], href);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "HtmlPresentationImportService.ApplyShapeHyperlink");
            }
        }

        private void SetHyperlink(ActionSetting actionSetting, string href)
        {
            if (actionSetting == null || string.IsNullOrWhiteSpace(href)) return;
            string decodedHref = WebUtility.HtmlDecode(href.Trim());
            actionSetting.Action = PpActionType.ppActionHyperlink;
            if (decodedHref.StartsWith("#", StringComparison.Ordinal))
            {
                actionSetting.Hyperlink.SubAddress = decodedHref.Substring(1);
                return;
            }

            actionSetting.Hyperlink.Address = ResolveLinkAddress(decodedHref);
        }

        private string ResolveLinkAddress(string href)
        {
            Uri uri;
            if (Uri.TryCreate(href, UriKind.Absolute, out uri)) return uri.AbsoluteUri;
            try
            {
                return new Uri(Path.GetFullPath(Path.Combine(_htmlBaseDirectory ?? string.Empty, href.Replace('/', Path.DirectorySeparatorChar)))).AbsoluteUri;
            }
            catch
            {
                return href;
            }
        }

        private static string FindWrappingAnchorHref(string html, string childHtml)
        {
            if (string.IsNullOrWhiteSpace(html) || string.IsNullOrWhiteSpace(childHtml)) return null;
            MatchCollection anchors = Regex.Matches(html, "<a\\b(?<attrs>[^>]*)>(?<content>.*?)</a\\s*>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            foreach (Match anchor in anchors)
            {
                if (anchor.Groups["content"].Value.IndexOf(childHtml, StringComparison.Ordinal) >= 0)
                {
                    return ReadAttribute("<a" + anchor.Groups["attrs"].Value + ">", "href");
                }
            }
            return null;
        }

        private static string GetAccentColor(bool useSourceStyling)
        {
            return useSourceStyling ? AccentHex : HeaderHex;
        }

        private void RenderDarkGradientBackground(Slide slide, float slideWidth, float slideHeight)
        {
            try
            {
                var background = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, slideWidth, slideHeight);
                background.Name = "oPE_HtmlImport_Background_" + Guid.NewGuid().ToString("N");
                background.Line.Visible = Office.MsoTriState.msoFalse;
                background.Fill.Visible = Office.MsoTriState.msoTrue;
                background.Fill.Transparency = 0f;
                background.Fill.ForeColor.RGB = ToRgb(DarkBackgroundStartHex);
                background.Fill.BackColor.RGB = ToRgb(DarkBackgroundAltHex);
                // CSS 135deg starts at the top-left and ends at the bottom-right.
                // PowerPoint's "Diagonal Up" preset maps to that direction; "Diagonal Down"
                // mirrors the source horizontally on the generated rectangle.
                background.Fill.TwoColorGradient(Office.MsoGradientStyle.msoGradientDiagonalUp, 1);

                try
                {
                    if (background.Fill.GradientStops.Count >= 1)
                    {
                        background.Fill.GradientStops[1].Color.RGB = ToRgb(DarkBackgroundStartHex);
                        background.Fill.GradientStops[1].Position = 0f;
                        background.Fill.GradientStops[1].Transparency = 0f;
                    }

                    if (background.Fill.GradientStops.Count >= 2)
                    {
                        background.Fill.GradientStops[2].Color.RGB = ToRgb(DarkBackgroundAltHex);
                        background.Fill.GradientStops[2].Position = 1f;
                        background.Fill.GradientStops[2].Transparency = 0f;
                    }
                }
                catch
                {
                    // Keep the native two-color gradient even if stop-level access fails.
                }

                background.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
            }
            catch
            {
                var background = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, slideWidth, slideHeight);
                background.Name = "oPE_HtmlImport_Background_" + Guid.NewGuid().ToString("N");
                background.Line.Visible = Office.MsoTriState.msoFalse;
                background.Fill.ForeColor.RGB = ToRgb("#264652");
                background.Fill.Solid();
                background.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
            }
        }

        private void ApplyTitleTextGradient(Shape titleShape)
        {
            if (titleShape == null)
            {
                return;
            }

            try
            {
                dynamic fontFill = titleShape.TextFrame2.TextRange.Font.Fill;
                fontFill.Visible = Office.MsoTriState.msoTrue;
                fontFill.Transparency = 0f;
                fontFill.TwoColorGradient(Office.MsoGradientStyle.msoGradientHorizontal, 1);
                fontFill.GradientStops[1].Color.RGB = ToRgb(AccentHex);
                fontFill.GradientStops[1].Position = 0f;
                fontFill.GradientStops[1].Transparency = 0f;
                fontFill.GradientStops[fontFill.GradientStops.Count].Color.RGB = ToRgb(TitleGradientEndHex);
                fontFill.GradientStops[fontFill.GradientStops.Count].Position = 1f;
                fontFill.GradientStops[fontFill.GradientStops.Count].Transparency = 0f;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "HtmlPresentationImportService.ApplyTitleTextGradient");
                try
                {
                    titleShape.TextFrame.TextRange.Font.Color.RGB = ToRgb(AccentHex);
                }
                catch (Exception formattingEx) { ReportOptionalImportFailure(formattingEx, "HtmlPresentationImportService.ApplyDefaultTitleStyle"); }
            }
        }

        private List<HtmlSlideModel> LoadSlides(string htmlFilePath)
        {
            HtmlImportDocument document = new HtmlImportDocumentPipeline(_allowRemoteResources).Load(htmlFilePath);
            foreach (string warning in document.Warnings) AddWarning(warning, null);
            return document.Slides.Select(slide => new HtmlSlideModel
            {
                Number = slide.Number,
                Id = slide.Id,
                IsGptDeck = slide.IsGptDeck,
                IsGenericDom = slide.IsGenericDom,
                SectionOpenTag = slide.SectionOpenTag,
                StyleSheet = slide.StyleSheet,
                Classes = slide.Classes,
                IsDark = slide.IsDark,
                RawHtml = slide.RawHtml,
                HeaderHtml = slide.HeaderHtml,
                ContentHtml = slide.ContentHtml,
                FooterHtml = slide.FooterHtml,
                TitleSlideHtml = slide.TitleSlideHtml,
                Title = slide.Title,
                FooterLeft = slide.FooterLeft,
                FooterRight = slide.FooterRight
            }).ToList();
        }

        private void AddWarning(string message, int? slideNumber)
        {
            LastReport?.AddWarning(message, slideNumber);
        }

        private static void ReportOptionalImportFailure(Exception exception, string context)
        {
            if (!ExceptionPolicy.IsExpectedComFailure(exception))
                ExceptionLogger.Log(exception, context);
            // Optional formatting or rollback continues so the primary import result remains diagnosable.
        }

        private static string GetTitleColor(bool useSourceStyling)
        {
            return useSourceStyling ? AccentHex : HeaderHex;
        }

        private static string GetSubtitleColor(bool isDark, bool useSourceStyling)
        {
            return useSourceStyling ? (isDark ? "#b9c2c9" : BodySecondaryHex) : BodySecondaryHex;
        }

        private static string GetHeaderColor(bool isDark, bool useSourceStyling)
        {
            return useSourceStyling ? (isDark ? DarkTextHex : HeaderHex) : HeaderHex;
        }

        private static string GetSmallHeadingColor(bool isDark, bool useSourceStyling)
        {
            return useSourceStyling ? (isDark ? AccentHex : HeaderHex) : HeaderHex;
        }

        private static string GetFooterColor(bool isDark, bool useSourceStyling)
        {
            return useSourceStyling ? (isDark ? "#8ea0ad" : FooterHex) : FooterHex;
        }

        private static string GetDividerColor(bool isDark, bool useSourceStyling)
        {
            return useSourceStyling ? (isDark ? "#506672" : "#e2e8f0") : "#d7dde5";
        }

        private static string GetBodyColor(bool isDark, bool useSourceStyling)
        {
            return useSourceStyling ? (isDark ? DarkBodySecondaryHex : BodySecondaryHex) : BodySecondaryHex;
        }

        private static string GetTextColor(bool isDark, bool useSourceStyling)
        {
            return useSourceStyling ? (isDark ? DarkTextHex : LightTextHex) : LightTextHex;
        }

        private static string GetBadgeBackgroundColor(bool isDark, bool useSourceStyling)
        {
            return useSourceStyling ? (isDark ? "#4d5d4d" : "#fff4df") : "#eef3f8";
        }

        private static string GetFontName(HtmlImportOptions options)
        {
            if (options != null && options.UsePresentationTheme) return null;
            return options != null && options.PreserveSourceFonts ? SourceBodyFont : null;
        }

        private Shape AddTextBox(Slide slide, float left, float top, float width, float height, string text, float fontSize, string colorHex, bool bold, PpParagraphAlignment alignment, Office.MsoVerticalAnchor verticalAnchor, string fontName)
        {
            var shape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            shape.Line.Visible = Office.MsoTriState.msoFalse;
            shape.Fill.Visible = Office.MsoTriState.msoFalse;
            shape.TextFrame.MarginLeft = 0;
            shape.TextFrame.MarginRight = 0;
            shape.TextFrame.MarginTop = 0;
            shape.TextFrame.MarginBottom = 0;
            shape.TextFrame.VerticalAnchor = verticalAnchor;
            shape.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
            SetShapeText(shape, text, fontSize, colorHex, bold, alignment, verticalAnchor, fontName);
            return shape;
        }

        private void SetShapeText(Shape shape, string text, float fontSize, string colorHex, bool bold, PpParagraphAlignment alignment, Office.MsoVerticalAnchor verticalAnchor, string fontName)
        {
            shape.TextFrame.TextRange.Text = text ?? string.Empty;
            string resolvedFont = _presentationTheme == null
                ? fontName
                : (bold ? _presentationTheme.HeadingFont : _presentationTheme.BodyFont);
            if (!string.IsNullOrWhiteSpace(resolvedFont))
            {
                shape.TextFrame.TextRange.Font.Name = resolvedFont;
            }
            shape.TextFrame.TextRange.Font.Size = Math.Max(1f, (float)Math.Round(fontSize, MidpointRounding.AwayFromZero));
            shape.TextFrame.TextRange.Font.Bold = bold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            shape.TextFrame.TextRange.Font.Color.RGB = ToRgb(string.IsNullOrWhiteSpace(colorHex) ? LightTextHex : colorHex);
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = alignment;
            shape.TextFrame.VerticalAnchor = verticalAnchor;
        }

        private void ApplyHtmlBoldFormatting(Shape shape, string html)
        {
            if (shape == null || string.IsNullOrWhiteSpace(html))
            {
                return;
            }

            string visibleText = shape.TextFrame.TextRange.Text ?? string.Empty;
            int searchStart = 0;
            MatchCollection matches = Regex.Matches(html, "<(?:strong|b)\\b[^>]*>(.*?)</(?:strong|b)\\s*>", RegexOptions.IgnoreCase | RegexOptions.Singleline);

            foreach (Match match in matches)
            {
                string boldText = CleanInlineText(match.Groups[1].Value);
                if (string.IsNullOrWhiteSpace(boldText))
                {
                    continue;
                }

                int position = visibleText.IndexOf(boldText, searchStart, StringComparison.Ordinal);
                if (position < 0)
                {
                    continue;
                }

                try
                {
                    shape.TextFrame.TextRange.Characters(position + 1, boldText.Length).Font.Bold = Office.MsoTriState.msoTrue;
                    searchStart = position + boldText.Length;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "HtmlPresentationImportService.ApplyHtmlBoldFormatting");
                }
            }
        }

        private void ApplyHtmlItalicFormatting(Shape shape, string html)
        {
            if (shape == null || string.IsNullOrWhiteSpace(html))
            {
                return;
            }

            string visibleText = shape.TextFrame.TextRange.Text ?? string.Empty;
            int searchStart = 0;
            MatchCollection matches = Regex.Matches(html, "<(?:em|i)\\b[^>]*>(.*?)</(?:em|i)\\s*>", RegexOptions.IgnoreCase | RegexOptions.Singleline);

            foreach (Match match in matches)
            {
                string italicText = CleanInlineText(match.Groups[1].Value);
                if (string.IsNullOrWhiteSpace(italicText))
                {
                    continue;
                }

                int position = visibleText.IndexOf(italicText, searchStart, StringComparison.Ordinal);
                if (position < 0)
                {
                    continue;
                }

                try
                {
                    shape.TextFrame.TextRange.Characters(position + 1, italicText.Length).Font.Italic = Office.MsoTriState.msoTrue;
                    searchStart = position + italicText.Length;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "HtmlPresentationImportService.ApplyHtmlItalicFormatting");
                }
            }
        }

        private static string ExtractBlockTextByClass(string html, string tagName, string className)
        {
            string block = ExtractBlockByClass(html, tagName, className);
            return CleanInlineText(block);
        }

        private static string ExtractBlockByClass(string html, string tagName, string className)
        {
            HtmlBlock block = FindBlocksByClass(html, tagName, className).FirstOrDefault();
            return block != null ? block.InnerHtml : null;
        }

        private static string ExtractFirstTagInnerHtml(string html, string tagName)
        {
            HtmlBlock block = FindTagBlocks(html, tagName).FirstOrDefault();
            return block != null ? block.InnerHtml : null;
        }

        private static string ExtractFirstTagBlock(string html, string tagName)
        {
            HtmlBlock block = FindTagBlocks(html, tagName).FirstOrDefault();
            return block != null ? block.FullHtml : null;
        }

        private static List<string> ExtractTagContents(string html, string tagName)
        {
            return FindTagBlocks(html, tagName).Select(b => b.InnerHtml).ToList();
        }

        private static List<HtmlBlock> FindDirectChildBlocks(string html, string tagName)
        {
            return FindBlocks(html, tagName, null, true);
        }

        private static List<HtmlBlock> FindDirectChildTagBlocks(string html, string tagName)
        {
            return FindBlocks(html, tagName, null, true);
        }

        private static List<HtmlBlock> FindBlocksByClass(string html, string tagName, string className)
        {
            return FindBlocks(html, tagName, className, false);
        }

        private static List<HtmlBlock> FindTagBlocks(string html, string tagName)
        {
            return FindBlocks(html, tagName, null, false);
        }

        private static List<HtmlBlock> FindBlocks(string html, string tagName, string requiredClass, bool directChildrenOnly)
        {
            var blocks = new List<HtmlBlock>();
            if (string.IsNullOrWhiteSpace(html))
            {
                return blocks;
            }

            string openToken = "<" + tagName;
            string closeToken = "</" + tagName;
            int index = 0;
            int rootDepth = 0;

            while (index < html.Length)
            {
                int nextOpen = IndexOfIgnoreCase(html, openToken, index);
                int nextClose = IndexOfIgnoreCase(html, closeToken, index);
                if (nextOpen == -1 && nextClose == -1)
                {
                    break;
                }

                bool isOpen = nextOpen != -1 && (nextClose == -1 || nextOpen < nextClose);
                if (!isOpen)
                {
                    rootDepth = Math.Max(0, rootDepth - 1);
                    int closeEnd = html.IndexOf('>', nextClose);
                    index = closeEnd == -1 ? html.Length : closeEnd + 1;
                    continue;
                }

                int tagEnd = html.IndexOf('>', nextOpen);
                if (tagEnd == -1)
                {
                    break;
                }

                string openTag = html.Substring(nextOpen, tagEnd - nextOpen + 1);
                bool selfClosing = openTag.EndsWith("/>", StringComparison.Ordinal);

                if (!selfClosing)
                {
                    int depthBeforeOpen = rootDepth;
                    rootDepth++;

                    bool classMatches = string.IsNullOrWhiteSpace(requiredClass) || ClassContains(openTag, requiredClass);
                    bool depthMatches = !directChildrenOnly || depthBeforeOpen == 0;
                    if (classMatches && depthMatches)
                    {
                        int endIndex = FindMatchingClosingTag(html, tagName, nextOpen);
                        if (endIndex > tagEnd)
                        {
                            int closingStart = endIndex - (tagName.Length + 3);
                            string innerHtml = html.Substring(tagEnd + 1, closingStart - tagEnd - 1);
                            string fullHtml = html.Substring(nextOpen, endIndex - nextOpen);

                            blocks.Add(new HtmlBlock
                            {
                                TagName = tagName,
                                OpenTag = openTag,
                                InnerHtml = innerHtml,
                                FullHtml = fullHtml,
                                Classes = SplitClasses(ReadAttribute(openTag, "class")),
                                Style = ReadAttribute(openTag, "style") ?? string.Empty
                            });
                        }
                    }
                }

                index = tagEnd + 1;
            }

            return blocks;
        }

        private static int FindMatchingClosingTag(string html, string tagName, int openingTagIndex)
        {
            string openToken = "<" + tagName;
            string closeToken = "</" + tagName;
            int openEnd = html.IndexOf('>', openingTagIndex);
            if (openEnd == -1)
            {
                return -1;
            }

            int depth = 1;
            int index = openEnd + 1;

            while (index < html.Length)
            {
                int nextOpen = IndexOfIgnoreCase(html, openToken, index);
                int nextClose = IndexOfIgnoreCase(html, closeToken, index);

                if (nextClose == -1)
                {
                    return -1;
                }

                if (nextOpen != -1 && nextOpen < nextClose)
                {
                    int nestedOpenEnd = html.IndexOf('>', nextOpen);
                    if (nestedOpenEnd == -1)
                    {
                        return -1;
                    }

                    string nestedOpenTag = html.Substring(nextOpen, nestedOpenEnd - nextOpen + 1);
                    if (!nestedOpenTag.EndsWith("/>", StringComparison.Ordinal))
                    {
                        depth++;
                    }
                    index = nestedOpenEnd + 1;
                    continue;
                }

                depth--;
                int closeEnd = html.IndexOf('>', nextClose);
                if (closeEnd == -1)
                {
                    return -1;
                }

                if (depth == 0)
                {
                    return closeEnd + 1;
                }

                index = closeEnd + 1;
            }

            return -1;
        }

        private static int IndexOfIgnoreCase(string value, string search, int startIndex)
        {
            return value.IndexOf(search, startIndex, StringComparison.OrdinalIgnoreCase);
        }

        private static bool ClassContains(HtmlBlock block, string requiredClass)
        {
            return block != null && block.Classes.Any(c => string.Equals(c, requiredClass, StringComparison.OrdinalIgnoreCase));
        }

        private static bool ClassContains(string openTag, string requiredClass)
        {
            string classValue = ReadAttribute(openTag, "class");
            return SplitClasses(classValue).Any(c => string.Equals(c, requiredClass, StringComparison.OrdinalIgnoreCase));
        }

        private static List<string> SplitClasses(string classes)
        {
            return string.IsNullOrWhiteSpace(classes)
                ? new List<string>()
                : classes.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(c => c.Trim()).ToList();
        }

        private static string ReadAttribute(string openTag, string attributeName)
        {
            if (string.IsNullOrWhiteSpace(openTag) || string.IsNullOrWhiteSpace(attributeName))
            {
                return null;
            }

            Match match = Regex.Match(
                openTag,
                "(?:^|\\s)" + Regex.Escape(attributeName) + "\\s*=\\s*(?:\"(?<double>[^\"]*)\"|'(?<single>[^']*)'|(?<bare>[^\\s>]+))",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!match.Success) return null;
            if (match.Groups["double"].Success) return match.Groups["double"].Value;
            if (match.Groups["single"].Success) return match.Groups["single"].Value;
            return match.Groups["bare"].Value;
        }

        private static string ReadStyle(string style, string propertyName)
        {
            if (string.IsNullOrWhiteSpace(style) || string.IsNullOrWhiteSpace(propertyName))
            {
                return null;
            }

            Match match = Regex.Match(style, Regex.Escape(propertyName) + "\\s*:\\s*([^;]+)", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value.Trim() : null;
        }

        private static HtmlDomNode ParseDomFragment(string html)
        {
            var root = new HtmlDomNode { TagName = "root", InnerHtml = html ?? string.Empty, FullHtml = html ?? string.Empty };
            var stack = new Stack<HtmlDomNode>();
            stack.Push(root);
            string source = html ?? string.Empty;
            MatchCollection tags = Regex.Matches(source, "<!--[\\s\\S]*?-->|<![^>]*>|<(?<close>/)?(?<tag>[a-zA-Z][a-zA-Z0-9:-]*)(?<attrs>[^>]*)>", RegexOptions.Singleline);
            int cursor = 0;

            foreach (Match tagMatch in tags)
            {
                if (tagMatch.Index > cursor)
                {
                    string textValue = source.Substring(cursor, tagMatch.Index - cursor);
                    if (!string.IsNullOrWhiteSpace(WebUtility.HtmlDecode(textValue)))
                    {
                        var textNode = new HtmlDomNode { TagName = "#text", Text = textValue, Parent = stack.Peek(), FullHtml = textValue, InnerHtml = textValue };
                        stack.Peek().Children.Add(textNode);
                    }
                }

                cursor = tagMatch.Index + tagMatch.Length;
                if (!tagMatch.Groups["tag"].Success) continue;
                string tagName = tagMatch.Groups["tag"].Value.ToLowerInvariant();
                bool isClosing = tagMatch.Groups["close"].Success;

                if (isClosing)
                {
                    while (stack.Count > 1)
                    {
                        HtmlDomNode completed = stack.Pop();
                        completed.InnerHtml = source.Substring(completed.ContentStart, Math.Max(0, tagMatch.Index - completed.ContentStart));
                        completed.FullHtml = source.Substring(completed.Start, Math.Max(0, cursor - completed.Start));
                        if (string.Equals(completed.TagName, tagName, StringComparison.OrdinalIgnoreCase)) break;
                    }
                    continue;
                }

                var node = new HtmlDomNode
                {
                    TagName = tagName,
                    OpenTag = tagMatch.Value,
                    Start = tagMatch.Index,
                    ContentStart = cursor,
                    Parent = stack.Peek(),
                    Attributes = ParseHtmlAttributes(tagMatch.Value)
                };
                node.Classes = SplitClasses(node.GetAttribute("class"));
                stack.Peek().Children.Add(node);

                bool selfClosing = tagMatch.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal) || IsVoidHtmlTag(tagName);
                if (selfClosing)
                {
                    node.InnerHtml = string.Empty;
                    node.FullHtml = tagMatch.Value;
                }
                else
                {
                    stack.Push(node);
                }
            }

            while (stack.Count > 1)
            {
                HtmlDomNode node = stack.Pop();
                node.InnerHtml = source.Substring(node.ContentStart);
                node.FullHtml = source.Substring(node.Start);
            }
            return root;
        }

        private static Dictionary<string, string> ParseHtmlAttributes(string openTag)
        {
            var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            MatchCollection matches = Regex.Matches(openTag ?? string.Empty, "(?<name>[a-zA-Z_:][a-zA-Z0-9_:.\\-]*)\\s*=\\s*(?:\"(?<double>[^\"]*)\"|'(?<single>[^']*)'|(?<bare>[^\\s>]+))", RegexOptions.Singleline);
            foreach (Match match in matches)
            {
                string value = match.Groups["double"].Success ? match.Groups["double"].Value : match.Groups["single"].Success ? match.Groups["single"].Value : match.Groups["bare"].Value;
                attributes[match.Groups["name"].Value] = WebUtility.HtmlDecode(value);
            }
            return attributes;
        }

        private static bool IsVoidHtmlTag(string tagName)
        {
            switch (tagName)
            {
                case "img": case "br": case "hr": case "input": case "meta": case "link": case "source": case "area": case "base": case "embed": case "param": case "track": case "wbr":
                    return true;
                default:
                    return false;
            }
        }

        private sealed class CssStyleResolver
        {
            private readonly List<CssRule> _rules = new List<CssRule>();
            private readonly Dictionary<string, string> _variables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<HtmlDomNode, Dictionary<string, string>> _cache = new Dictionary<HtmlDomNode, Dictionary<string, string>>();

            public CssStyleResolver(string styleSheet)
            {
                string css = Regex.Replace(styleSheet ?? string.Empty, "/\\*[\\s\\S]*?\\*/", string.Empty);
                MatchCollection rules = Regex.Matches(css, "(?<selectors>[^{}]+)\\{(?<body>[^{}]*)\\}", RegexOptions.Singleline);
                int order = 0;
                foreach (Match ruleMatch in rules)
                {
                    Dictionary<string, string> declarations = ParseDeclarations(ruleMatch.Groups["body"].Value);
                    foreach (var declaration in declarations.Where(item => item.Key.StartsWith("--", StringComparison.Ordinal))) _variables[declaration.Key] = declaration.Value;
                    foreach (string selector in ruleMatch.Groups["selectors"].Value.Split(','))
                    {
                        string cleanedSelector = selector.Trim();
                        if (string.IsNullOrWhiteSpace(cleanedSelector) || cleanedSelector.StartsWith("@", StringComparison.Ordinal)) continue;
                        _rules.Add(new CssRule { Selector = cleanedSelector, Declarations = declarations, Order = order++, Specificity = CalculateSpecificity(cleanedSelector) });
                    }
                }
            }

            public Dictionary<string, string> GetStyle(HtmlDomNode node)
            {
                Dictionary<string, string> cached;
                if (_cache.TryGetValue(node, out cached)) return cached;
                var style = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (node.Parent != null)
                {
                    Dictionary<string, string> parent = GetStyle(node.Parent);
                    foreach (string inherited in new[] { "color", "font-family", "font-size", "font-weight", "font-style", "line-height", "text-align", "text-transform" })
                    {
                        string value;
                        if (parent.TryGetValue(inherited, out value)) style[inherited] = value;
                    }
                }

                foreach (CssRule rule in _rules.Where(rule => MatchesSelector(node, rule.Selector)).OrderBy(rule => rule.Specificity).ThenBy(rule => rule.Order))
                {
                    foreach (var declaration in rule.Declarations) style[declaration.Key] = ResolveVariables(declaration.Value);
                }
                foreach (var declaration in ParseDeclarations(node.GetAttribute("style"))) style[declaration.Key] = ResolveVariables(declaration.Value);
                _cache[node] = style;
                return style;
            }

            private string ResolveVariables(string value)
            {
                return Regex.Replace(value ?? string.Empty, "var\\(\\s*(--[a-zA-Z0-9_-]+)\\s*(?:,([^)]*))?\\)", match =>
                {
                    string replacement;
                    return _variables.TryGetValue(match.Groups[1].Value, out replacement) ? replacement : match.Groups[2].Value.Trim();
                });
            }

            private static Dictionary<string, string> ParseDeclarations(string declarations)
            {
                var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (string part in (declarations ?? string.Empty).Split(';'))
                {
                    int colon = part.IndexOf(':');
                    if (colon <= 0) continue;
                    result[part.Substring(0, colon).Trim()] = part.Substring(colon + 1).Trim().Replace("!important", string.Empty).Trim();
                }
                return result;
            }

            private static int CalculateSpecificity(string selector)
            {
                return Regex.Matches(selector, "#").Count * 100 + Regex.Matches(selector, "\\.").Count * 10 + Regex.Matches(selector, "(^|[ >+~])([a-zA-Z][a-zA-Z0-9-]*)").Count;
            }

            private static bool MatchesSelector(HtmlDomNode node, string selector)
            {
                string cleaned = Regex.Replace(selector, ":[a-zA-Z-]+(?:\\([^)]*\\))?", string.Empty).Trim();
                string[] parts = Regex.Split(cleaned.Replace(">", " "), "\\s+").Where(part => !string.IsNullOrWhiteSpace(part)).ToArray();
                if (parts.Length == 0) return false;
                if (!MatchesSimpleSelector(node, parts[parts.Length - 1])) return false;

                HtmlDomNode current = node.Parent;
                for (int index = parts.Length - 2; index >= 0; index--)
                {
                    while (current != null && !MatchesSimpleSelector(current, parts[index])) current = current.Parent;
                    if (current == null) return false;
                    current = current.Parent;
                }
                return true;
            }

            private static bool MatchesSimpleSelector(HtmlDomNode node, string selector)
            {
                if (node == null || node.TagName == "#text") return false;
                Match idMatch = Regex.Match(selector, "#([a-zA-Z0-9_-]+)");
                if (idMatch.Success && !string.Equals(node.GetAttribute("id"), idMatch.Groups[1].Value, StringComparison.OrdinalIgnoreCase)) return false;
                foreach (Match classMatch in Regex.Matches(selector, "\\.([a-zA-Z0-9_-]+)"))
                {
                    if (!node.Classes.Any(value => string.Equals(value, classMatch.Groups[1].Value, StringComparison.OrdinalIgnoreCase))) return false;
                }
                Match tagMatch = Regex.Match(selector, "^[a-zA-Z][a-zA-Z0-9-]*");
                return !tagMatch.Success || string.Equals(node.TagName, tagMatch.Value, StringComparison.OrdinalIgnoreCase);
            }
        }

        private sealed class CssRule
        {
            public string Selector { get; set; }
            public Dictionary<string, string> Declarations { get; set; }
            public int Specificity { get; set; }
            public int Order { get; set; }
        }

        private sealed class HtmlDomNode
        {
            public HtmlDomNode()
            {
                Children = new List<HtmlDomNode>();
                Attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                Classes = new List<string>();
            }
            public string TagName { get; set; }
            public string OpenTag { get; set; }
            public string InnerHtml { get; set; }
            public string FullHtml { get; set; }
            public string Text { get; set; }
            public int Start { get; set; }
            public int ContentStart { get; set; }
            public HtmlDomNode Parent { get; set; }
            public List<HtmlDomNode> Children { get; private set; }
            public Dictionary<string, string> Attributes { get; set; }
            public List<string> Classes { get; set; }
            public string GetAttribute(string name) { string value; return Attributes.TryGetValue(name, out value) ? value : null; }
            public string GetText() { return TagName == "#text" ? WebUtility.HtmlDecode(Text ?? string.Empty) : string.Join(" ", Children.Select(child => child.GetText())).Trim(); }
        }

        private static string CleanInlineText(string html)
        {
            if (string.IsNullOrWhiteSpace(html))
            {
                return string.Empty;
            }

            string withLineBreaks = Regex.Replace(html, "<br\\s*/?>", "\n", RegexOptions.IgnoreCase);
            string withoutTags = Regex.Replace(withLineBreaks, "<.*?>", string.Empty, RegexOptions.Singleline);
            string decoded = WebUtility.HtmlDecode(withoutTags);
            decoded = decoded.Replace("\r", string.Empty);
            decoded = Regex.Replace(decoded, "[ \t]+", " ");
            decoded = Regex.Replace(decoded, "\\n\\s+", "\n");
            return decoded.Trim();
        }

        private static bool TryParseCssPixels(string cssValue, out float pixels)
        {
            pixels = 0f;
            if (string.IsNullOrWhiteSpace(cssValue))
            {
                return false;
            }

            string cleaned = cssValue.Replace("px", string.Empty).Trim();
            return float.TryParse(cleaned, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out pixels);
        }

        private static float CssPixelsToPoints(float pixels)
        {
            return pixels * 0.75f;
        }

        private int ToRgb(string colorHex)
        {
            if (string.IsNullOrWhiteSpace(colorHex))
            {
                return ColorTranslator.ToOle(Color.Black);
            }

            try
            {
                string mappedColor = _presentationTheme == null ? colorHex : _presentationTheme.MapColor(colorHex);
                return ColorTranslator.ToOle(ColorTranslator.FromHtml(mappedColor.Trim()));
            }
            catch
            {
                return ColorTranslator.ToOle(Color.Black);
            }
        }

        private static float ScaleX(float value, SlideScale scale)
        {
            return value * scale.X;
        }

        private static float ScaleY(float value, SlideScale scale)
        {
            return value * scale.Y;
        }

        private static float MapX(float value, SlideScale scale)
        {
            return scale.OffsetX + ScaleX(value, scale);
        }

        private static float MapY(float value, SlideScale scale)
        {
            return scale.OffsetY + ScaleY(value, scale);
        }

        private class HtmlSlideModel
        {
            public int Number { get; set; }
            public string Id { get; set; }
            public bool IsGptDeck { get; set; }
            public bool IsGenericDom { get; set; }
            public string SectionOpenTag { get; set; }
            public string StyleSheet { get; set; }
            public List<string> Classes { get; set; }
            public bool IsDark { get; set; }
            public string RawHtml { get; set; }
            public string HeaderHtml { get; set; }
            public string ContentHtml { get; set; }
            public string FooterHtml { get; set; }
            public string TitleSlideHtml { get; set; }
            public string Title { get; set; }
            public string FooterLeft { get; set; }
            public string FooterRight { get; set; }
        }

        private class HtmlBlock
        {
            public string TagName { get; set; }
            public string OpenTag { get; set; }
            public string InnerHtml { get; set; }
            public string FullHtml { get; set; }
            public List<string> Classes { get; set; }
            public string Style { get; set; }
        }

        private sealed class FlexItem
        {
            public HtmlDomNode Node { get; set; }
            public Dictionary<string, string> Style { get; set; }
            public float MainSize { get; set; }
            public float CrossSize { get; set; }
            public float Grow { get; set; }
        }

        private sealed class TableLayout
        {
            public TableLayout() { Cells = new List<TableCellPlacement>(); }
            public int RowCount { get; set; }
            public int ColumnCount { get; set; }
            public List<TableCellPlacement> Cells { get; private set; }
        }

        private sealed class TableCellPlacement
        {
            public int Row { get; set; }
            public int Column { get; set; }
            public int RowSpan { get; set; }
            public int ColumnSpan { get; set; }
            public HtmlBlock Block { get; set; }
        }

        private class SlideScale
        {
            public SlideScale(float x, float y, float offsetX, float offsetY)
            {
                X = x;
                Y = y;
                OffsetX = offsetX;
                OffsetY = offsetY;
                Average = (x + y) / 2f;
            }

            public float X { get; private set; }
            public float Y { get; private set; }
            public float OffsetX { get; private set; }
            public float OffsetY { get; private set; }
            public float Average { get; private set; }
        }

        private class LayoutRect
        {
            public LayoutRect(float left, float top, float width, float height)
            {
                Left = left;
                Top = top;
                Width = width;
                Height = height;
            }

            public float Left { get; private set; }
            public float Top { get; private set; }
            public float Width { get; private set; }
            public float Height { get; private set; }
        }

        /// <summary>Resolved colors and fonts from the active PowerPoint master theme.</summary>
        private sealed class PresentationThemePalette
        {
            private PresentationThemePalette() { }

            public string Dark1 { get; private set; }
            public string Light1 { get; private set; }
            public string Dark2 { get; private set; }
            public string Light2 { get; private set; }
            public string Accent1 { get; private set; }
            public string Accent2 { get; private set; }
            public string Accent3 { get; private set; }
            public string Accent4 { get; private set; }
            public string Accent5 { get; private set; }
            public string Accent6 { get; private set; }
            public string HeadingFont { get; private set; }
            public string BodyFont { get; private set; }

            public static PresentationThemePalette TryCreate(Presentation presentation)
            {
                try
                {
                    var scheme = presentation.SlideMaster.Theme.ThemeColorScheme;
                    var fonts = presentation.SlideMaster.Theme.ThemeFontScheme;
                    var palette = new PresentationThemePalette
                    {
                        Dark1 = ReadColor(scheme, Office.MsoThemeColorSchemeIndex.msoThemeDark1, "#1f2937"),
                        Light1 = ReadColor(scheme, Office.MsoThemeColorSchemeIndex.msoThemeLight1, "#ffffff"),
                        Dark2 = ReadColor(scheme, Office.MsoThemeColorSchemeIndex.msoThemeDark2, "#334155"),
                        Light2 = ReadColor(scheme, Office.MsoThemeColorSchemeIndex.msoThemeLight2, "#f8fafc"),
                        Accent1 = ReadColor(scheme, Office.MsoThemeColorSchemeIndex.msoThemeAccent1, "#2563eb"),
                        Accent2 = ReadColor(scheme, Office.MsoThemeColorSchemeIndex.msoThemeAccent2, "#0f766e"),
                        Accent3 = ReadColor(scheme, Office.MsoThemeColorSchemeIndex.msoThemeAccent3, "#7c3aed"),
                        Accent4 = ReadColor(scheme, Office.MsoThemeColorSchemeIndex.msoThemeAccent4, "#dc2626"),
                        Accent5 = ReadColor(scheme, Office.MsoThemeColorSchemeIndex.msoThemeAccent5, "#d97706"),
                        Accent6 = ReadColor(scheme, Office.MsoThemeColorSchemeIndex.msoThemeAccent6, "#0891b2"),
                        HeadingFont = ReadFont(fonts.MajorFont, DefaultFont),
                        BodyFont = ReadFont(fonts.MinorFont, DefaultFont)
                    };
                    return palette;
                }
                catch
                {
                    return null;
                }
            }

            public string MapColor(string sourceColor)
            {
                try
                {
                    Color color = ColorTranslator.FromHtml(sourceColor.Trim());
                    float saturation = color.GetSaturation();
                    float brightness = color.GetBrightness();
                    if (brightness <= 0.18f) return Dark1;
                    if (brightness <= 0.42f) return Dark2;
                    if (brightness >= 0.92f) return Light1;
                    if (saturation < 0.12f)
                    {
                        if (brightness >= 0.62f) return Light2;
                        return Dark2;
                    }

                    float hue = color.GetHue();
                    if (hue < 18f || hue >= 340f) return Accent2;
                    if (hue < 68f) return Accent1;
                    if (hue < 165f) return Accent3;
                    if (hue < 218f) return Accent4;
                    if (hue < 270f) return Accent5;
                    return Accent6;
                }
                catch
                {
                    return sourceColor;
                }
            }

            private static string ReadColor(Office.ThemeColorScheme scheme, Office.MsoThemeColorSchemeIndex index, string fallback)
            {
                try { return ToHex(ColorTranslator.FromOle(scheme.Colors(index).RGB)); }
                catch { return fallback; }
            }

            private static string ReadFont(Office.ThemeFonts fonts, string fallback)
            {
                try
                {
                    string name = fonts.Item(Office.MsoFontLanguageIndex.msoThemeLatin).Name;
                    return string.IsNullOrWhiteSpace(name) ? fallback : name;
                }
                catch { return fallback; }
            }

            private static string ToHex(Color color)
            {
                return "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
            }
        }
    }

    public sealed class HtmlImportOptions
    {
        public HtmlImportOptions()
        {
            SelectedSlideNumbers = new HashSet<int>();
            PreserveSourceStyling = true;
            PreserveSourceBackground = true;
            PreserveSourceFonts = true;
            UsePresentationTheme = false;
        }

        public HashSet<int> SelectedSlideNumbers { get; private set; }
        public bool PreserveSourceStyling { get; set; }
        public bool PreserveSourceBackground { get; set; }
        public bool PreserveSourceFonts { get; set; }
        public bool UsePresentationTheme { get; set; }
    }

    public sealed class HtmlImportPreview
    {
        public HtmlImportPreview()
        {
            Slides = new List<HtmlImportSlidePreview>();
        }

        public List<HtmlImportSlidePreview> Slides { get; private set; }
        public HtmlImportReport Report { get; internal set; }
    }

    public sealed class HtmlImportReport
    {
        private readonly List<HtmlImportReportItem> _warnings = new List<HtmlImportReportItem>();

        public IReadOnlyList<HtmlImportReportItem> Warnings { get { return _warnings; } }
        public bool HasWarnings { get { return _warnings.Count > 0; } }

        internal void AddWarning(string message, int? slideNumber)
        {
            if (string.IsNullOrWhiteSpace(message) || _warnings.Any(item => item.Message == message && item.SlideNumber == slideNumber)) return;
            _warnings.Add(new HtmlImportReportItem { Message = message, SlideNumber = slideNumber });
        }

        public string ToDisplayText(int maximumItems)
        {
            if (!HasWarnings) return "No import warnings.";
            int take = Math.Max(1, maximumItems);
            var lines = _warnings.Take(take).Select(item => (item.SlideNumber.HasValue ? "Slide " + item.SlideNumber.Value + ": " : string.Empty) + item.Message).ToList();
            if (_warnings.Count > take) lines.Add("… and " + (_warnings.Count - take) + " more warning(s).");
            return string.Join(Environment.NewLine, lines);
        }
    }

    public sealed class HtmlImportReportItem
    {
        public int? SlideNumber { get; internal set; }
        public string Message { get; internal set; }
    }

    /// <summary>
    /// Developer-facing visual regression helper. Baseline PNGs can be approved once,
    /// then compared against freshly exported imported slides without a test framework.
    /// </summary>
    public static class HtmlImportVisualRegression
    {
        public static void ExportSlidesToPng(Presentation presentation, string outputDirectory, int width, int height)
        {
            if (presentation == null) throw new ArgumentNullException("presentation");
            if (string.IsNullOrWhiteSpace(outputDirectory)) throw new ArgumentException("Output directory is required.", "outputDirectory");
            Directory.CreateDirectory(outputDirectory);
            for (int index = 1; index <= presentation.Slides.Count; index++)
            {
                string path = Path.Combine(outputDirectory, "slide-" + index.ToString("D3") + ".png");
                presentation.Slides[index].Export(path, "PNG", width, height);
            }
        }

        public static HtmlVisualRegressionResult ComparePngDirectories(string baselineDirectory, string actualDirectory, byte perChannelTolerance, float maximumDifferentPixelRatio)
        {
            var result = new HtmlVisualRegressionResult();
            if (!Directory.Exists(baselineDirectory) || !Directory.Exists(actualDirectory))
            {
                result.Messages.Add("Baseline or actual PNG directory does not exist.");
                return result;
            }

            string[] baselineFiles = Directory.GetFiles(baselineDirectory, "*.png").OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToArray();
            string[] actualFiles = Directory.GetFiles(actualDirectory, "*.png").OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToArray();
            if (baselineFiles.Length != actualFiles.Length)
            {
                result.Messages.Add("Slide image count differs: expected " + baselineFiles.Length + ", got " + actualFiles.Length + ".");
                return result;
            }

            for (int index = 0; index < baselineFiles.Length; index++)
            {
                float difference = ComparePng(baselineFiles[index], actualFiles[index], perChannelTolerance);
                result.DifferentPixelRatios.Add(difference);
                if (difference > maximumDifferentPixelRatio)
                {
                    result.Messages.Add(Path.GetFileName(baselineFiles[index]) + " differs by " + difference.ToString("P2", System.Globalization.CultureInfo.InvariantCulture) + ".");
                }
            }
            return result;
        }

        private static float ComparePng(string baselinePath, string actualPath, byte tolerance)
        {
            using (var baseline = new Bitmap(baselinePath))
            using (var actual = new Bitmap(actualPath))
            {
                if (baseline.Width != actual.Width || baseline.Height != actual.Height) return 1f;
                long different = 0;
                long total = (long)baseline.Width * baseline.Height;
                for (int y = 0; y < baseline.Height; y++)
                {
                    for (int x = 0; x < baseline.Width; x++)
                    {
                        Color expected = baseline.GetPixel(x, y);
                        Color received = actual.GetPixel(x, y);
                        if (Math.Abs(expected.R - received.R) > tolerance || Math.Abs(expected.G - received.G) > tolerance || Math.Abs(expected.B - received.B) > tolerance) different++;
                    }
                }
                return total == 0 ? 0f : different / (float)total;
            }
        }
    }

    public sealed class HtmlVisualRegressionResult
    {
        public HtmlVisualRegressionResult()
        {
            Messages = new List<string>();
            DifferentPixelRatios = new List<float>();
        }

        public List<string> Messages { get; private set; }
        public List<float> DifferentPixelRatios { get; private set; }
        public bool Passed { get { return Messages.Count == 0; } }
    }

    public sealed class HtmlImportSlidePreview
    {
        public int Number { get; set; }
        public string Title { get; set; }
    }
}
