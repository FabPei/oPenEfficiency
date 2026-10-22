using System;
using oPenEfficiency.Utils;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Exports PowerPoint slides as an editable HTML deck. Text, tables and common
    /// shapes become DOM/CSS; unsupported COM objects are exported as PNG fallbacks.
    /// </summary>
    public sealed class HtmlPresentationExportService
    {
        private const float CanvasWidth = 900f;

        public HtmlExportReport Export(Presentation presentation, IEnumerable<int> slideIndices, string outputHtmlPath)
        {
            if (presentation == null) throw new ArgumentNullException("presentation");
            if (string.IsNullOrWhiteSpace(outputHtmlPath)) throw new ArgumentException("Output path is required.", "outputHtmlPath");

            string fullOutputPath = Path.GetFullPath(outputHtmlPath);
            string outputDirectory = Path.GetDirectoryName(fullOutputPath);
            Directory.CreateDirectory(outputDirectory);
            string assetDirectoryName = Path.GetFileNameWithoutExtension(fullOutputPath) + "_assets";
            string assetDirectory = Path.Combine(outputDirectory, assetDirectoryName);
            Directory.CreateDirectory(assetDirectory);
            foreach (string staleAsset in Directory.GetFiles(assetDirectory, "slide-*-shape-*.png", SearchOption.TopDirectoryOnly))
            {
                File.Delete(staleAsset);
            }

            var report = new HtmlExportReport(fullOutputPath, assetDirectory);
            List<int> selectedIndices = (slideIndices ?? Enumerable.Range(1, presentation.Slides.Count)).Distinct().Where(index => index >= 1 && index <= presentation.Slides.Count).ToList();
            if (selectedIndices.Count == 0) throw new InvalidOperationException("No slides are selected for HTML export.");

            float scaleX = CanvasWidth / presentation.PageSetup.SlideWidth;
            float scaleY = scaleX;
            float canvasHeight = presentation.PageSetup.SlideHeight * scaleY;
            var css = new StringBuilder();
            var body = new StringBuilder();
            css.AppendLine(GetBaseCss(canvasHeight));

            int exportNumber = 0;
            foreach (int slideIndex in selectedIndices)
            {
                Slide slide = presentation.Slides[slideIndex];
                exportNumber++;
                body.AppendLine("<section class=\"slide\" data-slide=\"" + exportNumber + "\" id=\"slide-" + exportNumber + "\" style=\"" + GetSlideBackgroundCss(slide, report) + "\">");
                for (int shapeIndex = 1; shapeIndex <= slide.Shapes.Count; shapeIndex++)
                {
                    Shape shape = slide.Shapes[shapeIndex];
                    body.AppendLine(RenderShape(slide, shape, shapeIndex, exportNumber, scaleX, scaleY, assetDirectory, assetDirectoryName, report));
                }
                body.AppendLine("</section>");
            }

            string title = GetPresentationTitle(presentation);
            string html = "<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n<title>" + Encode(title) + "</title>\n<style>\n" + css + "</style>\n</head>\n<body>\n<div class=\"deck-toolbar\"><button type=\"button\" id=\"previous\">Previous</button><span id=\"counter\"></span><button type=\"button\" id=\"next\">Next</button><button type=\"button\" id=\"edit\">Edit text</button></div>\n<main class=\"deck\">\n" + body + "</main>\n<script>" + GetDeckScript() + "</script>\n</body>\n</html>";
            File.WriteAllText(fullOutputPath, html, new UTF8Encoding(false));
            return report;
        }

        private string RenderShape(Slide slide, Shape shape, int shapeIndex, int slideNumber, float scaleX, float scaleY, string assetDirectory, string assetDirectoryName, HtmlExportReport report)
        {
            try
            {
                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    return RenderTable(shape, shapeIndex, scaleX, scaleY);
                }

                if (shape.Type == Office.MsoShapeType.msoPicture || shape.Type == Office.MsoShapeType.msoLinkedPicture)
                {
                    return RenderPicture(shape, shapeIndex, slideNumber, scaleX, scaleY, assetDirectory, assetDirectoryName, report);
                }

                if (shape.Type == Office.MsoShapeType.msoAutoShape || shape.Type == Office.MsoShapeType.msoTextBox || shape.Type == Office.MsoShapeType.msoPlaceholder || shape.Type == Office.MsoShapeType.msoLine)
                {
                    return RenderEditableShape(shape, shapeIndex, scaleX, scaleY);
                }

                return RenderFallbackPicture(shape, shapeIndex, slideNumber, scaleX, scaleY, assetDirectory, assetDirectoryName, report);
            }
            catch (Exception ex)
            {
                report.AddWarning("Slide " + slideNumber + ", shape " + shapeIndex + " could not be exported as editable HTML: " + ex.Message);
                return RenderFallbackPicture(shape, shapeIndex, slideNumber, scaleX, scaleY, assetDirectory, assetDirectoryName, report);
            }
        }

        private string RenderEditableShape(Shape shape, int shapeIndex, float scaleX, float scaleY)
        {
            bool isLine = shape.Type == Office.MsoShapeType.msoLine;
            bool hasText = shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue;
            string style = GetPositionCss(shape, scaleX, scaleY) + GetShapeAppearanceCss(shape, scaleX, scaleY, isLine) + GetTextCss(shape, scaleX, scaleY);
            string link = GetShapeHyperlink(shape);
            string tag = hasText ? "div" : "div";
            string opening = "<" + tag + " class=\"ppt-shape" + (isLine ? " ppt-line" : string.Empty) + "\" data-ppt-shape=\"" + shapeIndex + "\" style=\"" + style + "\"";
            if (hasText) opening += " contenteditable=\"true\"";
            opening += ">";
            string content = hasText ? GetTextHtml(shape) : string.Empty;
            string element = opening + content + "</" + tag + ">";
            return string.IsNullOrWhiteSpace(link) ? element : "<a class=\"ppt-link\" href=\"" + EncodeAttribute(link) + "\">" + element + "</a>";
        }

        private string RenderPicture(Shape shape, int shapeIndex, int slideNumber, float scaleX, float scaleY, string assetDirectory, string assetDirectoryName, HtmlExportReport report)
        {
            string relativePath = ExportShapePng(shape, shapeIndex, slideNumber, assetDirectory, assetDirectoryName, report);
            if (string.IsNullOrWhiteSpace(relativePath)) return string.Empty;
            string style = GetPositionCss(shape, scaleX, scaleY) + "object-fit:contain;" + GetRotationCss(shape);
            string image = "<img class=\"ppt-image\" data-ppt-shape=\"" + shapeIndex + "\" src=\"" + EncodeAttribute(relativePath) + "\" alt=\"" + EncodeAttribute(shape.AlternativeText ?? string.Empty) + "\" style=\"" + style + "\">";
            string link = GetShapeHyperlink(shape);
            return string.IsNullOrWhiteSpace(link) ? image : "<a class=\"ppt-link\" href=\"" + EncodeAttribute(link) + "\">" + image + "</a>";
        }

        private string RenderFallbackPicture(Shape shape, int shapeIndex, int slideNumber, float scaleX, float scaleY, string assetDirectory, string assetDirectoryName, HtmlExportReport report)
        {
            string relativePath = ExportShapePng(shape, shapeIndex, slideNumber, assetDirectory, assetDirectoryName, report);
            if (string.IsNullOrWhiteSpace(relativePath)) return string.Empty;
            report.AddWarning("Slide " + slideNumber + ", shape " + shapeIndex + " (" + shape.Type + ") was exported as a PNG fallback.");
            return "<img class=\"ppt-fallback\" data-ppt-shape=\"" + shapeIndex + "\" src=\"" + EncodeAttribute(relativePath) + "\" alt=\"PowerPoint fallback\" style=\"" + GetPositionCss(shape, scaleX, scaleY) + GetRotationCss(shape) + "\">";
        }

        private string RenderTable(Shape shape, int shapeIndex, float scaleX, float scaleY)
        {
            Table table = shape.Table;
            var builder = new StringBuilder();
            builder.Append("<table class=\"ppt-table\" data-ppt-shape=\"").Append(shapeIndex).Append("\" style=\"").Append(GetPositionCss(shape, scaleX, scaleY)).Append(GetShapeAppearanceCss(shape, scaleX, scaleY, false)).Append("\">");
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                builder.Append("<tr>");
                for (int column = 1; column <= table.Columns.Count; column++)
                {
                    Cell cell = table.Cell(row, column);
                    string fill = GetFillCss(cell.Shape);
                    string textCss = GetTextCss(cell.Shape, scaleX, scaleY);
                    builder.Append("<td contenteditable=\"true\" style=\"").Append(fill).Append(textCss).Append("\">").Append(GetTextHtml(cell.Shape)).Append("</td>");
                }
                builder.Append("</tr>");
            }
            builder.Append("</table>");
            return builder.ToString();
        }

        private string ExportShapePng(Shape shape, int shapeIndex, int slideNumber, string assetDirectory, string assetDirectoryName, HtmlExportReport report)
        {
            try
            {
                string name = "slide-" + slideNumber.ToString("D3") + "-shape-" + shapeIndex.ToString("D3") + ".png";
                string path = Path.Combine(assetDirectory, name);
                shape.Export(path, PpShapeFormat.ppShapeFormatPNG, 0, 0, PpExportMode.ppRelativeToSlide);
                return assetDirectoryName + "/" + name;
            }
            catch (Exception ex)
            {
                report.AddWarning("Could not create PNG fallback for shape " + shapeIndex + ": " + ex.Message);
                return null;
            }
        }

        private static string GetPositionCss(Shape shape, float scaleX, float scaleY)
        {
            return "position:absolute;left:" + CssNumber(shape.Left * scaleX) + "px;top:" + CssNumber(shape.Top * scaleY) + "px;width:" + CssNumber(shape.Width * scaleX) + "px;height:" + CssNumber(shape.Height * scaleY) + "px;z-index:" + shape.ZOrderPosition + ";";
        }

        private static string GetShapeAppearanceCss(Shape shape, float scaleX, float scaleY, bool isLine)
        {
            if (isLine)
            {
                string lineColor = ToCssColor(shape.Line.ForeColor.RGB);
                return "height:0;border:0;border-top:" + CssNumber(Math.Max(1f, shape.Line.Weight * scaleY)) + "px solid " + lineColor + ";transform-origin:0 0;";
            }

            var builder = new StringBuilder();
            builder.Append(GetFillCss(shape));
            try
            {
                if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                {
                    builder.Append("border:").Append(CssNumber(Math.Max(0.25f, shape.Line.Weight * scaleX))).Append("px solid ").Append(ToCssColor(shape.Line.ForeColor.RGB)).Append(";");
                }
                else builder.Append("border:none;");
            }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "HtmlPresentationExportService.GetShapeCss.Fill"); }

            try
            {
                if (shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeOval) builder.Append("border-radius:50%;");
                else if (shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRoundedRectangle) builder.Append("border-radius:12px;");
            }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "HtmlPresentationExportService.GetShapeCss.Line"); }
            try
            {
                if (shape.Shadow.Visible == Office.MsoTriState.msoTrue)
                {
                    builder.Append("box-shadow:").Append(CssNumber(shape.Shadow.OffsetX * scaleX)).Append("px ").Append(CssNumber(shape.Shadow.OffsetY * scaleY)).Append("px ").Append(CssNumber(shape.Shadow.Blur * scaleX)).Append("px rgba(0,0,0,").Append(CssNumber(1f - shape.Shadow.Transparency / 100f)).Append(");");
                }
            }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "HtmlPresentationExportService.GetShapeCss.Shadow"); }
            builder.Append(GetRotationCss(shape));
            return builder.ToString();
        }

        private static string GetFillCss(Shape shape)
        {
            try
            {
                if (shape.Fill.Visible != Office.MsoTriState.msoTrue) return "background:transparent;";
                if (shape.Fill.Type == Office.MsoFillType.msoFillGradient && shape.Fill.GradientStops.Count >= 2)
                {
                    var stops = new List<string>();
                    for (int index = 1; index <= shape.Fill.GradientStops.Count; index++)
                    {
                        var stop = shape.Fill.GradientStops[index];
                        stops.Add(ToCssColor(stop.Color.RGB) + " " + CssNumber(stop.Position * 100f) + "%");
                    }
                    return "background:linear-gradient(135deg," + string.Join(",", stops) + ");";
                }
                return "background:" + ToCssColor(shape.Fill.ForeColor.RGB) + ";opacity:" + CssNumber(1f - shape.Fill.Transparency / 100f) + ";";
            }
            catch
            {
                return "background:transparent;";
            }
        }

        private static string GetTextCss(Shape shape, float scaleX, float scaleY)
        {
            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue) return string.Empty;
                TextRange range = shape.TextFrame.TextRange;
                string font = range.Font.Name;
                float size = range.Font.Size;
                string color = ToCssColor(range.Font.Color.RGB);
                string alignment = GetTextAlignment(range.ParagraphFormat.Alignment);
                string weight = range.Font.Bold == Office.MsoTriState.msoTrue ? "font-weight:700;" : string.Empty;
                string italic = range.Font.Italic == Office.MsoTriState.msoTrue ? "font-style:italic;" : string.Empty;
                string underline = range.Font.Underline == Office.MsoTriState.msoTrue ? "text-decoration:underline;" : string.Empty;
                return "box-sizing:border-box;padding:0;overflow:hidden;font-family:" + EncodeCssFont(font) + ";font-size:" + CssNumber(size * 4f / 3f * ((scaleX + scaleY) / 2f)) + "px;color:" + color + ";text-align:" + alignment + ";line-height:1.15;" + weight + italic + underline + "display:flex;flex-direction:column;justify-content:" + GetVerticalAlignment(shape) + ";";
            }
            catch
            {
                return "font-family:Segoe UI,sans-serif;font-size:16px;";
            }
        }

        private static string GetTextHtml(Shape shape)
        {
            try
            {
                string text = shape.TextFrame.TextRange.Text ?? string.Empty;
                return Encode(text).Replace("\r\n", "<br>").Replace("\n", "<br>").Replace("\r", "<br>");
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string GetSlideBackgroundCss(Slide slide, HtmlExportReport report)
        {
            try
            {
                if (slide.Background.Fill.Type == Office.MsoFillType.msoFillGradient && slide.Background.Fill.GradientStops.Count >= 2)
                {
                    var stops = new List<string>();
                    for (int index = 1; index <= slide.Background.Fill.GradientStops.Count; index++)
                    {
                        var stop = slide.Background.Fill.GradientStops[index];
                        stops.Add(ToCssColor(stop.Color.RGB) + " " + CssNumber(stop.Position * 100f) + "%");
                    }
                    return "background:linear-gradient(135deg," + string.Join(",", stops) + ");";
                }
                return "background:" + ToCssColor(slide.Background.Fill.ForeColor.RGB) + ";";
            }
            catch
            {
                report.AddWarning("Slide background could not be read; using white.");
                return "background:#ffffff;";
            }
        }

        private static string GetShapeHyperlink(Shape shape)
        {
            try
            {
                ActionSetting action = shape.ActionSettings[PpMouseActivation.ppMouseClick];
                if (action.Action != PpActionType.ppActionHyperlink) return null;
                if (!string.IsNullOrWhiteSpace(action.Hyperlink.Address)) return action.Hyperlink.Address;
                if (!string.IsNullOrWhiteSpace(action.Hyperlink.SubAddress)) return "#" + action.Hyperlink.SubAddress;
            }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "HtmlPresentationExportService.GetImageSource"); }
            return null;
        }

        private static string GetRotationCss(Shape shape)
        {
            try { return Math.Abs(shape.Rotation) > 0.01f ? "transform:rotate(" + CssNumber(shape.Rotation) + "deg);transform-origin:center center;" : string.Empty; } catch { return string.Empty; }
        }

        private static string GetVerticalAlignment(Shape shape)
        {
            try
            {
                if (shape.TextFrame.VerticalAnchor == Office.MsoVerticalAnchor.msoAnchorMiddle) return "center";
                if (shape.TextFrame.VerticalAnchor == Office.MsoVerticalAnchor.msoAnchorBottom) return "flex-end";
            }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "HtmlPresentationExportService.GetVerticalAlignment"); }
            return "flex-start";
        }

        private static string GetTextAlignment(PpParagraphAlignment alignment)
        {
            if (alignment == PpParagraphAlignment.ppAlignCenter) return "center";
            if (alignment == PpParagraphAlignment.ppAlignRight) return "right";
            if (alignment == PpParagraphAlignment.ppAlignJustify) return "justify";
            return "left";
        }

        private static string GetPresentationTitle(Presentation presentation)
        {
            try { return Path.GetFileNameWithoutExtension(presentation.Name); } catch { return "PowerPoint export"; }
        }

        private static string ToCssColor(int bgr)
        {
            int red = bgr & 0xff;
            int green = (bgr >> 8) & 0xff;
            int blue = (bgr >> 16) & 0xff;
            return string.Format(CultureInfo.InvariantCulture, "#{0:X2}{1:X2}{2:X2}", red, green, blue);
        }

        private static string CssNumber(float value) { return value.ToString("0.###", CultureInfo.InvariantCulture); }
        private static string Encode(string value) { return WebUtility.HtmlEncode(value ?? string.Empty); }
        private static string EncodeAttribute(string value) { return WebUtility.HtmlEncode(value ?? string.Empty).Replace("'", "&#39;"); }
        private static string EncodeCssFont(string value) { return "'" + (value ?? "Segoe UI").Replace("'", "") + "',sans-serif"; }

        private static string GetBaseCss(float canvasHeight)
        {
            string height = CssNumber(canvasHeight);
            return "*{box-sizing:border-box}body{margin:0;background:#111827;font-family:'Segoe UI',sans-serif;color:#1f2937}.deck{min-height:100vh;display:grid;place-items:center;padding:72px 24px 24px}.slide{position:relative;width:900px;height:" + height + "px;overflow:hidden;box-shadow:0 18px 48px rgba(0,0,0,.32);display:none}.slide.is-active{display:block}.ppt-shape,.ppt-image,.ppt-fallback,.ppt-table{position:absolute}.ppt-image,.ppt-fallback{display:block}.ppt-link{color:inherit;text-decoration:none}.ppt-table{border-collapse:collapse;table-layout:fixed}.ppt-table td{min-width:1px;vertical-align:middle}.ppt-shape:focus{outline:2px dashed #38bdf8;outline-offset:2px}.deck-toolbar{position:fixed;z-index:999;top:16px;left:50%;transform:translateX(-50%);display:flex;gap:8px;align-items:center;background:rgba(15,23,42,.92);padding:9px 12px;border-radius:999px;color:#fff}.deck-toolbar button{border:0;border-radius:999px;padding:7px 12px;background:#0ea5e9;color:#fff;cursor:pointer}.deck-toolbar button:hover{background:#0284c7}@media(max-width:960px){.deck{display:block;padding:64px 0 0}.slide{transform:scale(calc(100vw / 900));transform-origin:top left;margin-bottom:calc(" + height + "px * (100vw / 900) - " + height + "px)}}";
        }

        private static string GetDeckScript()
        {
            return "(()=>{const slides=[...document.querySelectorAll('.slide')];let index=0;const counter=document.querySelector('#counter');const show=next=>{index=(next+slides.length)%slides.length;slides.forEach((slide,i)=>slide.classList.toggle('is-active',i===index));counter.textContent=`${index+1} / ${slides.length}`};document.querySelector('#previous').onclick=()=>show(index-1);document.querySelector('#next').onclick=()=>show(index+1);document.querySelector('#edit').onclick=()=>document.body.classList.toggle('is-editing');document.addEventListener('keydown',event=>{if(event.target.isContentEditable)return;if(event.key==='ArrowRight')show(index+1);if(event.key==='ArrowLeft')show(index-1)});show(0)})();";
        }
    }

    public sealed class HtmlExportReport
    {
        internal HtmlExportReport(string htmlPath, string assetDirectory)
        {
            HtmlPath = htmlPath;
            AssetDirectory = assetDirectory;
            Warnings = new List<string>();
        }

        public string HtmlPath { get; private set; }
        public string AssetDirectory { get; private set; }
        public List<string> Warnings { get; private set; }
        internal void AddWarning(string warning) { if (!string.IsNullOrWhiteSpace(warning) && !Warnings.Contains(warning)) Warnings.Add(warning); }
    }
}
