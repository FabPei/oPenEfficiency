using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Windows.Data.Pdf;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage;
using Windows.Storage.Streams;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnLocalSlideRebuild",
        Name = "Rebuild Slide from PDF/Image",
        SearchAliases = new[] { "Rebuild Slide Locally" },
        Tooltip = "Rebuild a PDF page or image as a local PowerPoint slide",
        IconData = "M5,3H19A2,2 0 0,1 21,5V19A2,2 0 0,1 19,21H5A2,2 0 0,1 3,19V5A2,2 0 0,1 5,3M7,7V17H17V7H7M9,9H15V11H9V9M9,13H13V15H9V13Z",
        Color = "#0EA5E9",
        Description = "Uses Windows OCR and PDF rendering on this device to create a local reference slide with editable text, background, and accent shapes.",
        DetailedHelpText = "### Rebuild Slide from PDF/Image\n\nChoose an image or PDF. Right-click can restrict the picker to images or PDFs, or request selected PDF pages such as 1-3,5. The source remains on your computer: Windows renders PDF pages locally and Windows OCR detects text locally. A progress window shows the current rendering and OCR step.\n\nThe feature creates a native background color and long horizontal accent lines, then adds the original page as a reference background and overlays editable PowerPoint text boxes at the detected positions. It automatically masks detected source text when the surrounding background is a single color, so edits do not leave duplicate text behind. Complex graphics, diagrams, text formatting, photos, and gradients remain part of the reference image. For best results, use high-resolution source files and install the Windows language pack used by the source text.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class LocalSlideRebuildFeature
    {
        public enum SourceMode { AnyFile, Image, EntirePdf, PdfPageRange }

        private const string FeatureTag = "oPE_LocalSlideRebuild";
        private static readonly object RebuildLock = new object();
        private static bool _isRebuilding;

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, SourceMode.AnyFile);
        }

        public static bool Execute(PowerPointManager manager, SourceMode mode)
        {
            lock (RebuildLock)
            {
                if (_isRebuilding)
                {
                    MessageBox.Show("A local slide rebuild is already running.", "Rebuild Slide from PDF/Image", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                _isRebuilding = true;
            }

            try
            {
                var application = manager?.GetApplication();
                if (application?.ActivePresentation == null)
                {
                    MessageBox.Show("Please open a PowerPoint presentation before rebuilding a slide.", "Rebuild Slide from PDF/Image", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                using (var dialog = new OpenFileDialog())
                {
                    dialog.Title = "Choose a PDF page or image to rebuild locally";
                    dialog.Filter = mode == SourceMode.Image
                        ? "Image files (*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff)|*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff"
                        : mode == SourceMode.EntirePdf || mode == SourceMode.PdfPageRange
                            ? "PDF files (*.pdf)|*.pdf"
                            : "PDF and image files (*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff)|*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff|PDF files (*.pdf)|*.pdf|Image files (*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff)|*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff";
                    dialog.CheckFileExists = true;
                    dialog.Multiselect = false;

                    if (dialog.ShowDialog() != DialogResult.OK)
                    {
                        return false;
                    }

                    HashSet<uint> pageNumbers = null;
                    if (mode == SourceMode.PdfPageRange)
                    {
                        using (var rangeDialog = new PdfPageRangeDialog())
                        {
                            if (rangeDialog.ShowDialog(new PowerPointWindowOwner((IntPtr)application.HWND)) != DialogResult.OK)
                                return false;
                            if (!TryParsePageRange(rangeDialog.PageRange, out pageNumbers))
                            {
                                MessageBox.Show("Enter page numbers such as 1-3,5,8.", "Invalid PDF Page Range", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return false;
                            }
                        }
                    }

                    return RebuildOnPowerPointThread(application, dialog.FileName, pageNumbers);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LocalSlideRebuildFeature.Execute");
                MessageBox.Show(
                    "The file could not be rebuilt locally. Ensure that Windows has an OCR language pack matching the source text, then check the openEfficiency error log for details.",
                    "Rebuild Slide from PDF/Image",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                lock (RebuildLock)
                {
                    _isRebuilding = false;
                }
            }
        }

        private static bool RebuildOnPowerPointThread(PowerPoint.Application application, string sourcePath, HashSet<uint> pageNumbers)
        {
            var progress = new LocalRebuildProgressDialog();
            List<LocalPage> pages = null;
            try
            {
                progress.Show(new PowerPointWindowOwner((IntPtr)application.HWND));
                Action<string> reportProgress = status => ReportProgress(progress, status);
                Task<List<LocalPage>> loadTask = Task.Run(async () => await LoadPagesAsync(sourcePath, reportProgress, pageNumbers));
                while (!loadTask.IsCompleted)
                {
                    // Keep the visible progress form responsive without moving any PowerPoint
                    // COM object out of the original VSTO apartment.
                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(40);
                }

                pages = loadTask.GetAwaiter().GetResult();
                CompleteRebuild(application, pages, progress);
                return pages.Count > 0;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LocalSlideRebuildFeature.RebuildOnPowerPointThread");
                ShowRebuildError(progress);
                return false;
            }
            finally
            {
                if (!progress.IsDisposed)
                {
                    progress.Close();
                }
                CleanupTemporaryPages(pages);
            }
        }

        private static void CleanupTemporaryPages(IEnumerable<LocalPage> pages)
        {
            if (pages == null) return;
            foreach (var page in pages.Where(item => item.IsTemporaryRender))
            {
                try
                {
                    if (File.Exists(page.RenderPath)) File.Delete(page.RenderPath);
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "LocalSlideRebuildFeature.CleanupTemporaryPages");
                }
            }
        }

        private static void CompleteRebuild(PowerPoint.Application application, List<LocalPage> pages, LocalRebuildProgressDialog progress)
        {
            if (pages.Count == 0)
            {
                MessageBox.Show("No pages could be read from the selected file.", "Rebuild Slide from PDF/Image", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            progress.SetStatus("Creating editable PowerPoint shapes…");
            progress.Refresh();
            try { application.StartNewUndoEntry(); } catch { /* Rebuild remains valid when undo grouping is unavailable. */ }
            int created = AddSlides(application, pages);
            MessageBox.Show(
                string.Format("Created {0} reconstructed slide(s). The analysis ran locally using Windows PDF rendering and OCR.\n\nThe original page is kept as a reference background. Detected text is editable, and its source text is automatically masked on flat-color backgrounds.", created),
                "Rebuild Slide from PDF/Image",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static void ShowRebuildError(LocalRebuildProgressDialog progress)
        {
            progress.Close();
            MessageBox.Show(
                "The file could not be rebuilt locally. Check the openEfficiency error log for details.",
                "Rebuild Slide from PDF/Image",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        private static void ReportProgress(LocalRebuildProgressDialog progress, string status)
        {
            if (progress.IsDisposed || !progress.IsHandleCreated)
            {
                return;
            }

            try
            {
                progress.BeginInvoke(new Action(() => progress.SetStatus(status)));
            }
            catch (InvalidOperationException)
            {
                // PowerPoint is shutting down and the progress window is no longer available.
            }
        }

        private static async Task<List<LocalPage>> LoadPagesAsync(string sourcePath, Action<string> reportProgress, HashSet<uint> pageNumbers)
        {
            var extension = Path.GetExtension(sourcePath);
            if (string.Equals(extension, ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                return await LoadPdfPagesAsync(sourcePath, reportProgress, pageNumbers);
            }

            reportProgress?.Invoke("Reading image and running local OCR…");
            return new List<LocalPage> { await LoadImagePageAsync(sourcePath, reportProgress) };
        }

        private static async Task<LocalPage> LoadImagePageAsync(string sourcePath, Action<string> reportProgress)
        {
            StorageFile sourceFile = await StorageFile.GetFileFromPathAsync(sourcePath);
            using (IRandomAccessStream stream = await sourceFile.OpenAsync(FileAccessMode.Read))
            {
                BitmapDecoder decoder = await BitmapDecoder.CreateAsync(stream);
                using (SoftwareBitmap bitmap = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied))
                {
                    var text = await RecognizeTextAsync(bitmap);
                    reportProgress?.Invoke("Analyzing background and accent lines…");
                    return CreateLocalPage(sourcePath, bitmap.PixelWidth, bitmap.PixelHeight, text, isTemporaryRender: false);
                }
            }
        }

        private static async Task<List<LocalPage>> LoadPdfPagesAsync(string sourcePath, Action<string> reportProgress, HashSet<uint> pageNumbers)
        {
            var pages = new List<LocalPage>();
            string pendingRenderPath = null;
            try
            {
                StorageFile sourceFile = await StorageFile.GetFileFromPathAsync(sourcePath);
                PdfDocument document = await PdfDocument.LoadFromFileAsync(sourceFile);
                if (pageNumbers != null && pageNumbers.Any(number => number == 0 || number > document.PageCount))
                    throw new ArgumentOutOfRangeException(nameof(pageNumbers), "The requested PDF page range exceeds the document page count.");
                uint selectedCount = pageNumbers == null ? document.PageCount : (uint)pageNumbers.Count;
                uint processedCount = 0;
                for (uint pageIndex = 0; pageIndex < document.PageCount; pageIndex++)
                {
                    uint pageNumber = pageIndex + 1;
                    if (pageNumbers != null && !pageNumbers.Contains(pageNumber)) continue;
                    processedCount++;
                    reportProgress?.Invoke(string.Format("Rendering and recognizing PDF page {0} ({1} of {2} selected)…", pageNumber, processedCount, selectedCount));
                    using (PdfPage page = document.GetPage(pageIndex))
                    using (var renderStream = new InMemoryRandomAccessStream())
                    {
                        await page.RenderToStreamAsync(renderStream);
                        renderStream.Seek(0);
                        BitmapDecoder decoder = await BitmapDecoder.CreateAsync(renderStream);
                        using (SoftwareBitmap bitmap = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied))
                        {
                            pendingRenderPath = await SaveRenderedPageAsync(bitmap, pageIndex + 1);
                            var text = await RecognizeTextAsync(bitmap);
                            reportProgress?.Invoke(string.Format("Analyzing background and accent lines on PDF page {0}…", pageNumber));
                            pages.Add(CreateLocalPage(pendingRenderPath, bitmap.PixelWidth, bitmap.PixelHeight, text, isTemporaryRender: true));
                            pendingRenderPath = null;
                        }
                    }
                }
            }
            catch
            {
                if (!string.IsNullOrWhiteSpace(pendingRenderPath)) TryDeleteTemporaryRender(pendingRenderPath);
                foreach (var page in pages) TryDeleteTemporaryRender(page.RenderPath);
                throw;
            }

            return pages;
        }

        private static bool TryParsePageRange(string value, out HashSet<uint> pageNumbers)
        {
            return oPenEfficiency.Core.Services.SlideRebuildPageRangePolicy.TryParsePageRange(value, out pageNumbers);
        }

        private static void TryDeleteTemporaryRender(string path)
        {
            try
            {
                if (File.Exists(path)) File.Delete(path);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LocalSlideRebuildFeature.TryDeleteTemporaryRender");
            }
        }

        private static LocalPage CreateLocalPage(string renderPath, int pixelWidth, int pixelHeight, List<RecognizedText> text, bool isTemporaryRender)
        {
            using (var sourceBitmap = new System.Drawing.Bitmap(renderPath))
            {
                HorizontalAccentLine accent;
                TryFindHorizontalAccentLine(sourceBitmap, out accent);
                return new LocalPage(renderPath, pixelWidth, pixelHeight, text, GetAverageCornerColor(sourceBitmap), accent, isTemporaryRender);
            }
        }

        private static async Task<string> SaveRenderedPageAsync(SoftwareBitmap bitmap, uint pageNumber)
        {
            string path = Path.Combine(
                Path.GetTempPath(),
                string.Format("oPE_Rebuild_{0}_{1}.png", Guid.NewGuid().ToString("N"), pageNumber));

            using (var stream = new InMemoryRandomAccessStream())
            {
                BitmapEncoder encoder = await BitmapEncoder.CreateAsync(BitmapEncoder.PngEncoderId, stream);
                encoder.SetSoftwareBitmap(bitmap);
                await encoder.FlushAsync();

                stream.Seek(0);
                using (var source = stream.AsStreamForRead())
                using (var destination = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await source.CopyToAsync(destination);
                }
            }

            return path;
        }

        private static async Task<List<RecognizedText>> RecognizeTextAsync(SoftwareBitmap bitmap)
        {
            OcrEngine engine = OcrEngine.TryCreateFromUserProfileLanguages();
            if (engine == null)
            {
                throw new InvalidOperationException("Windows OCR is unavailable for the current profile languages.");
            }

            OcrResult result = await engine.RecognizeAsync(bitmap);
            var text = new List<RecognizedText>();
            foreach (OcrLine line in result.Lines)
            {
                string lineText = string.Join(" ", line.Words.Select(word => word.Text)).Trim();
                if (string.IsNullOrWhiteSpace(lineText) || line.Words.Count == 0)
                {
                    continue;
                }

                var first = line.Words[0].BoundingRect;
                var last = line.Words[line.Words.Count - 1].BoundingRect;
                text.Add(new RecognizedText(lineText, first.X, first.Y, Math.Max(1, last.X + last.Width - first.X), Math.Max(1, Math.Max(first.Height, last.Height))));
            }

            return text;
        }

        private static int AddSlides(PowerPoint.Application application, IEnumerable<LocalPage> pages)
        {
            var presentation = application.ActivePresentation;
            float slideWidth = presentation.PageSetup.SlideWidth;
            float slideHeight = presentation.PageSetup.SlideHeight;
            var activeSlide = application.ActiveWindow?.View?.Slide as PowerPoint.Slide;
            int insertAt = activeSlide?.SlideIndex ?? presentation.Slides.Count;
            int created = 0;
            var createdSlides = new List<PowerPoint.Slide>();

            try
            {
                foreach (LocalPage page in pages)
                {
                    var slide = presentation.Slides.Add(insertAt + ++created, PowerPoint.PpSlideLayout.ppLayoutBlank);
                    createdSlides.Add(slide);

                    using (var sourceBitmap = new System.Drawing.Bitmap(page.RenderPath))
                    {
                        AddNativeBackground(slide, slideWidth, slideHeight, page.BackgroundColor);

                        var reference = slide.Shapes.AddPicture(page.RenderPath, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, slideWidth, slideHeight);
                        reference.Name = FeatureTag + "_Reference";
                        reference.AlternativeText = "Local source reference for editable OCR text overlays and complex visual details.";
                        reference.Tags.Add(FeatureTag, "ReferenceBackground");

                        AddNativeAccentLines(slide, page.AccentLine, page.PixelWidth, page.PixelHeight, slideWidth, slideHeight);

                        foreach (RecognizedText text in page.Text)
                        {
                            float left = (float)(text.X / page.PixelWidth * slideWidth);
                            float top = (float)(text.Y / page.PixelHeight * slideHeight);
                            float width = (float)(text.Width / page.PixelWidth * slideWidth);
                            float height = Math.Max(8f, (float)(text.Height / page.PixelHeight * slideHeight * 1.35));
                            System.Drawing.Color backgroundColor;
                            bool canMaskText = TryGetFlatBackground(sourceBitmap, text, out backgroundColor);

                            if (canMaskText)
                            {
                                const float maskPadding = 1.5f;
                                var mask = slide.Shapes.AddShape(
                                    MsoAutoShapeType.msoShapeRectangle,
                                    Math.Max(0, left - maskPadding),
                                    Math.Max(0, top - maskPadding),
                                    Math.Min(slideWidth - left + maskPadding, width + maskPadding * 2),
                                    Math.Min(slideHeight - top + maskPadding, height + maskPadding * 2));
                                mask.Name = FeatureTag + "_Mask_" + slide.Shapes.Count;
                                mask.Tags.Add(FeatureTag, "OcrTextMask");
                                mask.Line.Visible = MsoTriState.msoFalse;
                                mask.Fill.ForeColor.RGB = ColorToOle(backgroundColor);
                            }

                            var overlay = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, Math.Max(8f, width), height);
                            overlay.Name = FeatureTag + "_Text_" + slide.Shapes.Count;
                            overlay.Tags.Add(FeatureTag, "EditableOcrText");
                            overlay.Line.Visible = MsoTriState.msoFalse;
                            overlay.Fill.Visible = MsoTriState.msoFalse;
                            overlay.TextFrame.MarginLeft = 0;
                            overlay.TextFrame.MarginRight = 0;
                            overlay.TextFrame.MarginTop = 0;
                            overlay.TextFrame.MarginBottom = 0;
                            overlay.TextFrame.WordWrap = MsoTriState.msoFalse;
                            overlay.TextFrame.TextRange.Text = text.Value;
                            overlay.TextFrame.TextRange.Font.Name = "Arial";
                            overlay.TextFrame.TextRange.Font.Size = Math.Max(6f, height * 0.72f);
                            overlay.TextFrame.TextRange.Font.Color.RGB = ColorToOle(canMaskText && GetBrightness(backgroundColor) < 128 ? System.Drawing.Color.White : System.Drawing.Color.FromArgb(28, 28, 28));
                        }
                    }
                }
            }
            catch
            {
                for (int index = createdSlides.Count - 1; index >= 0; index--)
                {
                    try { createdSlides[index].Delete(); } catch { /* Rollback slide may already have been removed after a failed rebuild. */ }
                }
                throw;
            }

            return created;
        }

        private static int ColorToOle(System.Drawing.Color color)
        {
            return color.R | (color.G << 8) | (color.B << 16);
        }

        private static void AddNativeBackground(PowerPoint.Slide slide, float slideWidth, float slideHeight, System.Drawing.Color color)
        {
            var background = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, slideWidth, slideHeight);
            background.Name = FeatureTag + "_NativeBackground";
            background.Tags.Add(FeatureTag, "NativeBackground");
            background.Line.Visible = MsoTriState.msoFalse;
            background.Fill.ForeColor.RGB = ColorToOle(color);
        }

        private static void AddNativeAccentLines(PowerPoint.Slide slide, HorizontalAccentLine accent, int pixelWidth, int pixelHeight, float slideWidth, float slideHeight)
        {
            if (accent == null)
            {
                return;
            }

            const int segmentCount = 3;
            float segmentWidth = (float)accent.Width / segmentCount / pixelWidth * slideWidth;
            float top = (float)accent.Top / pixelHeight * slideHeight;
            float height = Math.Max(1.5f, (float)accent.Height / pixelHeight * slideHeight);
            for (int segment = 0; segment < segmentCount; segment++)
            {
                float left = (float)(accent.Left + accent.Width * segment / (double)segmentCount) / pixelWidth * slideWidth;
                var lineSegment = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top, segmentWidth + 0.5f, height);
                lineSegment.Name = FeatureTag + "_AccentLine_" + (segment + 1);
                lineSegment.Tags.Add(FeatureTag, "NativeAccentLine");
                lineSegment.Line.Visible = MsoTriState.msoFalse;
                lineSegment.Fill.ForeColor.RGB = ColorToOle(accent.Colors[segment]);
            }
        }

        private static System.Drawing.Color GetAverageCornerColor(System.Drawing.Bitmap bitmap)
        {
            int insetX = Math.Min(12, Math.Max(0, bitmap.Width / 10));
            int insetY = Math.Min(12, Math.Max(0, bitmap.Height / 10));
            var corners = new[]
            {
                bitmap.GetPixel(insetX, insetY),
                bitmap.GetPixel(bitmap.Width - 1 - insetX, insetY),
                bitmap.GetPixel(insetX, bitmap.Height - 1 - insetY),
                bitmap.GetPixel(bitmap.Width - 1 - insetX, bitmap.Height - 1 - insetY)
            };
            return System.Drawing.Color.FromArgb(
                corners.Sum(color => color.R) / corners.Length,
                corners.Sum(color => color.G) / corners.Length,
                corners.Sum(color => color.B) / corners.Length);
        }

        private static bool TryFindHorizontalAccentLine(System.Drawing.Bitmap bitmap, out HorizontalAccentLine accent)
        {
            accent = null;
            int minimumWidth = (int)(bitmap.Width * 0.30);
            int bestWidth = 0;
            int bestLeft = 0;
            int bestTop = 0;

            for (int y = 0; y < bitmap.Height; y++)
            {
                int runStart = -1;
                for (int x = 0; x < bitmap.Width; x++)
                {
                    if (IsAccentColor(bitmap.GetPixel(x, y)))
                    {
                        if (runStart < 0) runStart = x;
                    }
                    else if (runStart >= 0)
                    {
                        int width = x - runStart;
                        if (width >= minimumWidth && width > bestWidth)
                        {
                            bestWidth = width;
                            bestLeft = runStart;
                            bestTop = y;
                        }
                        runStart = -1;
                    }
                }
            }

            if (bestWidth == 0)
            {
                return false;
            }

            int bottom = bestTop;
            while (bottom + 1 < bitmap.Height && IsAccentColor(bitmap.GetPixel(bestLeft + bestWidth / 2, bottom + 1))) bottom++;
            int top = bestTop;
            while (top > 0 && IsAccentColor(bitmap.GetPixel(bestLeft + bestWidth / 2, top - 1))) top--;

            var colors = new System.Drawing.Color[3];
            for (int index = 0; index < colors.Length; index++)
            {
                int sampleX = bestLeft + (int)(bestWidth * ((index + 0.5) / colors.Length));
                colors[index] = bitmap.GetPixel(Math.Min(bitmap.Width - 1, sampleX), bestTop);
            }
            accent = new HorizontalAccentLine(bestLeft, top, bestWidth, bottom - top + 1, colors);
            return true;
        }

        private static bool IsAccentColor(System.Drawing.Color color)
        {
            int maximum = Math.Max(color.R, Math.Max(color.G, color.B));
            int minimum = Math.Min(color.R, Math.Min(color.G, color.B));
            return maximum - minimum >= 35 && GetBrightness(color) >= 70;
        }

        private static bool TryGetFlatBackground(System.Drawing.Bitmap bitmap, RecognizedText text, out System.Drawing.Color backgroundColor)
        {
            backgroundColor = System.Drawing.Color.Empty;
            int left = Math.Max(0, (int)Math.Floor(text.X) - 3);
            int top = Math.Max(0, (int)Math.Floor(text.Y) - 3);
            int right = Math.Min(bitmap.Width - 1, (int)Math.Ceiling(text.X + text.Width) + 3);
            int bottom = Math.Min(bitmap.Height - 1, (int)Math.Ceiling(text.Y + text.Height) + 3);
            if (right <= left || bottom <= top)
            {
                return false;
            }

            var samples = new List<System.Drawing.Color>();
            for (int x = left; x <= right; x += Math.Max(1, (right - left) / 8))
            {
                samples.Add(bitmap.GetPixel(x, top));
                samples.Add(bitmap.GetPixel(x, bottom));
            }
            for (int y = top; y <= bottom; y += Math.Max(1, (bottom - top) / 4))
            {
                samples.Add(bitmap.GetPixel(left, y));
                samples.Add(bitmap.GetPixel(right, y));
            }

            int redRange = samples.Max(color => color.R) - samples.Min(color => color.R);
            int greenRange = samples.Max(color => color.G) - samples.Min(color => color.G);
            int blueRange = samples.Max(color => color.B) - samples.Min(color => color.B);
            if (redRange > 20 || greenRange > 20 || blueRange > 20)
            {
                return false;
            }

            backgroundColor = System.Drawing.Color.FromArgb(
                samples.Sum(color => color.R) / samples.Count,
                samples.Sum(color => color.G) / samples.Count,
                samples.Sum(color => color.B) / samples.Count);
            return true;
        }

        private static int GetBrightness(System.Drawing.Color color)
        {
            return (color.R * 299 + color.G * 587 + color.B * 114) / 1000;
        }

        private sealed class LocalPage
        {
            public LocalPage(string renderPath, int pixelWidth, int pixelHeight, List<RecognizedText> text, System.Drawing.Color backgroundColor, HorizontalAccentLine accentLine, bool isTemporaryRender)
            {
                RenderPath = renderPath;
                PixelWidth = Math.Max(1, pixelWidth);
                PixelHeight = Math.Max(1, pixelHeight);
                Text = text ?? new List<RecognizedText>();
                BackgroundColor = backgroundColor;
                AccentLine = accentLine;
                IsTemporaryRender = isTemporaryRender;
            }

            public string RenderPath { get; }
            public int PixelWidth { get; }
            public int PixelHeight { get; }
            public List<RecognizedText> Text { get; }
            public System.Drawing.Color BackgroundColor { get; }
            public HorizontalAccentLine AccentLine { get; }
            public bool IsTemporaryRender { get; }
        }

        private sealed class RecognizedText
        {
            public RecognizedText(string value, double x, double y, double width, double height)
            {
                Value = value;
                X = x;
                Y = y;
                Width = width;
                Height = height;
            }

            public string Value { get; }
            public double X { get; }
            public double Y { get; }
            public double Width { get; }
            public double Height { get; }
        }

        private sealed class HorizontalAccentLine
        {
            public HorizontalAccentLine(int left, int top, int width, int height, System.Drawing.Color[] colors)
            {
                Left = left;
                Top = top;
                Width = width;
                Height = height;
                Colors = colors;
            }

            public int Left { get; }
            public int Top { get; }
            public int Width { get; }
            public int Height { get; }
            public System.Drawing.Color[] Colors { get; }
        }

        private sealed class PdfPageRangeDialog : Form
        {
            private readonly TextBox _rangeBox;

            public string PageRange => _rangeBox.Text;

            public PdfPageRangeDialog()
            {
                Text = "PDF Page Range";
                Width = 360;
                Height = 165;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                StartPosition = FormStartPosition.CenterParent;
                MinimizeBox = false;
                MaximizeBox = false;

                Controls.Add(new Label
                {
                    Left = 14,
                    Top = 14,
                    Width = 315,
                    Text = "Pages to rebuild (for example: 1-3,5,8)"
                });
                _rangeBox = new TextBox { Left = 14, Top = 40, Width = 315, Text = "1" };
                Controls.Add(_rangeBox);

                var ok = new Button { Text = "OK", Left = 174, Top = 78, Width = 75, DialogResult = DialogResult.OK };
                var cancel = new Button { Text = "Cancel", Left = 254, Top = 78, Width = 75, DialogResult = DialogResult.Cancel };
                Controls.Add(ok);
                Controls.Add(cancel);
                AcceptButton = ok;
                CancelButton = cancel;
            }
        }

        private sealed class LocalRebuildProgressDialog : Form
        {
            private readonly Label _statusLabel;

            public LocalRebuildProgressDialog()
            {
                Text = "Rebuild Slide from PDF/Image";
                FormBorderStyle = FormBorderStyle.FixedDialog;
                StartPosition = FormStartPosition.CenterParent;
                ControlBox = false;
                ShowInTaskbar = false;
                Width = 460;
                Height = 135;

                _statusLabel = new Label
                {
                    AutoSize = false,
                    Left = 20,
                    Top = 18,
                    Width = 400,
                    Height = 34,
                    Text = "Preparing local reconstruction…"
                };
                var progressBar = new ProgressBar
                {
                    Left = 20,
                    Top = 62,
                    Width = 400,
                    Height = 18,
                    Style = ProgressBarStyle.Marquee,
                    MarqueeAnimationSpeed = 30
                };
                Controls.Add(_statusLabel);
                Controls.Add(progressBar);
            }

            public void SetStatus(string status)
            {
                _statusLabel.Text = status;
            }
        }

        private sealed class PowerPointWindowOwner : IWin32Window
        {
            public PowerPointWindowOwner(IntPtr handle) { Handle = handle; }
            public IntPtr Handle { get; }
        }
    }
}
