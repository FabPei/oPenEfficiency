using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

namespace oPenEfficiency.Excel.Features.ChartScanner
{
    /// <summary>
    /// OCR backed by the built-in Windows engine (Windows.Media.Ocr), used offline via the
    /// WinRT projection under .NET Framework 4.8. Recognises numeric axis labels and reports
    /// each number's value plus the centre of its bounding box in image pixels.
    ///
    /// Everything is wrapped defensively: if the engine or a language pack is unavailable,
    /// <see cref="IsAvailable"/> is false and <see cref="ReadNumbers"/> returns nothing, so
    /// the digitizer simply falls back to manual value entry.
    /// </summary>
    public sealed class WindowsChartOcr : IChartOcr
    {
        private readonly OcrEngine _engine;

        // Matches integer/decimal tokens, optionally with thousands/decimal separators and a sign.
        private static readonly Regex NumberToken = new Regex(@"-?\d[\d.,\s]*\d|-?\d", RegexOptions.Compiled);

        public WindowsChartOcr()
        {
            try { _engine = OcrEngine.TryCreateFromUserProfileLanguages(); }
            catch { _engine = null; }
        }

        public bool IsAvailable => _engine != null;

        public IReadOnlyList<OcrNumber> ReadNumbers(Bitmap image, Rectangle region)
        {
            var results = new List<OcrNumber>();
            if (_engine == null || image == null) return results;

            try
            {
                // Optionally work on a sub-region; remember the offset to map centres back.
                Bitmap working = image;
                int offsetX = 0, offsetY = 0;
                bool cropped = false;
                if (region.Width > 0 && region.Height > 0)
                {
                    var clip = Rectangle.Intersect(region, new Rectangle(0, 0, image.Width, image.Height));
                    if (clip.Width <= 0 || clip.Height <= 0) return results;
                    working = image.Clone(clip, image.PixelFormat);
                    offsetX = clip.X; offsetY = clip.Y; cropped = true;
                }

                // The engine rejects images beyond MaxImageDimension; downscale if needed.
                double scale = 1.0;
                int maxDim = (int)OcrEngine.MaxImageDimension;
                int longest = Math.Max(working.Width, working.Height);
                Bitmap scaled = working;
                bool scaledCreated = false;
                if (longest > maxDim)
                {
                    scale = (double)maxDim / longest;
                    scaled = new Bitmap((int)(working.Width * scale), (int)(working.Height * scale));
                    using (var g = Graphics.FromImage(scaled))
                    {
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.DrawImage(working, 0, 0, scaled.Width, scaled.Height);
                    }
                    scaledCreated = true;
                }

                OcrResult ocr = Recognize(scaled);
                if (ocr != null)
                {
                    foreach (var line in ocr.Lines)
                    {
                        foreach (var word in line.Words)
                        {
                            if (!TryParseLabelNumber(word.Text, out double value)) continue;
                            var r = word.BoundingRect;
                            double cx = (r.X + r.Width / 2.0) / scale + offsetX;
                            double cy = (r.Y + r.Height / 2.0) / scale + offsetY;
                            results.Add(new OcrNumber { Value = value, CenterX = cx, CenterY = cy });
                        }
                    }
                }

                if (scaledCreated) scaled.Dispose();
                if (cropped) working.Dispose();
            }
            catch
            {
                // Best effort: any WinRT/interop failure degrades to manual entry.
            }
            return results;
        }

        private OcrResult Recognize(Bitmap bmp)
        {
            using (var ms = new MemoryStream())
            {
                bmp.Save(ms, ImageFormat.Png);
                ms.Position = 0;
                using (IRandomAccessStream ras = ms.AsRandomAccessStream())
                {
                    var decoder = BitmapDecoder.CreateAsync(ras).AsTask().GetAwaiter().GetResult();
                    using (var software = decoder.GetSoftwareBitmapAsync().AsTask().GetAwaiter().GetResult())
                    using (var bgra = SoftwareBitmap.Convert(software, BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied))
                    {
                        return _engine.RecognizeAsync(bgra).AsTask().GetAwaiter().GetResult();
                    }
                }
            }
        }

        /// <summary>
        /// Parses a recognised token into a number, tolerating thousands separators, decimal
        /// commas/points and a trailing unit (e.g. "1.234,5", "40M", "75%"). Heuristic but only
        /// used to pre-fill the manual calibration prompt, which the user can still correct.
        /// </summary>
        internal static bool TryParseLabelNumber(string text, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;

            var match = NumberToken.Match(text);
            if (!match.Success) return false;
            string token = match.Value.Replace(" ", "");

            bool hasDot = token.Contains(".");
            bool hasComma = token.Contains(",");
            if (hasDot && hasComma)
            {
                // The rightmost separator is the decimal one; the other groups thousands.
                char dec = token.LastIndexOf('.') > token.LastIndexOf(',') ? '.' : ',';
                char thou = dec == '.' ? ',' : '.';
                token = token.Replace(thou.ToString(), "").Replace(dec, '.');
            }
            else if (hasComma)
            {
                token = NormalizeSingleSeparator(token, ',');
            }
            else if (hasDot)
            {
                token = NormalizeSingleSeparator(token, '.');
            }

            return double.TryParse(token, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }

        // One separator present: exactly 3 trailing digits => thousands grouping (remove),
        // otherwise treat it as the decimal separator.
        private static string NormalizeSingleSeparator(string token, char sep)
        {
            int idx = token.LastIndexOf(sep);
            string after = token.Substring(idx + 1);
            if (after.Length == 3 && CountChar(token, sep) == 1 && idx > 0)
                return token.Replace(sep.ToString(), "");
            return token.Replace(sep, '.');
        }

        private static int CountChar(string s, char c)
        {
            int n = 0;
            foreach (var ch in s) if (ch == c) n++;
            return n;
        }
    }
}
