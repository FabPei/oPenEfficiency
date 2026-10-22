using System.Collections.Generic;
using System.Drawing;

namespace oPenEfficiency.Excel.Features.ChartScanner
{
    /// <summary>A number recognised in the image, with the centre of its bounding box (image pixels).</summary>
    public struct OcrNumber
    {
        public double Value;
        public double CenterX;
        public double CenterY;
    }

    /// <summary>
    /// Abstraction over optical character recognition of axis numbers, so the digitizer
    /// works with or without OCR. The default implementation returns nothing, which makes
    /// the whole feature fall back to manual value entry.
    /// </summary>
    public interface IChartOcr
    {
        bool IsAvailable { get; }

        /// <summary>Reads numeric labels from the given image region (or the whole image if region is empty).</summary>
        IReadOnlyList<OcrNumber> ReadNumbers(Bitmap image, Rectangle region);
    }

    /// <summary>
    /// No-op OCR: always returns an empty result. Used as the default so the Chart Scanner
    /// is fully functional via manual calibration even when no OCR engine is wired up.
    ///
    /// To enable real recognition with the built-in Windows OCR (Windows.Media.Ocr):
    ///   1. Add the NuGet package "Microsoft.Windows.SDK.Contracts" (and reference
    ///      System.Runtime.WindowsRuntime) so the WinRT APIs resolve under .NET Framework 4.8.
    ///   2. Implement IChartOcr using
    ///        var engine = Windows.Media.Ocr.OcrEngine.TryCreateFromUserProfileLanguages();
    ///      convert the Bitmap to a SoftwareBitmap, call engine.RecognizeAsync, and parse
    ///      each OcrWord.Text into a number (respecting the user's decimal separator).
    ///   3. Swap NullChartOcr for the new class in ChartScannerWindow.
    /// This is intentionally isolated so wiring OCR cannot break the rest of the build.
    /// </summary>
    public sealed class NullChartOcr : IChartOcr
    {
        public bool IsAvailable => false;

        public IReadOnlyList<OcrNumber> ReadNumbers(Bitmap image, Rectangle region)
            => System.Array.Empty<OcrNumber>();
    }
}
