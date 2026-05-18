using System;
using System.IO;
using System.Net;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert a QR code image from a URL.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnQRCode",
        Name = "QR Code",
        Tooltip = "QR Code",
        IconData = "M3,3H7V7H3V3M3,17H7V21H3V17M17,3H21V7H17V3M17,17H21V21H17V17M5,5V5.01H6.99V5H5M5,19V19.01H6.99V19H5M19,5V5.01H20.99V5H19M19,19V19.01H20.99V19H19M9,9H5V13H9V9M13,9H17V13H13V9M9,17H13V15H15V17H19V13H15V15H13V9H15V7H9V9H11V13H9V17Z",
        Color = "#22C55E",
        Description = "Generates a QR code from a URL or text. Opens a dialog to enter the QR code content.",
        DetailedHelpText = "### QR Code\nGenerates a QR code from a URL or text string and inserts it as a native PowerPoint image shape directly onto the current slide.",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class QRCodeFeature
    {
        /// <summary>
        /// Opens the QR Code window for user input.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            new QRCodeWindow(manager).Show();
            return true;
        }

        public static bool Execute(PowerPointManager manager, string text)
        {
            if (manager == null || string.IsNullOrWhiteSpace(text)) return false;

            // Simple QR code API URL (e.g., using goqr.me)
            string encodedText = Uri.EscapeDataString(text);
            string qrUrl = $"https://api.qrserver.com/v1/create-qr-code/?size=150x150&data={encodedText}";

            return InsertImageFromUrl(manager, qrUrl);
        }

        public static bool InsertImageFromUrl(PowerPointManager manager, string imageUrl)
        {
            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                
                // Ensure TLS 1.2
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                string tempFile = Path.Combine(Path.GetTempPath(), "qr_" + Guid.NewGuid().ToString().Substring(0, 8) + ".png");

                using (var client = new WebClient())
                {
                    client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
                    client.DownloadFile(imageUrl, tempFile);
                }

                if (File.Exists(tempFile))
                {
                    float size = 150f;
                    float left = (app.ActivePresentation.PageSetup.SlideWidth / 2) - (size / 2);
                    float top = (app.ActivePresentation.PageSetup.SlideHeight / 2) - (size / 2);

                    var shape = slide.Shapes.AddPicture(tempFile, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, left, top, size, size);
                    shape.Select();

                    try { File.Delete(tempFile); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"QRCodeFeature.Cleanup tempFile error: {ex.Message}"); }
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("QR Generation Error: " + ex.Message);
                return false;
            }
        }
    }
}
