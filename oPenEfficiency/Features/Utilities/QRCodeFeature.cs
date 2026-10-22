using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls.Primitives;
using System.Windows.Interop;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert a QR code image from a URL.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnQRCode",
        Name = "QR Code",
        Tooltip = "Generate and insert a QR code",
        IconData = "M3,3H7V7H3V3M3,17H7V21H3V17M17,3H21V7H17V3M17,17H21V21H17V17M5,5V5.01H6.99V5H5M5,19V19.01H6.99V19H5M19,5V5.01H20.99V5H19M19,19V19.01H20.99V19H19M9,9H5V13H9V9M13,9H17V13H13V9M9,17H13V15H15V17H19V13H15V15H13V9H15V7H9V9H11V13H9V17Z",
        Color = "#22C55E",
        Description = "Downloads a QR code for entered text and inserts it as a PowerPoint image on the current slide.",
        DetailedHelpText = "### QR Code\n\nEnter a URL or text string. openEfficiency sends that content to the external QR Server service, downloads the generated PNG, and inserts it at the center of the current slide. The temporary image is always removed after the attempt. Generation requires an internet connection and creates one PowerPoint Undo entry.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None,
        IsToggle = true)]
    public static class QRCodeFeature
    {
        private static QRCodeWindow _window;
        private static ToggleButton _toggleButton;

        /// <summary>
        /// Opens the QR Code window for user input.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, (ToggleButton)null);
        }

        public static bool Execute(PowerPointManager manager, ToggleButton toggleButton)
        {
            if (manager == null) return false;

            try
            {
                if (toggleButton != null && toggleButton.IsChecked == false)
                {
                    _window?.Close();
                    return true;
                }

                if (_window != null && _window.IsLoaded)
                {
                    if (toggleButton != null)
                    {
                        _toggleButton = toggleButton;
                        toggleButton.IsChecked = true;
                    }
                    _window.Activate();
                    return true;
                }

                _toggleButton = toggleButton;
                _window = new QRCodeWindow(manager);
                var app = manager.GetApplication();
                if (app != null) new WindowInteropHelper(_window).Owner = (IntPtr)app.HWND;
                _window.Closed += (sender, args) =>
                {
                    if (_toggleButton != null) _toggleButton.IsChecked = false;
                    _toggleButton = null;
                    _window = null;
                };
                if (toggleButton != null) toggleButton.IsChecked = true;
                _window.Show();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "QRCodeFeature.Execute");
                if (toggleButton != null) toggleButton.IsChecked = false;
                return false;
            }
        }

        public static Task<bool> GenerateAndInsertAsync(PowerPointManager manager, string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return Task.FromResult(false);
            string qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" +
                Uri.EscapeDataString(text);
            return InsertImageFromUrlAsync(manager, qrUrl);
        }

        public static async Task<bool> InsertImageFromUrlAsync(
            PowerPointManager manager,
            string imageUrl,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            string tempFile = null;
            try
            {
                if (manager == null || string.IsNullOrWhiteSpace(imageUrl)) return false;
                var app = manager.GetApplication();
                if (app?.ActiveWindow == null || app.ActivePresentation == null) return false;
                var sourcePresentation = app.ActivePresentation;
                
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                string tempDirectory = Path.Combine(Path.GetTempPath(), "oPenEfficiency");
                Directory.CreateDirectory(tempDirectory);
                tempFile = Path.Combine(tempDirectory, "qr_" + Guid.NewGuid().ToString("N") + ".png");

                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(30);
                    client.DefaultRequestHeaders.UserAgent.ParseAdd("oPenEfficiency/1.0");
                    using (var response = await client.GetAsync(imageUrl, cancellationToken))
                    {
                        response.EnsureSuccessStatusCode();
                        byte[] imageBytes = await response.Content.ReadAsByteArrayAsync();
                        cancellationToken.ThrowIfCancellationRequested();
                        if (imageBytes == null || imageBytes.Length == 0) return false;
                        File.WriteAllBytes(tempFile, imageBytes);
                    }
                }

                if (!File.Exists(tempFile)) return false;
                cancellationToken.ThrowIfCancellationRequested();
                if (app.ActivePresentation != sourcePresentation) return false;

                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "QRCodeFeature.StartNewUndoEntry"); }

                float size = 150f;
                float left = (app.ActivePresentation.PageSetup.SlideWidth / 2) - (size / 2);
                float top = (app.ActivePresentation.PageSetup.SlideHeight / 2) - (size / 2);
                var shape = slide.Shapes.AddPicture(
                    tempFile,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue,
                    left,
                    top,
                    size,
                    size);
                shape.Select();
                return true;
            }
            catch (OperationCanceledException)
            {
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "QRCodeFeature.InsertImageFromUrlAsync");
                return false;
            }
            finally
            {
                if (!string.IsNullOrEmpty(tempFile))
                {
                    try { if (File.Exists(tempFile)) File.Delete(tempFile); }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "QRCodeFeature.Cleanup"); }
                }
            }
        }
    }
}
