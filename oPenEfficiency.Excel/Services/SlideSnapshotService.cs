using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using oPenEfficiency.Utils;
using PP = global::Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Excel.Services
{
    /// <summary>
    /// Renders a single slide of a PowerPoint file to a PNG image by automating
    /// PowerPoint via COM. Used to insert and refresh linked slide snapshots in Excel.
    ///
    /// Safety: reuses a running PowerPoint instance if present (and never quits it);
    /// only quits a PowerPoint instance that this service started itself.
    /// The presentation is opened read-only and without a window.
    /// </summary>
    public static class SlideSnapshotService
    {
        /// <summary>
        /// Exports the given slide to a temporary PNG and returns its path,
        /// or null on failure (with a human-readable reason in <paramref name="error"/>).
        /// </summary>
        public static string ExportSlideToPng(string pptxPath, int slideIndex, out string error)
        {
            error = null;

            if (string.IsNullOrWhiteSpace(pptxPath) || !File.Exists(pptxPath))
            {
                error = "PowerPoint file not found.";
                return null;
            }

            PP.Application ppApp = null;
            bool createdApp = false;
            PP.Presentation pres = null;

            try
            {
                ppApp = GetOrCreatePowerPoint(out createdApp);
                if (ppApp == null)
                {
                    error = "PowerPoint konnte nicht gestartet werden.";
                    return null;
                }

                pres = ppApp.Presentations.Open(
                    pptxPath,
                    MsoTriState.msoTrue,   // ReadOnly
                    MsoTriState.msoFalse,  // Untitled
                    MsoTriState.msoFalse); // WithWindow

                int slideCount = pres.Slides.Count;
                if (slideIndex < 1 || slideIndex > slideCount)
                {
                    error = $"Slide {slideIndex} does not exist (available: 1..{slideCount}).";
                    return null;
                }

                PP.Slide slide = pres.Slides[slideIndex];

                // Export at 2x the on-screen size for a crisp image.
                const double pxPerPoint = 96.0 / 72.0;
                int width = (int)Math.Round(pres.PageSetup.SlideWidth * pxPerPoint * 2);
                int height = (int)Math.Round(pres.PageSetup.SlideHeight * pxPerPoint * 2);

                string outDir = Path.Combine(Path.GetTempPath(), "oPenEfficiency", "SlideSnapshots");
                Directory.CreateDirectory(outDir);
                string png = Path.Combine(outDir, $"slide_{Guid.NewGuid():N}.png");

                slide.Export(png, "PNG", width, height);

                if (!File.Exists(png))
                {
                    error = "Export did not produce a file.";
                    return null;
                }

                return png;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SlideSnapshotService.ExportSlideToPng");
                error = ex.Message;
                return null;
            }
            finally
            {
                try { pres?.Close(); }
                catch (COMException)
                {
                    // The temporary read-only presentation is already closed.
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "SlideSnapshotService.ExportSlideToPng.ClosePresentation");
                }

                // Only quit PowerPoint if we started it AND nothing else is open in it,
                // so we never close a presentation the user is working on.
                try
                {
                    if (createdApp && ppApp != null && ppApp.Presentations.Count == 0)
                    {
                        ppApp.Quit();
                    }
                }
                catch (COMException)
                {
                    // PowerPoint may already be shutting down; do not affect a user session.
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "SlideSnapshotService.ExportSlideToPng.QuitOwnedApplication");
                }
            }
        }

        private static PP.Application GetOrCreatePowerPoint(out bool created)
        {
            created = false;
            try
            {
                // Reuse an already running instance (do not quit it later).
                return (PP.Application)Marshal.GetActiveObject("PowerPoint.Application");
            }
            catch (COMException)
            {
                // Not running: start a dedicated automation instance.
                created = true;
                return new PP.Application();
            }
        }
    }
}
