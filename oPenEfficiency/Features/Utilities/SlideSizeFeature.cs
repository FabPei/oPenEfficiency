using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI.Dialogs;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnSlideSize",
        Name = "Show / Hide Slide Size",
        Tooltip = "Analyze Slide Media Heaviness",
        Color = "#F43F5E",
        MinSelection = 0,
        RequiredType = PowerPoint.PpSelectionType.ppSelectionNone,
        Description = "An analysis tool that gauges the heaviness of media objects on each slide to identify space hogs.",
        Keywords = "size, media, images, video, weight, bloat, heavy",
        DetailedHelpText = "### Show / Hide Slide Size\n\nToggles an overlay displaying the individual file size of every embedded media object (images, videos) on the slide, helping identify large files that bloat presentation size."
    )]
    public static class SlideSizeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var pres = manager.GetApplication().ActivePresentation;
                string tempDir = Path.Combine(Path.GetTempPath(), "oPenEfficiency_MediaScan");

                if (!Directory.Exists(tempDir))
                    Directory.CreateDirectory(tempDir);
                else
                {
                    // Clean previous
                    foreach (var f in Directory.GetFiles(tempDir)) File.Delete(f);
                }

                var items = new List<SlideSizeItem>();

                // Helper dictionary to store heavy shapes we want to label later
                var labelCandidates = new Dictionary<PowerPoint.Shape, double>();

                // First, check if there are ALREADY tracking labels, and if so, REMOVE THEM (Toggle off)
                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    for (int i = slide.Shapes.Count; i >= 1; i--)
                    {
                        var shp = slide.Shapes[i];
                        if (shp.Name.StartsWith("oPE_SizeLabel_"))
                        {
                            shp.Delete();
                        }
                    }
                }

                // If we just removed labels and user just clicked to hide them, we can exit early.
                // But the user requested the Float Window as default. So we continue parsing.

                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    long totalBytes = 0;
                    int mediaCount = 0;

                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape.Type == MsoShapeType.msoPicture || shape.Type == MsoShapeType.msoMedia || shape.Type == MsoShapeType.msoEmbeddedOLEObject)
                        {
                            mediaCount++;
                            string tempFile = Path.Combine(tempDir, $"s{slide.SlideNumber}_shp{shape.Id}.png");
                            
                            try
                            {
                                shape.Export(tempFile, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                                if (File.Exists(tempFile))
                                {
                                    long fSize = new FileInfo(tempFile).Length;
                                    totalBytes += fSize;
                                    
                                    // If greater than 0.5 MB, tag it as heavy
                                    double localMb = fSize / 1048576.0;
                                    if (localMb > 0.5)
                                    {
                                        labelCandidates.Add(shape, localMb);
                                    }
                                }
                            }
                            catch
                            {
                                // Shape can't be exported (e.g. video without thumbnail frame available to API)
                            }
                        }
                    }

                    if (mediaCount > 0)
                    {
                        items.Add(new SlideSizeItem
                        {
                            SlideNumber = slide.SlideNumber,
                            SizeMB = totalBytes / 1048576.0,
                            Details = $"{mediaCount} Media Object(s)"
                        });
                    }
                }

                // Show Float Window
                var win = new SlideSizeWindow(items.OrderByDescending(x => x.SizeMB).ToList());
                bool? dialogResult = win.ShowDialog();

                if (win.CreateLabelsRequested)
                {
                    // Create floating labels (The REQ-28 right-click action equivalent)
                    foreach (var kvp in labelCandidates)
                    {
                        var targetShape = kvp.Key;
                        double mb = kvp.Value;
                        var slide = targetShape.Parent as PowerPoint.Slide;

                        if (slide != null)
                        {
                            var lbl = slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, targetShape.Left, targetShape.Top, 80, 25);
                            lbl.Name = "oPE_SizeLabel_" + Guid.NewGuid().ToString();
                            lbl.Fill.ForeColor.RGB = 0x3232E8; // Red in BGR
                            lbl.Line.Visible = MsoTriState.msoFalse;
                            lbl.TextFrame.TextRange.Text = $"{mb:F1} MB";
                            lbl.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF; // White
                            lbl.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex.Message, "SlideSizeFeature");
                return false;
            }
        }
    }
}
