using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI.Dialogs;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnSlideSize",
        Name = "Analyze Slide Media Size",
        SearchAliases = new[] { "Slide Media Size Analysis", "Slide Size Analysis" },
        Tooltip = "Estimate slide media heaviness",
        Color = "#F43F5E",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None,
        Description = "Estimates the rendered PNG footprint of pictures, media, and embedded OLE objects on each slide.",
        DetailedHelpText = "### Analyze Slide Media Size\nExports render previews of pictures, media, and embedded OLE objects to an isolated temporary folder, then reports an approximation of relative slide heaviness. Values are rendered PNG estimates, not the compressed bytes stored inside the PPTX. Right-click to run the analysis or directly remove all existing analysis labels.")]
    public static class SlideSizeFeature
    {
        private const string LabelTag = "oPE_SlideSizeLabel";
        private const string LegacyLabelPrefix = "oPE_SizeLabel_";

        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            string tempDirectory = null;
            try
            {
                var app = manager.GetApplication();
                var presentation = app?.ActivePresentation;
                if (presentation == null) return false;

                tempDirectory = Path.Combine(Path.GetTempPath(), "oPenEfficiency_MediaScan_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tempDirectory);
                var items = new List<SlideSizeItem>();
                var candidates = new List<LabelCandidate>();

                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    long totalBytes = 0;
                    int analyzed = 0;
                    int failed = 0;
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (!IsMediaShape(shape)) continue;
                        string tempFile = Path.Combine(tempDirectory, $"s{slide.SlideID}_shape{shape.Id}.png");
                        try
                        {
                            shape.Export(tempFile, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                            if (!File.Exists(tempFile)) { failed++; continue; }
                            long bytes = new FileInfo(tempFile).Length;
                            totalBytes += bytes;
                            analyzed++;
                            candidates.Add(new LabelCandidate { SlideId = slide.SlideID, ShapeId = shape.Id, EstimatedMb = bytes / 1048576.0 });
                            File.Delete(tempFile);
                        }
                        catch (Exception ex)
                        {
                            failed++;
                            ExceptionLogger.Log(ex, $"SlideSizeFeature.ExportPreview.Slide{slide.SlideID}.Shape{shape.Id}");
                        }
                    }

                    if (analyzed + failed > 0)
                    {
                        items.Add(new SlideSizeItem
                        {
                            SlideNumber = slide.SlideNumber,
                            SizeMB = totalBytes / 1048576.0,
                            Details = failed == 0 ? $"{analyzed} analyzed" : $"{analyzed} analyzed, {failed} unavailable"
                        });
                    }
                }

                var window = new SlideSizeWindow(items.OrderByDescending(item => item.SizeMB).ToList(), HasLabels(presentation));
                new System.Windows.Interop.WindowInteropHelper(window).Owner = (IntPtr)app.HWND;
                window.ShowDialog();

                if (window.LabelAction == SlideSizeLabelAction.CreateOrRefresh)
                {
                    try { app.StartNewUndoEntry(); }
                    catch (Exception undoEx) { ExceptionPolicy.ReportOptionalComFailure(undoEx, "SlideSizeFeature.Show.StartUndoEntry"); }
                    RemoveLabels(presentation);
                    CreateLabels(presentation, candidates);
                }
                else if (window.LabelAction == SlideSizeLabelAction.Remove)
                {
                    try { app.StartNewUndoEntry(); }
                    catch (Exception undoEx) { ExceptionPolicy.ReportOptionalComFailure(undoEx, "SlideSizeFeature.Refresh.StartUndoEntry"); }
                    RemoveLabels(presentation);
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SlideSizeFeature.Execute");
                throw;
            }
            finally
            {
                if (!string.IsNullOrEmpty(tempDirectory))
                {
                    try { if (Directory.Exists(tempDirectory)) Directory.Delete(tempDirectory, true); }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "SlideSizeFeature.TempCleanup"); }
                }
            }
        }

        public static bool RemoveAnalysisLabels(PowerPointManager manager)
        {
            try
            {
                var app = manager?.GetApplication();
                var presentation = app?.ActivePresentation;
                if (presentation == null || !HasLabels(presentation)) return false;
                try { app.StartNewUndoEntry(); }
                catch (Exception undoEx) { ExceptionPolicy.ReportOptionalComFailure(undoEx, "SlideSizeFeature.Hide.StartUndoEntry"); }
                RemoveLabels(presentation);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SlideSizeFeature.RemoveAnalysisLabels");
                return false;
            }
        }

        private static bool IsMediaShape(PowerPoint.Shape shape)
        {
            return shape.Type == MsoShapeType.msoPicture || shape.Type == MsoShapeType.msoMedia ||
                   shape.Type == MsoShapeType.msoEmbeddedOLEObject || shape.Type == MsoShapeType.msoLinkedPicture;
        }

        private static bool HasLabels(PowerPoint.Presentation presentation)
        {
            foreach (PowerPoint.Slide slide in presentation.Slides)
                foreach (PowerPoint.Shape shape in slide.Shapes)
                    if (IsLabel(shape)) return true;
            return false;
        }

        private static void RemoveLabels(PowerPoint.Presentation presentation)
        {
            foreach (PowerPoint.Slide slide in presentation.Slides)
                for (int i = slide.Shapes.Count; i >= 1; i--)
                    if (IsLabel(slide.Shapes[i])) slide.Shapes[i].Delete();
        }

        private static bool IsLabel(PowerPoint.Shape shape)
        {
            try
            {
                if (string.Equals(shape.Tags[LabelTag], "true", StringComparison.OrdinalIgnoreCase)) return true;
            }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "SlideSizeFeature.IsLabelShape.Tag"); }
            return shape.Name != null && shape.Name.StartsWith(LegacyLabelPrefix, StringComparison.OrdinalIgnoreCase);
        }

        private static void CreateLabels(PowerPoint.Presentation presentation, IEnumerable<LabelCandidate> candidates)
        {
            foreach (var candidate in candidates)
            {
                var slide = FindSlide(presentation, candidate.SlideId);
                var target = FindShape(slide, candidate.ShapeId);
                if (target == null) continue;
                var label = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, target.Left, target.Top, 90, 25);
                label.Name = LegacyLabelPrefix + Guid.NewGuid().ToString("N");
                label.Tags.Add(LabelTag, "true");
                label.Fill.ForeColor.RGB = 0x3232E8;
                label.Line.Visible = MsoTriState.msoFalse;
                label.TextFrame.TextRange.Text = $"~{candidate.EstimatedMb:F2} MB";
                label.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF;
                label.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            }
        }

        private static PowerPoint.Slide FindSlide(PowerPoint.Presentation presentation, int slideId)
        {
            foreach (PowerPoint.Slide slide in presentation.Slides) if (slide.SlideID == slideId) return slide;
            return null;
        }

        private static PowerPoint.Shape FindShape(PowerPoint.Slide slide, int shapeId)
        {
            if (slide == null) return null;
            foreach (PowerPoint.Shape shape in slide.Shapes) if (shape.Id == shapeId) return shape;
            return null;
        }

        private sealed class LabelCandidate
        {
            public int SlideId { get; set; }
            public int ShapeId { get; set; }
            public double EstimatedMb { get; set; }
        }
    }
}
