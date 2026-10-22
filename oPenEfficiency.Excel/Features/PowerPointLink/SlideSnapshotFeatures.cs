using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using WinForms = System.Windows.Forms;
using Xl = global::Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using oPenEfficiency.Excel.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Excel.Features.PowerPointLink
{
    /// <summary>
    /// Shared bits for the linked slide-snapshot feature: the tag stored on each
    /// inserted picture so it can be refreshed later, plus selection helpers.
    /// The tag lives in Shape.AlternativeText as "oPE_SlideLink|&lt;path&gt;|&lt;slideIndex&gt;".
    /// </summary>
    internal static class SlideLink
    {
        public const string TagPrefix = "oPE_SlideLink";

        public static bool IsTagged(Xl.Shape shape)
        {
            try
            {
                var alt = shape.AlternativeText;
                return !string.IsNullOrEmpty(alt) && alt.StartsWith(TagPrefix + "|", StringComparison.Ordinal);
            }
            catch (COMException)
            {
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SlideLink.IsTagged");
                return false;
            }
        }

        public static bool TryParse(Xl.Shape shape, out string pptxPath, out int slideIndex)
        {
            pptxPath = null;
            slideIndex = 0;
            try
            {
                var alt = shape.AlternativeText;
                if (string.IsNullOrEmpty(alt)) return false;
                var parts = alt.Split('|');
                if (parts.Length < 3 || parts[0] != TagPrefix) return false;
                pptxPath = parts[1];
                return int.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out slideIndex);
            }
            catch (COMException)
            {
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SlideLink.TryParse");
                return false;
            }
        }

        public static string BuildTag(string pptxPath, int slideIndex)
            => $"{TagPrefix}|{pptxPath}|{slideIndex.ToString(CultureInfo.InvariantCulture)}";

        /// <summary>Tagged snapshot shapes in the current selection, if any.</summary>
        public static List<Xl.Shape> GetSelectedTaggedShapes(Xl.Application app)
        {
            var result = new List<Xl.Shape>();
            try
            {
                dynamic selection = app.Selection;
                Xl.ShapeRange shapeRange = null;
                try { shapeRange = selection.ShapeRange as Xl.ShapeRange; }
                catch (COMException)
                {
                    // The current Excel selection is not a shape selection.
                }
                if (shapeRange != null)
                {
                    foreach (Xl.Shape shape in shapeRange)
                    {
                        if (IsTagged(shape)) result.Add(shape);
                    }
                }
            }
            catch (COMException)
            {
                // A stale selection is treated as containing no tagged shapes.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SlideLink.GetSelectedTaggedShapes");
            }
            return result;
        }

        /// <summary>All tagged snapshot shapes on a worksheet.</summary>
        public static List<Xl.Shape> GetAllTaggedShapes(Xl.Worksheet sheet)
        {
            var result = new List<Xl.Shape>();
            try
            {
                foreach (Xl.Shape shape in sheet.Shapes)
                {
                    if (IsTagged(shape)) result.Add(shape);
                }
            }
            catch (COMException)
            {
                // A worksheet closing during enumeration yields the valid shapes found so far.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SlideLink.GetAllTaggedShapes");
            }
            return result;
        }
    }

    /// <summary>
    /// Inserts a PNG snapshot of a chosen PowerPoint slide into the active worksheet
    /// and links it (source file + slide index) so it can be refreshed later.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnInsertSlideSnapshot",
        Name = "Insert Slide",
        Tooltip = "Insert a PowerPoint slide as an image",
        IconData = "",
        Color = "#D24726",
        Description = "Inserts an image of a PowerPoint slide into Excel and stores the file path plus slide number so it can be refreshed later.",
        RequiredType = OfficeSelectionType.None)]
    public static class InsertSlideSnapshotFeature
    {
        public static bool Execute(ExcelManager manager)
        {
            try
            {
                var app = manager?.Application;
                var sheet = manager?.ActiveSheet;
                if (app == null || sheet == null) return false;

                // 1. Pick the PowerPoint file.
                string pptxPath;
                using (var dlg = new WinForms.OpenFileDialog
                {
                    Title = "Choose PowerPoint File",
                    Filter = "PowerPoint Presentations|*.pptx;*.pptm;*.ppt|All Files|*.*"
                })
                {
                    if (dlg.ShowDialog() != WinForms.DialogResult.OK) return false;
                    pptxPath = dlg.FileName;
                }

                // 2. Ask for the slide number (native Excel InputBox, Type 1 = number).
                object input = app.InputBox("Slide number:", "Insert Slide as Image", 1, Type: 1);
                if (input is bool) return false; // user cancelled
                int slideIndex = Convert.ToInt32(input);
                if (slideIndex < 1)
                {
                    WinForms.MessageBox.Show("Please enter a slide number greater than or equal to 1.");
                    return false;
                }

                // 3. Export the slide to PNG.
                string png = SlideSnapshotService.ExportSlideToPng(pptxPath, slideIndex, out string error);
                if (png == null)
                {
                    WinForms.MessageBox.Show("The slide could not be exported:\n" + error);
                    return false;
                }

                // 4. Position at the selected cell; size to a sensible width keeping aspect.
                float left = 100f, top = 100f;
                try
                {
                    var target = manager.SelectedRange;
                    if (target != null)
                    {
                        left = Convert.ToSingle(target.Left);
                        top = Convert.ToSingle(target.Top);
                    }
                }
                catch (COMException)
                {
                    // Use the default insertion position when Excel has no usable range.
                }
                catch (InvalidCastException)
                {
                    // Mixed selection coordinates are not scalar values.
                }

                const float displayWidth = 480f;
                float displayHeight = displayWidth * 9f / 16f;
                try
                {
                    using (var img = System.Drawing.Image.FromFile(png))
                    {
                        if (img.Width > 0)
                            displayHeight = displayWidth * img.Height / (float)img.Width;
                    }
                }
                catch (ArgumentException ex)
                {
                    ExceptionLogger.Log(ex, "InsertSlideSnapshotFeature.ReadExportDimensions");
                }
                catch (IOException ex)
                {
                    ExceptionLogger.Log(ex, "InsertSlideSnapshotFeature.ReadExportDimensions");
                }

                Xl.Shape shape = sheet.Shapes.AddPicture(
                    png,
                    MsoTriState.msoFalse, // LinkToFile
                    MsoTriState.msoTrue,  // SaveWithDocument
                    left, top, displayWidth, displayHeight);

                shape.AlternativeText = SlideLink.BuildTag(pptxPath, slideIndex);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertSlideSnapshotFeature.Execute");
                return false;
            }
        }
    }

    /// <summary>
    /// Refreshes linked slide snapshots: re-exports the source slide and replaces the
    /// picture in place (same position, size and rotation). Acts on the selected
    /// snapshot(s); if none are selected, refreshes all snapshots on the active sheet.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnRefreshSlideSnapshot",
        Name = "Refresh Slide",
        Tooltip = "Refresh a linked slide image",
        IconData = "",
        Color = "#D24726",
        Description = "Re-exports the linked PowerPoint slide and replaces the selected image at the same position and size. If nothing is selected, all linked images on the sheet are refreshed.",
        RequiredType = OfficeSelectionType.None)]
    public static class RefreshSlideSnapshotFeature
    {
        public static bool Execute(ExcelManager manager)
        {
            try
            {
                var app = manager?.Application;
                var sheet = manager?.ActiveSheet;
                if (app == null || sheet == null) return false;

                var targets = SlideLink.GetSelectedTaggedShapes(app);
                if (targets.Count == 0)
                {
                    targets = SlideLink.GetAllTaggedShapes(sheet);
                }

                if (targets.Count == 0)
                {
                    WinForms.MessageBox.Show("No linked slide image found. Please insert a slide first.");
                    return false;
                }

                int refreshed = 0;
                string lastError = null;
                foreach (var shape in targets)
                {
                    if (RefreshShape(sheet, shape, out string err)) refreshed++;
                    else if (err != null) lastError = err;
                }

                if (refreshed == 0 && lastError != null)
                {
                    WinForms.MessageBox.Show("Refresh failed:\n" + lastError);
                }

                return refreshed > 0;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "RefreshSlideSnapshotFeature.Execute");
                return false;
            }
        }

        private static bool RefreshShape(Xl.Worksheet sheet, Xl.Shape shape, out string error)
        {
            error = null;
            if (!SlideLink.TryParse(shape, out string pptxPath, out int slideIndex))
            {
                return false;
            }

            string png = SlideSnapshotService.ExportSlideToPng(pptxPath, slideIndex, out error);
            if (png == null) return false;

            Xl.Shape replacement = null;
            bool originalDeleted = false;
            try
            {
                // Capture geometry of the existing picture.
                float left = Convert.ToSingle(shape.Left);
                float top = Convert.ToSingle(shape.Top);
                float width = Convert.ToSingle(shape.Width);
                float height = Convert.ToSingle(shape.Height);
                float rotation = Convert.ToSingle(shape.Rotation);
                string name = shape.Name;
                string tag = shape.AlternativeText;

                replacement = sheet.Shapes.AddPicture(
                    png,
                    MsoTriState.msoFalse,
                    MsoTriState.msoTrue,
                    left, top, width, height);

                try { replacement.Rotation = rotation; }
                catch (COMException)
                {
                    // Rotation is optional; preserve the refreshed image and its geometry.
                }
                replacement.AlternativeText = tag;

                try
                {
                    shape.Delete();
                    originalDeleted = true;
                }
                catch (COMException ex)
                {
                    TryDeleteReplacement(replacement);
                    replacement = null;
                    ExceptionLogger.Log(ex, "RefreshSlideSnapshotFeature.RefreshShape.DeleteOriginal");
                    error = "The original snapshot could not be replaced; the new image was rolled back.";
                    return false;
                }

                try { replacement.Name = name; }
                catch (COMException)
                {
                    // Excel can reject a duplicate or reserved shape name; the link tag remains authoritative.
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "RefreshSlideSnapshotFeature.RefreshShape.RestoreName");
                }

                return true;
            }
            catch (Exception ex)
            {
                if (!originalDeleted) TryDeleteReplacement(replacement);
                ExceptionLogger.Log(ex, "RefreshSlideSnapshotFeature.RefreshShape");
                error = ex.Message;
                return false;
            }
        }

        private static void TryDeleteReplacement(Xl.Shape replacement)
        {
            if (replacement == null) return;
            try { replacement.Delete(); }
            catch (COMException)
            {
                // The rollback image is already absent.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "RefreshSlideSnapshotFeature.RollbackReplacement");
            }
        }
    }
}
