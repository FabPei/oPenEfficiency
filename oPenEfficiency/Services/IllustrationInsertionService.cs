using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Services
{
    /// <summary>Outcome of one insertion, phrased for the window's status line.</summary>
    public sealed class IllustrationInsertionResult
    {
        public bool Inserted { get; set; }

        public string Message { get; set; }

        /// <summary>Number of shapes that ended up selected by the insertion.</summary>
        public int ShapeCount { get; set; }

        /// <summary>True when the artwork could not be converted and stayed an SVG picture.</summary>
        public bool StayedPicture { get; set; }

        /// <summary>True when PowerPoint accepted the explicit boundary for this insertion's undo unit.</summary>
        public bool UndoEntryStarted { get; set; }
    }

    /// <summary>
    /// Places one illustration on the active slide. The SVG text is prepared host-free first
    /// (<see cref="IllustrationInsertionPlanner"/>), written to a temp file, inserted as a vector picture
    /// and - when the user asked for editable shapes - converted through PowerPoint's own SVGEdit command.
    /// When that command is unavailable the artwork stays a tagged SVG picture; ungrouping is not a
    /// substitute because a picture is not a group. The inserted shape is tagged with its item id and
    /// collection so the license trail stays attached to the artwork.
    /// </summary>
    public sealed class IllustrationInsertionService
    {
        public IllustrationInsertionResult Insert(
            PowerPointManager manager,
            IllustrationItem item,
            string svgText,
            IllustrationInsertOptions options,
            string sourcePath = null)
        {
            var result = new IllustrationInsertionResult();
            if (manager == null)
            {
                result.Message = "The insert path is not connected to PowerPoint.";
                return result;
            }

            if (item == null)
            {
                result.Message = "No illustration was selected.";
                return result;
            }

            var effective = options ?? new IllustrationInsertOptions();
            string prepared = IllustrationInsertionPlanner.CreateInsertSvg(svgText, effective);
            if (string.IsNullOrWhiteSpace(prepared))
            {
                result.Message = "This file could not be prepared for insertion.";
                return result;
            }

            string tempFile = null;
            try
            {
                var app = manager.GetApplication();
                if (app?.ActivePresentation == null || app.ActiveWindow == null)
                {
                    result.Message = "Open a presentation in normal view first.";
                    return result;
                }

                PowerPoint.Slide slide = null;
                if (app.ActiveWindow.ViewType == PowerPoint.PpViewType.ppViewNormal)
                {
                    slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                }

                if (slide == null)
                {
                    result.Message = "Switch to a normal slide view first.";
                    return result;
                }

                // When the choices ask for no change, or the prepared copy came out identical to the source, the
                // original file is inserted directly and no copy is written.
                string insertPath;
                if ((IllustrationInsertionPlanner.IsNoOp(effective) ||
                     IllustrationInsertionPlanner.ChangesNothing(svgText, prepared)) &&
                    !string.IsNullOrWhiteSpace(sourcePath) &&
                    File.Exists(sourcePath))
                {
                    insertPath = sourcePath;
                }
                else
                {
                    tempFile = WriteTempSvg(prepared);
                    insertPath = tempFile;
                }

                if (string.IsNullOrWhiteSpace(insertPath))
                {
                    result.Message = "The temporary file for the insertion could not be written.";
                    return result;
                }

                float width = effective.DesiredWidth;
                float height = ResolveHeight(prepared, app.ActivePresentation, width);
                float left = Math.Max(0f, (app.ActivePresentation.PageSetup.SlideWidth - width) / 2f);
                float top = Math.Max(0f, (app.ActivePresentation.PageSetup.SlideHeight - height) / 2f);

                try
                {
                    app.StartNewUndoEntry();
                    result.UndoEntryStarted = true;
                }
                catch (Exception ex)
                {
                    oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationInsertionService.Insert.StartNewUndoEntry");
                }

                var shape = slide.Shapes.AddPicture(
                    insertPath,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue,
                    left,
                    top,
                    width,
                    height);

                if (shape == null)
                {
                    result.Message = "PowerPoint did not accept the illustration.";
                    return result;
                }

                shape.Name = BuildShapeName(item);
                TagShape(shape, item);

                var target = shape;
                if (effective.Decompose)
                {
                    // Captured before the conversion: SVGEdit's group bounds follow the artwork's geometry
                    // (strokes can overhang the viewBox), so the converted result is put back into the
                    // exact footprint the inserted picture showed.
                    float pictureLeft = shape.Left;
                    float pictureTop = shape.Top;
                    float pictureWidth = shape.Width;
                    float pictureHeight = shape.Height;
                    var converted = TryConvertToShapes(app, shape);
                    if (converted != null)
                    {
                        // Re-fit, re-tag and re-name inside one explicit undo unit so the editable insertion has a
                        // deterministic chain: insertion, conversion, cosmetics. Without the boundary each
                        // property write opens its own unit and one Ctrl+Z strips only the tag.
                        try
                        {
                            app.StartNewUndoEntry();
                        }
                        catch (Exception ex)
                        {
                            oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationInsertionService.Insert.CosmeticsUndoEntry");
                        }
                        MatchOriginalBounds(converted, pictureLeft, pictureTop, pictureWidth, pictureHeight);
                        TagShape(converted, item);
                        // Keep the provenance visible: SVGEdit replaces the named picture with its own
                        // result shapes, so the converted target takes over the illustration source name
                        // the Selection Pane shows for the picture route (owner finding 2026-10-19).
                        converted.Name = BuildShapeName(item);
                        target = converted;
                    }
                    else
                    {
                        result.StayedPicture = true;
                        result.Message = "Inserted as an SVG picture - converting to editable shapes is not available for this file in this Office build.";
                    }
                }

                TrySelectShape(slide, target);
                result.Inserted = true;
                result.ShapeCount = 1;
                if (string.IsNullOrWhiteSpace(result.Message))
                {
                    result.Message = effective.Decompose
                        ? "Inserted as editable shapes and tagged with its source."
                        : "Inserted as a scalable SVG picture.";
                }

                return result;
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationInsertionService.Insert");
                result.Message = "Insertion failed: " + ex.Message;
                return result;
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(tempFile))
                {
                    try
                    {
                        if (File.Exists(tempFile)) File.Delete(tempFile);
                    }
                    catch (Exception ex)
                    {
                        oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationInsertionService.Insert.DeleteTemp");
                    }
                }
            }
        }

        private static string BuildShapeName(IllustrationItem item)
        {
            string name = string.IsNullOrWhiteSpace(item.Name) ? item.Id : item.Name;
            return "Illustration " + name;
        }

        private static void TagShape(PowerPoint.Shape shape, IllustrationItem item)
        {
            if (shape == null || item == null) return;

            try
            {
                ShapeMetadataService.SetTag(shape, "oPE_Illustration", item.Id);
                ShapeMetadataService.SetTag(shape, "oPE_IllustrationCollection", item.CollectionId);
                ShapeMetadataService.SetTag(shape, "oPE_IllustrationUpdatedUtc", DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'"));
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationInsertionService.TagShape");
            }
        }

        private static float ResolveHeight(string svgText, PowerPoint.Presentation presentation, float width)
        {
            float viewWidth;
            float viewHeight;
            if (!IllustrationInsertionPlanner.TryReadViewBox(svgText, out viewWidth, out viewHeight) || viewWidth <= 0f)
            {
                return width;
            }

            float height = width * (viewHeight / viewWidth);
            float slideHeight = presentation == null ? height : presentation.PageSetup.SlideHeight;
            return slideHeight > 0f ? Math.Min(height, slideHeight) : height;
        }

        private static string WriteTempSvg(string svgText)
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "oPenEfficiency", "Illustrations");
            Directory.CreateDirectory(tempDir);

            string tempFile = Path.Combine(tempDir, Guid.NewGuid().ToString("N") + ".svg");
            File.WriteAllText(tempFile, svgText);
            return tempFile;
        }

        /// <summary>
        /// Converts the inserted picture through PowerPoint's SVGEdit command (the same path the icon browser
        /// uses). The command's enabled-state follows the selection, so the picture is selected before the
        /// probe and one retry covers a ribbon state that lags an external automation call; a command that
        /// silently does nothing is reported as unavailable instead of returning the unconverted picture.
        /// </summary>
        private static PowerPoint.Shape TryConvertToShapes(PowerPoint.Application app, PowerPoint.Shape picture)
        {
            try
            {
                // Read before the command runs: when SVGEdit succeeds it deletes the picture, and an Id
                // read from the deleted shape would throw instead of comparing.
                int pictureId = picture.Id;
                for (int attempt = 1; attempt <= 2; attempt++)
                {
                    try
                    {
                        picture.Select();
                    }
                    catch (Exception ex) when (oPenEfficiency.Utils.ExceptionPolicy.IsExpectedComFailure(ex))
                    {
                        return null;
                    }

                    // A failed probe means "cannot ask this host"; the execution attempt then decides.
                    if (!oPenEfficiency.Utils.ExceptionPolicy.TryComProbe(
                            () => app.CommandBars.GetEnabledMso("SVGEdit"),
                            out bool enabled) ||
                        enabled)
                    {
                        try
                        {
                            app.CommandBars.ExecuteMso("SVGEdit");
                        }
                        catch (Exception ex) when (oPenEfficiency.Utils.ExceptionPolicy.IsExpectedComFailure(ex))
                        {
                            return null;
                        }

                        PowerPoint.Shape converted = FindConvertedShape(app, pictureId);
                        if (converted != null)
                        {
                            return converted;
                        }
                    }

                    if (attempt < 2)
                    {
                        System.Threading.Thread.Sleep(300);
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationInsertionService.TryConvertToShapes");
                return null;
            }
        }

        /// <summary>
        /// Returns the shape or group the SVGEdit command selected, or null when the selection still holds
        /// the original picture - the command's signature for "nothing to convert here".
        /// </summary>
        private static PowerPoint.Shape FindConvertedShape(PowerPoint.Application app, int pictureId)
        {
            bool read = oPenEfficiency.Utils.ExceptionPolicy.TryComProbe(
                (Func<PowerPoint.Shape>)(() => ReadConvertedShape(app, pictureId)),
                out PowerPoint.Shape converted);
            return read ? converted : null;
        }

        private static PowerPoint.Shape ReadConvertedShape(PowerPoint.Application app, int pictureId)
        {
            PowerPoint.Selection selection = app.ActiveWindow == null ? null : app.ActiveWindow.Selection;
            if (selection == null ||
                selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes ||
                selection.ShapeRange == null ||
                selection.ShapeRange.Count == 0)
            {
                return null;
            }

            PowerPoint.Shape primary = selection.ShapeRange[1];
            if (primary == null || primary.Id == pictureId)
            {
                return null;
            }

            return selection.ShapeRange.Count == 1 ? primary : selection.ShapeRange.Group();
        }

        private static void MatchOriginalBounds(PowerPoint.Shape shape, float left, float top, float width, float height)
        {
            try
            {
                shape.Left = left;
                shape.Top = top;
                shape.Width = width;
                shape.Height = height;
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationInsertionService.MatchOriginalBounds");
            }
        }

        private static void TrySelectShape(PowerPoint.Slide slide, PowerPoint.Shape shape)
        {
            if (slide == null || shape == null) return;

            try
            {
                slide.Shapes.Range(new[] { shape.Name }).Select();
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationInsertionService.TrySelectShape");
            }
        }
    }
}
