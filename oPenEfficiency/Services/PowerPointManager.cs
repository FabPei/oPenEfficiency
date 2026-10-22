using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;

namespace oPenEfficiency
{
    /// <summary>
    /// Lightweight metadata about the current selection to avoid repeated COM calls.
    /// </summary>
    public class SelectionMetadata
    {
        public PpSelectionType Type { get; set; }
        public int ShapeCount { get; set; }
        public int TextShapeCount { get; set; }
        public int TextOrTableShapeCount { get; set; }
        public bool HasTextRange { get; set; }
        public bool HasTable { get; set; }
        public bool HasHiddenShapesOnSlide { get; set; }
        public Slide CurrentSlide { get; set; }
    }

    /// <summary>
    /// Core manager for PowerPoint operations.
    /// Acts as a central access point for the PowerPoint COM application and selection-based helpers.
    /// Most feature-specific logic has been migrated to the oPenEfficiency.Features namespace.
    /// </summary>
    public class PowerPointManager
    {
        private readonly Application _application;

        public PowerPointManager(Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public Application GetApplication() => _application;

        /// <summary>
        /// Retrieves the currently selected shapes as a ShapeRange.
        /// Returns null if no shapes are selected.
        /// </summary>
        public ShapeRange GetSelectedShapes()
        {
            if (ExceptionPolicy.TryComProbe(
                () =>
                {
                    var window = _application.ActiveWindow;
                    var selection = window?.Selection;
                    if (selection == null || selection.Type != PpSelectionType.ppSelectionShapes) return null;
                    return selection.ShapeRange;
                },
                out ShapeRange shapes))
            {
                return shapes;
            }

            return null;
        }

        /// <summary>
        /// Gathers metadata about the current selection and slide state.
        /// This is intended to be called once per selection change to avoid redundant COM calls.
        /// </summary>
        public SelectionMetadata GetCurrentSelectionMetadata()
        {
            var meta = new SelectionMetadata();
            // Transient COM RCWs are released once their values are copied into meta. These were
            // previously left to the GC on every selection probe; in a long session that leaks an
            // Application-window + Selection + ShapeRange RCW per call and keeps the out-of-process
            // PowerPoint server alive past Quit() (REG-LIFECYCLE-CRASH-001 forced-cleanup evidence).
            // The releaseWindow/releaseSelection handles are method-scoped copies of the probed
            // window/selection so a single outer finally can release them on every path while the
            // probed reads stay inside the guarded try (expected COM failures still classify as
            // optional below).
            DocumentWindow releaseWindow = null;
            Selection releaseSelection = null;
            ShapeRange shapes = null;
            try
            {
                if (!ExceptionPolicy.TryComProbe(
                    () => _application.ActiveWindow,
                    out DocumentWindow window) || window == null)
                {
                    return meta;
                }
                releaseWindow = window;

                var selection = window.Selection;
                releaseSelection = selection;
                meta.Type = selection.Type;

                if (meta.Type == PpSelectionType.ppSelectionText)
                {
                    try { meta.HasTextRange = selection.TextRange != null; }
                    catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "PowerPointManager.SelectionMetadata.TextRange"); }
                }

                if (meta.Type == PpSelectionType.ppSelectionShapes || meta.Type == PpSelectionType.ppSelectionText)
                {
                    shapes = selection.ShapeRange;
                    meta.ShapeCount = shapes.Count;

                    for (int i = 1; i <= meta.ShapeCount; i++)
                    {
                        var shape = shapes[i];
                        try
                        {
                            bool isTable = shape.Type == Office.MsoShapeType.msoTable;
                            bool hasText = shape.HasTextFrame == Office.MsoTriState.msoTrue;
                            if (isTable) meta.HasTable = true;
                            if (hasText) meta.TextShapeCount++;
                            if (hasText || isTable) meta.TextOrTableShapeCount++;
                        }
                        finally
                        {
                            ExceptionPolicy.RunBestEffortComCleanup(
                                () => { if (shape != null && Marshal.IsComObject(shape)) Marshal.ReleaseComObject(shape); },
                                "PowerPointManager.GetCurrentSelectionMetadata.ReleaseShape");
                        }
                    }
                }

                var slide = GetCurrentSlide();
                meta.CurrentSlide = slide;
                if (slide != null)
                {
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        if (slide.Shapes[i].Visible == Office.MsoTriState.msoFalse)
                        {
                            meta.HasHiddenShapesOnSlide = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Selection state is optional UI state: when the document window disappears while this
                // probe runs (final-presentation close, shutdown), the expected COM failure must stay
                // silent like the guarded ActiveWindow probe above; everything else still logs.
                ExceptionPolicy.ReportOptionalComFailure(ex, "PowerPointManager.GetCurrentSelectionMetadata");
            }
            finally
            {
                ExceptionPolicy.RunBestEffortComCleanup(
                    () => { if (shapes != null && Marshal.IsComObject(shapes)) Marshal.ReleaseComObject(shapes); },
                    "PowerPointManager.GetCurrentSelectionMetadata.ReleaseShapes");
                ExceptionPolicy.RunBestEffortComCleanup(
                    () => { if (releaseSelection != null && Marshal.IsComObject(releaseSelection)) Marshal.ReleaseComObject(releaseSelection); },
                    "PowerPointManager.GetCurrentSelectionMetadata.ReleaseSelection");
                ExceptionPolicy.RunBestEffortComCleanup(
                    () => { if (releaseWindow != null && Marshal.IsComObject(releaseWindow)) Marshal.ReleaseComObject(releaseWindow); },
                    "PowerPointManager.GetCurrentSelectionMetadata.ReleaseWindow");
            }
            return meta;
        }

        /// <summary>
        /// Retrieves the first selected shape as the "Reference" shape for alignment and spacing.
        /// </summary>
        public Shape GetReferenceShape()
        {
            var selection = GetSelectedShapes();
            if (selection == null || selection.Count == 0) return null;
            if (UsesLastSelectedAnchor())
            {
                try
                {
                    return selection[selection.Count];
                }
                catch (Exception ex)
                {
                    // Preserve the historical behavior: a COM failure on the last-element indexer
                    // degrades to the first shape with a log instead of propagating to callers.
                    ExceptionLogger.Log(ex, "PowerPointManager.GetReferenceShape");
                }
            }
            return selection[1];
        }

        /// <summary>
        /// Whether the configured anchor object is the last selected shape ("Last Selected") instead of
        /// the first. This is the single decision point for the reference-shape rule; the automation
        /// seams derive their anchor expectation from the raw configuration independently and cross-check
        /// against this method so an inverted rule cannot pass unnoticed.
        /// </summary>
        public bool UsesLastSelectedAnchor()
        {
            try
            {
                var config = JsonService.Load<Models.SidebarConfig>("sidebar_config.json");
                return AnchorModes.IsLastSelected(config?.MasterObjectMode);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PowerPointManager.UsesLastSelectedAnchor");
                return false;
            }
        }

        /// <summary>
        /// Gets the current active slide from the active window. PowerPoint can
        /// legitimately expose a window before a slide is available, so the
        /// lookup is an expected COM probe rather than an error boundary.
        /// </summary>
        public Slide GetCurrentSlide()
        {
            if (ExceptionPolicy.TryComProbe(
                () =>
                {
                    var window = _application.ActiveWindow;
                    var view = window?.View;
                    return view?.Slide as Slide;
                },
                out Slide slide))
            {
                return slide;
            }

            return null;
        }

        /// <summary>
        /// Retrieves the currently selected slides or the slide containing the current selection.
        /// </summary>
        public List<Slide> GetSelectedSlidesForOperation()
        {
            var slides = new List<Slide>();
            try
            {
                var selection = _application.ActiveWindow?.Selection;
                if (selection == null) return slides;

                if (selection.Type == PpSelectionType.ppSelectionSlides)
                {
                    foreach (Slide s in selection.SlideRange) slides.Add(s);
                }
                else if (selection.Type == PpSelectionType.ppSelectionShapes || selection.Type == PpSelectionType.ppSelectionText)
                {
                    if (_application.ActiveWindow.View.Slide is Slide s)
                    {
                        slides.Add(s);
                    }
                }
                return slides;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PowerPointManager.GetSelectedSlidesForOperation");
                return slides;
            }
        }

        /// <summary>
        /// Gets the current active presentation from the application.
        /// </summary>
        public Presentation GetCurrentPresentation()
        {
            try
            {
                return _application.ActivePresentation;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PowerPointManager.GetCurrentPresentation");
                return null;
            }
        }

        /// <summary>
        /// Recursively gathers all shapes from a shape or group.
        /// </summary>
        public void GetAllShapesRecursive(Shape parent, List<Shape> result)
        {
            if (parent.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (Shape child in parent.GroupItems) GetAllShapesRecursive(child, result);
            }
            else
            {
                result.Add(parent);
            }
        }

        /// <summary>
        /// Gets the screen coordinates (in pixels) of the current selection's center.
        /// Useful for positioning windows near the user's focus.
        /// </summary>
        public void GetSelectionScreenLocation(out int x, out int y) => GetSelectionScreenPos(out x, out y, true);

        /// <summary>
        /// Gets the screen coordinates (in pixels) of the current selection's top-center.
        /// Useful for positioning windows just above the object.
        /// </summary>
        public void GetSelectionScreenTopCenterLocation(out int x, out int y) => GetSelectionScreenPos(out x, out y, false);

        private void GetSelectionScreenPos(out int x, out int y, bool verticalCenter)
        {
            x = 0; y = 0;
            try
            {
                if (_application.ActiveWindow == null) return;
                var selection = _application.ActiveWindow.Selection;
                float left = 0, top = 0, width = 0, height = 0;

                if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    left = selection.ShapeRange.Left;
                    top = selection.ShapeRange.Top;
                    width = selection.ShapeRange.Width;
                    height = selection.ShapeRange.Height;
                }
                else if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    left = selection.TextRange.BoundLeft;
                    top = selection.TextRange.BoundTop;
                    width = selection.TextRange.BoundWidth;
                    height = selection.TextRange.BoundHeight;
                }
                else
                {
                    // Fallback to center of slide
                    left = _application.ActivePresentation.PageSetup.SlideWidth / 2;
                    top = _application.ActivePresentation.PageSetup.SlideHeight / 2;
                }

                x = _application.ActiveWindow.PointsToScreenPixelsX(left + (width / 2));
                float verticalOffset = verticalCenter ? (height / 2) : 0;
                y = _application.ActiveWindow.PointsToScreenPixelsY(top + verticalOffset);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PowerPointManager.GetSelectionScreenPos");
                x = 500; y = 400;
            }
        }

        /// <summary>
        /// Safely releases a COM object and sets the reference to null.
        /// Use this pattern consistently to prevent COM object leaks.
        /// </summary>
        public static void ReleaseComObject(ref object comObject, string debugName = null)
        {
            if (comObject != null)
            {
                try
                {
                    Marshal.ReleaseComObject(comObject);
                    if (debugName != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"Released COM object: {debugName}");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to release COM object {debugName}: {ex.Message}");
                }
                finally
                {
                    comObject = null;
                }
            }
        }
    }
}
