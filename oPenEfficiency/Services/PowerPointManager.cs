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
            try
            {
                if (_application.ActiveWindow == null) return null;
                var selection = _application.ActiveWindow.Selection;
                if (selection == null || selection.Type != PpSelectionType.ppSelectionShapes) return null;
                return selection.ShapeRange;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PowerPointManager.GetSelectedShapes");
                return null;
            }
        }

        /// <summary>
        /// Gathers metadata about the current selection and slide state.
        /// This is intended to be called once per selection change to avoid redundant COM calls.
        /// </summary>
        public SelectionMetadata GetCurrentSelectionMetadata()
        {
            var meta = new SelectionMetadata();
            try
            {
                if (_application.ActiveWindow == null) return meta;
                
                var selection = _application.ActiveWindow.Selection;
                meta.Type = selection.Type;

                if (meta.Type == PpSelectionType.ppSelectionShapes || meta.Type == PpSelectionType.ppSelectionText)
                {
                    var shapes = selection.ShapeRange;
                    meta.ShapeCount = shapes.Count;
                    
                    for (int i = 1; i <= meta.ShapeCount; i++)
                    {
                        if (shapes[i].Type == Office.MsoShapeType.msoTable)
                        {
                            meta.HasTable = true;
                            break;
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
                ExceptionLogger.Log(ex, "PowerPointManager.GetCurrentSelectionMetadata");
            }
            return meta;
        }

        /// <summary>
        /// Retrieves the Reference shape for alignment and spacing.
        /// Priority: First locked shape in selection -> First selected shape.
        /// </summary>
        public Shape GetReferenceShape()
        {
            var selection = GetSelectedShapes();
            if (selection == null || selection.Count == 0) return null;

            // 1. Check if any shape in the selection is locked (Master shape)
            for (int i = 1; i <= selection.Count; i++)
            {
                if (Features.ShapeLockingFeature.IsShapeLocked(selection[i]))
                {
                    return selection[i];
                }
            }

            // 2. Fallback to the first selected shape
            return selection[1];
        }

        /// <summary>
        /// Gets the current active slide from the active window.
        /// </summary>
        public Slide GetCurrentSlide()
        {
            try
            {
                if (_application.ActiveWindow != null && _application.ActiveWindow.View != null)
                {
                    return _application.ActiveWindow.View.Slide as Slide;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PowerPointManager.GetCurrentSlide");
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
