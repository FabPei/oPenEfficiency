using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency;
using oPenEfficiency.Utils;
using MsoLineDashStyle = Microsoft.Office.Core.MsoLineDashStyle;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Point = System.Drawing.Point;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Manages snap-to-objects functionality for shape dragging.
    /// When enabled, moving shapes will snap to edges and centers of nearby objects.
    /// </summary>
    public class SnapToObjectsManager : IDisposable
    {
        // Snap threshold in pixels - shapes snap when within this distance
        private const int SnapThresholdPixels = 10;

        // Minimum drag distance before snap activates (prevents accidental triggers)
        private const int DragThresholdPixels = 4;

        private readonly PowerPointManager _powerPointManager;
        private readonly Win32MouseHook _mouseHook;
        private bool _isEnabled;
        private bool _disposed;

        // Drag state
        private Shape _draggedShape;
        private Point _grabOffset; // Offset from shape's top-left to mouse position
        private float _originalLeft;
        private float _originalTop;
        private int _mouseDownScreenX;
        private int _mouseDownScreenY;
        private bool _dragReady;
        private bool _dragActive;

        // Pending move state
        private int _pendingScreenX;
        private int _pendingScreenY;
        private bool _moveUpdatePending;

        // Visual feedback
        private Shape _highlightShape;
        private Shape _snapGuideHorizontal;
        private Shape _snapGuideVertical;

        // Snap point structure
        private struct SnapPoint
        {
            public float X;
            public float Y;
            public Shape Owner;
            public SnapType Type;

            public SnapPoint(float x, float y, Shape owner, SnapType type)
            {
                X = x;
                Y = y;
                Owner = owner;
                Type = type;
            }
        }

        private enum SnapType
        {
            LeftEdge,
            RightEdge,
            CenterX,
            TopEdge,
            BottomEdge,
            CenterY
        }

        public SnapToObjectsManager(PowerPointManager powerPointManager)
        {
            _powerPointManager = powerPointManager;
            _isEnabled = false;

            // Initialize global mouse hook
            _mouseHook = new Win32MouseHook();
            _mouseHook.LeftButtonDown += OnGlobalMouseDown;
            _mouseHook.LeftButtonUp += OnGlobalMouseUp;
            _mouseHook.MouseMoveFilter = OnMouseMoveFilter;
        }

        /// <summary>
        /// Gets or sets whether snap-to-objects is enabled.
        /// </summary>
        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                if (_isEnabled == value) return;
                _isEnabled = value;

                if (!_isEnabled)
                {
                    CancelDrag();
                    RemoveVisualFeedback();
                }
            }
        }

        /// <summary>
        /// Toggles snap-to-objects on or off.
        /// </summary>
        public void Toggle()
        {
            IsEnabled = !IsEnabled;
        }

        private void OnGlobalMouseDown(object sender, Win32MouseHook.MouseHookEventArgs e)
        {
            if (!_isEnabled) return;

            _mouseDownScreenX = e.ScreenX;
            _mouseDownScreenY = e.ScreenY;

            try
            {
                var selection = _powerPointManager.GetSelectedShapes();
                if (selection == null || selection.Count != 1)
                {
                    _dragReady = false;
                    return;
                }

                var shape = selection[1];

                // Convert screen coordinates to slide coordinates
                if (TryScreenToSlidePoint(e.ScreenX, e.ScreenY, out float slideX, out float slideY))
                {
                    // Check if mouse is over the selected shape
                    float shapeLeft = shape.Left;
                    float shapeRight = shapeLeft + shape.Width;
                    float shapeTop = shape.Top;
                    float shapeBottom = shapeTop + shape.Height;

                    if (slideX >= shapeLeft && slideX <= shapeRight &&
                        slideY >= shapeTop && slideY <= shapeBottom)
                    {
                        _draggedShape = shape;
                        _grabOffset = new Point(
                            (int)(slideX - shapeLeft),
                            (int)(slideY - shapeTop)
                        );
                        _originalLeft = shape.Left;
                        _originalTop = shape.Top;
                        _dragReady = true;
                        _dragActive = false;
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SnapToObjectsManager.OnGlobalMouseDown");
            }
        }

        private bool OnMouseMoveFilter(int screenX, int screenY)
        {
            if (!_isEnabled || !_dragReady || _isRedrawing)
            {
                return false;
            }

            // Check if drag threshold is met
            int deltaX = Math.Abs(screenX - _mouseDownScreenX);
            int deltaY = Math.Abs(screenY - _mouseDownScreenY);

            if (!_dragActive && deltaX < DragThresholdPixels && deltaY < DragThresholdPixels)
            {
                return false; // Don't suppress small movements
            }

            // Activate drag
            if (!_dragActive)
            {
                _dragActive = true;
            }

            // Store pending coordinates
            _pendingScreenX = screenX;
            _pendingScreenY = screenY;

            // Coalesce updates to prevent COM overload
            if (!_moveUpdatePending)
            {
                _moveUpdatePending = true;
                System.Windows.Application.Current?.Dispatcher.BeginInvoke(
                    new Action(ProcessPendingMove),
                    System.Windows.Threading.DispatcherPriority.Input
                );
            }

            return true; // Suppress mouse move
        }

        private bool _isRedrawing;

        private void ProcessPendingMove()
        {
            _moveUpdatePending = false;
            _isRedrawing = true;

            try
            {
                if (_draggedShape == null) return;

                // Convert screen to slide coordinates
                if (!TryScreenToSlidePoint(_pendingScreenX, _pendingScreenY, out float slideX, out float slideY))
                {
                    return;
                }

                // Calculate target position (unsnapped)
                float targetLeft = slideX - _grabOffset.X;
                float targetTop = slideY - _grabOffset.Y;

                // Find snap targets
                var snapResult = CalculateSnapPoint(
                    targetLeft,
                    targetTop,
                    _draggedShape.Width,
                    _draggedShape.Height
                );

                // Apply snapped position
                float finalLeft = snapResult.SnappedX;
                float finalTop = snapResult.SnappedY;

                _draggedShape.Left = finalLeft;
                _draggedShape.Top = finalTop;

                // Show visual feedback
                if (snapResult.SnapTarget != null)
                {
                    ShowVisualFeedback(
                        snapResult.SnapTarget,
                        snapResult.SnapType,
                        finalLeft,
                        finalTop,
                        _draggedShape.Width,
                        _draggedShape.Height
                    );
                }
                else
                {
                    RemoveVisualFeedback();
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SnapToObjectsManager.ProcessPendingMove");
            }
            finally
            {
                _isRedrawing = false;
            }
        }

        /// <summary>
        /// Calculates the best snap point for the given position.
        /// Returns the snapped coordinates and the snap target if found.
        /// </summary>
        private SnapResult CalculateSnapPoint(float targetLeft, float targetTop, float width, float height)
        {
            var app = _powerPointManager.GetApplication();
            if (app == null || app.ActivePresentation == null)
            {
                return new SnapResult(targetLeft, targetTop, null, SnapType.CenterX);
            }

            var slide = _powerPointManager.GetCurrentSlide();
            if (slide == null)
            {
                return new SnapResult(targetLeft, targetTop, null, SnapType.CenterX);
            }

            var snapPoints = new List<SnapPoint>();
            float bestSnapX = targetLeft;
            float bestSnapY = targetTop;
            Shape snapTarget = null;
            SnapType? snapType = null;
            float minDistance = SnapThresholdPixels;

            // Collect snap points from all other shapes on the slide
            try
            {
                foreach (Shape shape in slide.Shapes)
                {
                    try
                    {
                        // Skip the dragged shape and hidden shapes
                        if (shape == null || shape == _draggedShape || shape.Visible == Office.MsoTriState.msoFalse)
                            continue;

                        // Add edge and center snap points
                        snapPoints.Add(new SnapPoint(shape.Left, shape.Top + shape.Height / 2, shape, SnapType.LeftEdge));
                        snapPoints.Add(new SnapPoint(shape.Left + shape.Width, shape.Top + shape.Height / 2, shape, SnapType.RightEdge));
                        snapPoints.Add(new SnapPoint(shape.Left + shape.Width / 2, shape.Top + shape.Height / 2, shape, SnapType.CenterX));
                        snapPoints.Add(new SnapPoint(shape.Left + shape.Width / 2, shape.Top, shape, SnapType.TopEdge));
                        snapPoints.Add(new SnapPoint(shape.Left + shape.Width / 2, shape.Top + shape.Height, shape, SnapType.BottomEdge));
                        snapPoints.Add(new SnapPoint(shape.Left + shape.Width / 2, shape.Top + shape.Height / 2, shape, SnapType.CenterY));
                    }
                    catch (COMException) { /* Shape might have been deleted */ }
                }
            }
            catch (COMException) { /* Shape collection might have changed */ }

            // Calculate snap points for the dragged shape
            float draggedLeft = targetLeft;
            float draggedRight = targetLeft + width;
            float draggedCenterX = targetLeft + width / 2;
            float draggedTop = targetTop;
            float draggedBottom = targetTop + height;
            float draggedCenterY = targetTop + height / 2;

            // Check horizontal snaps (X-axis)
            foreach (var point in snapPoints)
            {
                float distance = float.MaxValue;
                float snappedX = draggedLeft;

                switch (point.Type)
                {
                    case SnapType.LeftEdge:
                    case SnapType.CenterX:
                    case SnapType.RightEdge:
                        // Try snapping left edge of dragged to this point
                        distance = Math.Abs(draggedLeft - point.X);
                        snappedX = point.X;

                        // Try snapping right edge of dragged to this point
                        float rightDist = Math.Abs(draggedRight - point.X);
                        if (rightDist < distance)
                        {
                            distance = rightDist;
                            snappedX = point.X - width;
                        }

                        // Try snapping center of dragged to this point
                        float centerDist = Math.Abs(draggedCenterX - point.X);
                        if (centerDist < distance)
                        {
                            distance = centerDist;
                            snappedX = point.X - width / 2;
                        }
                        break;
                }

                if (distance <= minDistance)
                {
                    minDistance = distance;
                    bestSnapX = snappedX;
                    snapTarget = point.Owner;
                    snapType = point.Type;
                }
            }

            // Check vertical snaps (Y-axis) with separate threshold check
            minDistance = SnapThresholdPixels;
            foreach (var point in snapPoints)
            {
                float distance = float.MaxValue;
                float snappedY = draggedTop;

                switch (point.Type)
                {
                    case SnapType.TopEdge:
                    case SnapType.CenterY:
                    case SnapType.BottomEdge:
                        // Try snapping top edge of dragged to this point
                        distance = Math.Abs(draggedTop - point.Y);
                        snappedY = point.Y;

                        // Try snapping bottom edge of dragged to this point
                        float bottomDist = Math.Abs(draggedBottom - point.Y);
                        if (bottomDist < distance)
                        {
                            distance = bottomDist;
                            snappedY = point.Y - height;
                        }

                        // Try snapping center of dragged to this point
                        float centerDist = Math.Abs(draggedCenterY - point.Y);
                        if (centerDist < distance)
                        {
                            distance = centerDist;
                            snappedY = point.Y - height / 2;
                        }
                        break;
                }

                if (distance <= minDistance)
                {
                    minDistance = distance;
                    bestSnapY = snappedY;
                    if (snapTarget == null || point.Owner != snapTarget)
                    {
                        snapTarget = point.Owner;
                    }
                    snapType = point.Type;
                }
            }

            return new SnapResult(bestSnapX, bestSnapY, snapTarget, snapType ?? SnapType.CenterX);
        }

        private void ShowVisualFeedback(Shape target, SnapType snapType, float shapeLeft, float shapeTop, float shapeWidth, float shapeHeight)
        {
            RemoveVisualFeedback();

            try
            {
                var slide = _powerPointManager.GetCurrentSlide();
                if (slide == null || target == null) return;

                // Test if shape still exists before accessing properties
                try
                {
                    string dummy = target.Name; 
                }
                catch (COMException)
                {
                    return; // Shape is gone
                }

                // Highlight the snap target shape using BuildFreeform with proper node additions
                float l = target.Left, t = target.Top, w = target.Width, h = target.Height;
                var ff = slide.Shapes.BuildFreeform(Office.MsoEditingType.msoEditingCorner, l, t);
                ff.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, l + w, t);
                ff.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, l + w, t + h);
                ff.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, l, t + h);
                ff.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, l, t);
                _highlightShape = ff.ConvertToShape();

                _highlightShape.Fill.Visible = Office.MsoTriState.msoFalse;
                _highlightShape.Line.Weight = 2;
                _highlightShape.Line.ForeColor.RGB = Color.FromArgb(0, 120, 215).ToArgb() & 0xFFFFFF;
                _highlightShape.Line.DashStyle = MsoLineDashStyle.msoLineDash;
                _highlightShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);

                // Draw snap guide lines
                var guides = new List<Shape>();

                // Horizontal guide
                float guideY;
                switch (snapType)
                {
                    case SnapType.TopEdge:
                        guideY = shapeTop;
                        break;
                    case SnapType.BottomEdge:
                        guideY = shapeTop + shapeHeight;
                        break;
                    case SnapType.CenterY:
                        guideY = shapeTop + shapeHeight / 2;
                        break;
                    default:
                        guideY = target.Top + target.Height / 2;
                        break;
                }

                var hGuide = slide.Shapes.AddLine(0, guideY,
                    _powerPointManager.GetApplication().ActivePresentation.PageSetup.SlideWidth, guideY);
                hGuide.Line.ForeColor.RGB = Color.FromArgb(0, 120, 215).ToArgb() & 0xFFFFFF;
                hGuide.Line.DashStyle = MsoLineDashStyle.msoLineDashDot;
                hGuide.Line.Weight = 1;
                guides.Add(hGuide);

                // Vertical guide
                float guideX;
                switch (snapType)
                {
                    case SnapType.LeftEdge:
                        guideX = shapeLeft;
                        break;
                    case SnapType.RightEdge:
                        guideX = shapeLeft + shapeWidth;
                        break;
                    case SnapType.CenterX:
                        guideX = shapeLeft + shapeWidth / 2;
                        break;
                    default:
                        guideX = target.Left + target.Width / 2;
                        break;
                }

                var vGuide = slide.Shapes.AddLine(guideX, 0, guideX,
                    _powerPointManager.GetApplication().ActivePresentation.PageSetup.SlideHeight);
                vGuide.Line.ForeColor.RGB = Color.FromArgb(0, 120, 215).ToArgb() & 0xFFFFFF;
                vGuide.Line.DashStyle = MsoLineDashStyle.msoLineDashDot;
                vGuide.Line.Weight = 1;
                guides.Add(vGuide);

                _snapGuideHorizontal = hGuide;
                _snapGuideVertical = vGuide;

                // Tag shapes for cleanup
                _highlightShape.Tags.Add("oPE_SnapHighlight", "true");
                _snapGuideHorizontal.Tags.Add("oPE_SnapGuide", "true");
                _snapGuideVertical.Tags.Add("oPE_SnapGuide", "true");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SnapToObjectsManager.ShowVisualFeedback");
            }
        }

        private void RemoveVisualFeedback()
        {
            try
            {
                if (_highlightShape != null)
                {
                    try { _highlightShape.Delete(); } catch { }
                    _highlightShape = null;
                }

                if (_snapGuideHorizontal != null)
                {
                    try { _snapGuideHorizontal.Delete(); } catch { }
                    _snapGuideHorizontal = null;
                }

                if (_snapGuideVertical != null)
                {
                    try { _snapGuideVertical.Delete(); } catch { }
                    _snapGuideVertical = null;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SnapToObjectsManager.RemoveVisualFeedback");
            }
        }

        private void OnGlobalMouseUp(object sender, EventArgs e)
        {
            if (_dragActive)
            {
                // Drag completed
                RemoveVisualFeedback();
            }
            CancelDrag();
        }

        private void CancelDrag()
        {
            _draggedShape = null;
            _dragReady = false;
            _dragActive = false;
            _moveUpdatePending = false;
        }

        /// <summary>
        /// Converts screen coordinates to slide coordinates.
        /// </summary>
        private bool TryScreenToSlidePoint(int screenX, int screenY, out float pointX, out float pointY)
        {
            pointX = 0;
            pointY = 0;

            try
            {
                var app = _powerPointManager.GetApplication();
                if (app == null || app.ActiveWindow == null || app.ActivePresentation == null)
                    return false;

                var window = app.ActiveWindow;
                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;

                int px0 = window.PointsToScreenPixelsX(0);
                int px1 = window.PointsToScreenPixelsX(slideWidth);
                int py0 = window.PointsToScreenPixelsY(0);
                int py1 = window.PointsToScreenPixelsY(slideHeight);

                float scaleX = (float)(slideWidth / Math.Max(1, (px1 - px0)));
                float scaleY = (float)(slideHeight / Math.Max(1, (py1 - py0)));

                pointX = (screenX - px0) * scaleX;
                pointY = (screenY - py0) * scaleY;

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SnapToObjectsManager.TryScreenToSlidePoint");
                return false;
            }
        }

        public void Dispose()
        {
            if (_disposed) return;

            _disposed = true;
            RemoveVisualFeedback();

            if (_mouseHook != null)
            {
                _mouseHook.LeftButtonDown -= OnGlobalMouseDown;
                _mouseHook.LeftButtonUp -= OnGlobalMouseUp;
                _mouseHook.MouseMoveFilter = null;
                _mouseHook.Dispose();
            }
        }

        private class SnapResult
        {
            public float SnappedX { get; }
            public float SnappedY { get; }
            public Shape SnapTarget { get; }
            public SnapType SnapType { get; }

            public SnapResult(float snappedX, float snappedY, Shape snapTarget, SnapType snapType)
            {
                SnappedX = snappedX;
                SnappedY = snappedY;
                SnapTarget = snapTarget;
                SnapType = snapType;
            }
        }
    }
}
