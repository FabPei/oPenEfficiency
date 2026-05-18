using System;
using System.Drawing;
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
    /// Manages snap-to-grid functionality for shape dragging.
    /// When enabled, moving shapes will snap to grid lines.
    /// </summary>
    public class SnapToGridManager : IDisposable
    {
        // Snap threshold in pixels - shapes snap when within this distance
        private const int SnapThresholdPixels = 10;

        // Minimum drag distance before snap activates (prevents accidental triggers)
        private const int DragThresholdPixels = 4;

        // Grid spacing in points (PowerPoint units) - default 50 points
        private float _gridSpacingX = 50f;
        private float _gridSpacingY = 50f;

        // Grid offset (origin point)
        private float _gridOffsetX = 0f;
        private float _gridOffsetY = 0f;

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
        private Shape _snapGuideHorizontal;
        private Shape _snapGuideVertical;
        private bool _isRedrawing;

        public SnapToGridManager(PowerPointManager powerPointManager)
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
        /// Gets or sets whether snap-to-grid is enabled.
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
        /// Gets or sets the grid spacing in points (PowerPoint units).
        /// </summary>
        public float GridSpacingX
        {
            get => _gridSpacingX;
            set => _gridSpacingX = Math.Max(10f, value);
        }

        public float GridSpacingY
        {
            get => _gridSpacingY;
            set => _gridSpacingY = Math.Max(10f, value);
        }

        /// <summary>
        /// Toggles snap-to-grid on or off.
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
                ExceptionLogger.Log(ex, "SnapToGridManager.OnGlobalMouseDown");
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

                // Calculate snapped position
                var snapResult = CalculateSnapPoint(targetLeft, targetTop);

                // Apply snapped position
                _draggedShape.Left = snapResult.SnappedX;
                _draggedShape.Top = snapResult.SnappedY;

                // Show visual feedback
                ShowVisualFeedback(snapResult.SnappedX, snapResult.SnappedY, _draggedShape.Width, _draggedShape.Height);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SnapToGridManager.ProcessPendingMove");
            }
            finally
            {
                _isRedrawing = false;
            }
        }

        /// <summary>
        /// Calculates the snap point based on grid lines.
        /// </summary>
        private SnapResult CalculateSnapPoint(float targetLeft, float targetTop)
        {
            // Snap to nearest grid line
            float snappedX = SnapToGrid(targetLeft + _gridOffsetX, _gridSpacingX) - _gridOffsetX;
            float snappedY = SnapToGrid(targetTop + _gridOffsetY, _gridSpacingY) - _gridOffsetY;

            return new SnapResult(snappedX, snappedY);
        }

        /// <summary>
        /// Snaps a value to the nearest grid line.
        /// </summary>
        private float SnapToGrid(float value, float spacing)
        {
            return (float)Math.Round(value / spacing) * spacing;
        }

        private void ShowVisualFeedback(float shapeLeft, float shapeTop, float shapeWidth, float shapeHeight)
        {
            RemoveVisualFeedback();

            try
            {
                var slide = _powerPointManager.GetCurrentSlide();
                if (slide == null) return;

                var app = _powerPointManager.GetApplication();
                if (app == null) return;

                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;

                // Draw vertical guide line at snapped X position
                float guideX = shapeLeft + shapeWidth / 2; // Center of shape
                var vGuide = slide.Shapes.AddLine(guideX, 0, guideX, slideHeight);
                vGuide.Line.ForeColor.RGB = Color.FromArgb(0, 120, 215).ToArgb() & 0xFFFFFF;
                vGuide.Line.DashStyle = MsoLineDashStyle.msoLineDashDot;
                vGuide.Line.Weight = 1;
                _snapGuideVertical = vGuide;
                _snapGuideVertical.Tags.Add("oPE_SnapGuide", "true");

                // Draw horizontal guide line at snapped Y position
                float guideY = shapeTop + shapeHeight / 2; // Center of shape
                var hGuide = slide.Shapes.AddLine(0, guideY, slideWidth, guideY);
                hGuide.Line.ForeColor.RGB = Color.FromArgb(0, 120, 215).ToArgb() & 0xFFFFFF;
                hGuide.Line.DashStyle = MsoLineDashStyle.msoLineDashDot;
                hGuide.Line.Weight = 1;
                _snapGuideHorizontal = hGuide;
                _snapGuideHorizontal.Tags.Add("oPE_SnapGuide", "true");

                // Draw grid intersection point
                var intersection = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeOval,
                    guideX - 4,
                    guideY - 4,
                    8,
                    8
                );
                intersection.Fill.ForeColor.RGB = Color.FromArgb(0, 120, 215).ToArgb() & 0xFFFFFF;
                intersection.Line.Visible = Office.MsoTriState.msoFalse;
                intersection.Tags.Add("oPE_SnapGuide", "true");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SnapToGridManager.ShowVisualFeedback");
            }
        }

        private void RemoveVisualFeedback()
        {
            try
            {
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
                ExceptionLogger.Log(ex, "SnapToGridManager.RemoveVisualFeedback");
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
                ExceptionLogger.Log(ex, "SnapToGridManager.TryScreenToSlidePoint");
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

            public SnapResult(float snappedX, float snappedY)
            {
                SnappedX = snappedX;
                SnappedY = snappedY;
            }
        }
    }
}
