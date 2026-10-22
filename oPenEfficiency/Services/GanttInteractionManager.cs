using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Point = System.Drawing.Point;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Manages specific interactive behaviors for Gantt chart shapes.
    /// Handles constraints like horizontal-only movement for event lines.
    /// </summary>
    public class GanttInteractionManager : IDisposable
    {
        private const int DragThresholdPixels = 4;

        private readonly PowerPointManager _powerPointManager;

        // Drag state
        private Shape _draggedShape;
        private Point _grabOffset;
        private float _originalTop;
        private int _mouseDownScreenX;
        private int _mouseDownScreenY;
        private bool _dragReady;
        private bool _dragActive;

        // Coalescing state
        private int _pendingScreenX;
        private int _pendingScreenY;
        private bool _moveUpdatePending;
        private bool _isRedrawing;

        public GanttInteractionManager(PowerPointManager powerPointManager)
        {
            _powerPointManager = powerPointManager;
        }

        public void OnMouseDown(int screenX, int screenY)
        {
            CancelDrag();
            _mouseDownScreenX = screenX;
            _mouseDownScreenY = screenY;

            try
            {
                var selection = _powerPointManager.GetSelectedShapes();
                if (selection == null || selection.Count != 1)
                {
                    _dragReady = false;
                    return;
                }

                Shape shape = selection[1];
                bool retainShape = false;
                try
                {
                    // Check if this is an Event Line
                    if (!IsEventLine(shape))
                    {
                        _dragReady = false;
                        return;
                    }

                    // Convert screen coordinates to slide coordinates
                    if (TryScreenToSlidePoint(screenX, screenY, out float slideX, out float slideY))
                    {
                        // Check if mouse is over the shape bounds (loose check for lines)
                        float threshold = 10f;
                        float shapeLeft = shape.Left - threshold;
                        float shapeRight = shape.Left + shape.Width + threshold;
                        float shapeTop = shape.Top;
                        float shapeBottom = shapeTop + shape.Height;

                        if (slideX >= shapeLeft && slideX <= shapeRight &&
                            slideY >= shapeTop && slideY <= shapeBottom)
                        {
                            _draggedShape = shape;
                            retainShape = true;
                            _grabOffset = new Point(
                                (int)(slideX - shape.Left),
                                (int)(slideY - shape.Top)
                            );
                            _originalTop = shape.Top;
                            _dragReady = true;
                            _dragActive = false;
                        }
                    }
                }
                finally
                {
                    if (!retainShape && shape != null && Marshal.IsComObject(shape)) Marshal.ReleaseComObject(shape);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttInteractionManager.OnMouseDown");
            }
        }

        public bool OnMouseMove(int screenX, int screenY)
        {
            if (!_dragReady || _isRedrawing) return false;

            // Check if drag threshold is met
            int deltaX = Math.Abs(screenX - _mouseDownScreenX);
            int deltaY = Math.Abs(screenY - _mouseDownScreenY);

            if (!_dragActive && deltaX < DragThresholdPixels && deltaY < DragThresholdPixels)
            {
                return false;
            }

            if (!_dragActive) _dragActive = true;

            _pendingScreenX = screenX;
            _pendingScreenY = screenY;

            if (!_moveUpdatePending)
            {
                _moveUpdatePending = true;
                System.Windows.Application.Current?.Dispatcher.BeginInvoke(
                    new Action(ProcessPendingMove),
                    System.Windows.Threading.DispatcherPriority.Input
                );
            }

            return true; // Suppress native mouse move to control it ourselves
        }

        public void OnMouseUp()
        {
            CancelDrag();
        }

        public void Dispose()
        {
            CancelDrag();
            GC.SuppressFinalize(this);
        }

        private void ProcessPendingMove()
        {
            _moveUpdatePending = false;
            _isRedrawing = true;

            try
            {
                // Deferred drag work must never run after the button was released.
                if (_draggedShape == null) return;
                if (!Win32MouseHook.IsLeftMouseButtonDown())
                {
                    CancelDrag();
                    return;
                }

                if (!TryScreenToSlidePoint(_pendingScreenX, _pendingScreenY, out float slideX, out float slideY))
                {
                    return;
                }

                // Calculate target X position
                float targetLeft = slideX - _grabOffset.X;

                // Update shape position: MOVE HORIZONTALLY ONLY
                _draggedShape.Left = targetLeft;
                _draggedShape.Top = _originalTop; // Restore vertical position
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttInteractionManager.ProcessPendingMove");
                CancelDrag();
            }
            finally
            {
                _isRedrawing = false;
            }
        }

        private void CancelDrag()
        {
            if (_draggedShape != null && Marshal.IsComObject(_draggedShape))
            {
                try { Marshal.ReleaseComObject(_draggedShape); }
                catch (COMException)
                {
                    // The dragged shape was already released by PowerPoint.
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "GanttInteractionManager.CancelDrag.ReleaseShape");
                }
            }
            _draggedShape = null;
            _dragReady = false;
            _dragActive = false;
            _moveUpdatePending = false;
        }

        private bool IsEventLine(Shape shape)
        {
            try
            {
                string typeTag = PowerPointShapeManager.GetGanttType(shape);
                return typeTag == PowerPointShapeManager.GanttShapeType.EventLine.ToString();
            }
            catch
            {
                return false;
            }
        }

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
            catch
            {
                return false;
            }
        }
    }
}
