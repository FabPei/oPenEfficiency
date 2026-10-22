using System;
using System.Drawing;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using oPenEfficiency.Services;
using oPenEfficiency.UI.Panels;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI.Toolbars
{
    /// <summary>
    /// Displays non-persistent resize handles over a rendered Gantt task bar.
    /// The overlay never becomes part of the presentation or its printed output.
    /// </summary>
    public partial class GanttResizeHandlesWindow : Window
    {
        private const double HandleWidthDip = 10;
        private const double VerticalPaddingDip = 3;

        private readonly PowerPointManager _manager;
        private readonly GanttPanelViewModel _viewModel;
        private readonly PowerPoint.Shape _taskBar;
        private readonly Guid _taskId;

        private bool _draggingStart;
        private double _accumulatedDragDip;
        private float _originalLeft;
        private float _originalWidth;
        private float _minimumWidth;
        private double _dipToPoints = 1;

        public GanttResizeHandlesWindow(
            PowerPointManager manager,
            PowerPoint.Shape taskBar,
            GanttPanelViewModel viewModel,
            Guid taskId)
        {
            _manager = manager ?? throw new ArgumentNullException(nameof(manager));
            _taskBar = taskBar ?? throw new ArgumentNullException(nameof(taskBar));
            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            _taskId = taskId;

            InitializeComponent();

            var app = _manager.GetApplication();
            if (app != null)
            {
                new WindowInteropHelper(this).Owner = (IntPtr)app.HWND;
            }

            Loaded += (s, e) => UpdateOverlayBounds();
        }

        public bool MatchesTask(Guid taskId)
        {
            return taskId == _taskId;
        }

        public void RefreshPosition()
        {
            if (!IsLoaded) return;
            UpdateOverlayBounds();
        }

        public bool IsPointerOverHandle(int screenX, int screenY)
        {
            try
            {
                var app = _manager.GetApplication();
                if (app == null) return false;

                double dpiX;
                double dpiY;
                using (Graphics graphics = Graphics.FromHwnd((IntPtr)app.HWND))
                {
                    dpiX = graphics.DpiX;
                    dpiY = graphics.DpiY;
                }

                double xDip = screenX * 96.0 / dpiX;
                double yDip = screenY * 96.0 / dpiY;
                bool insideY = yDip >= Top && yDip <= Top + Height;
                bool insideLeft = xDip >= Left && xDip <= Left + HandleWidthDip;
                bool insideRight = xDip >= Left + Width - HandleWidthDip && xDip <= Left + Width;
                return insideY && (insideLeft || insideRight);
            }
            catch
            {
                return false;
            }
        }

        private void Handle_DragStarted(object sender, DragStartedEventArgs e)
        {
            try
            {
                _draggingStart = ReferenceEquals(sender, LeftHandle);
                _accumulatedDragDip = 0;
                _originalLeft = _taskBar.Left;
                _originalWidth = _taskBar.Width;
                _minimumWidth = CalculateMinimumWidth();
                _dipToPoints = CalculateDipToPoints();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttResizeHandlesWindow.Handle_DragStarted");
            }
        }

        private void Handle_DragDelta(object sender, DragDeltaEventArgs e)
        {
            try
            {
                _accumulatedDragDip += e.HorizontalChange;
                float deltaPoints = (float)(_accumulatedDragDip * _dipToPoints);

                if (_draggingStart)
                {
                    float right = _originalLeft + _originalWidth;
                    float newLeft = Math.Min(_originalLeft + deltaPoints, right - _minimumWidth);
                    _taskBar.Left = newLeft;
                    _taskBar.Width = right - newLeft;
                }
                else
                {
                    _taskBar.Width = Math.Max(_minimumWidth, _originalWidth + deltaPoints);
                }

                UpdateOverlayBounds();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttResizeHandlesWindow.Handle_DragDelta");
                Close();
            }
        }

        private void Handle_DragCompleted(object sender, DragCompletedEventArgs e)
        {
            try
            {
                float left = _taskBar.Left;
                float width = _taskBar.Width;
                Close();
                _viewModel.ApplyTaskBarResize(_taskId, left, width);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttResizeHandlesWindow.Handle_DragCompleted");
                try { Close(); }
                catch (InvalidOperationException)
                {
                    // The overlay is already closing after the drag failure.
                }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(closeEx, "GanttResizeHandlesWindow.Handle_DragCompleted.Close");
                }
            }
        }

        private float CalculateMinimumWidth()
        {
            var task = _viewModel.Project.Tasks.FirstOrDefault(item => item.Id == _taskId);
            if (task == null) return 4f;

            double durationDays = Math.Max(1, (task.End - task.Start).TotalDays);
            return Math.Max(4f, (float)(_originalWidth / durationDays));
        }

        private double CalculateDipToPoints()
        {
            var app = _manager.GetApplication();
            var window = app?.ActiveWindow;
            if (window == null) return 1;

            int px0 = window.PointsToScreenPixelsX(0);
            int px100 = window.PointsToScreenPixelsX(100);
            double pixelsPerPoint = Math.Abs(px100 - px0) / 100.0;
            if (pixelsPerPoint <= 0) return 1;

            using (Graphics graphics = Graphics.FromHwnd((IntPtr)app.HWND))
            {
                double pixelsPerDip = graphics.DpiX / 96.0;
                return pixelsPerDip / pixelsPerPoint;
            }
        }

        private void UpdateOverlayBounds()
        {
            try
            {
                var app = _manager.GetApplication();
                var window = app?.ActiveWindow;
                if (window == null) return;

                int leftPx = window.PointsToScreenPixelsX(_taskBar.Left);
                int rightPx = window.PointsToScreenPixelsX(_taskBar.Left + _taskBar.Width);
                int topPx = window.PointsToScreenPixelsY(_taskBar.Top);
                int bottomPx = window.PointsToScreenPixelsY(_taskBar.Top + _taskBar.Height);

                double dpiX;
                double dpiY;
                using (Graphics graphics = Graphics.FromHwnd((IntPtr)app.HWND))
                {
                    dpiX = graphics.DpiX;
                    dpiY = graphics.DpiY;
                }

                double leftDip = Math.Min(leftPx, rightPx) * 96.0 / dpiX;
                double rightDip = Math.Max(leftPx, rightPx) * 96.0 / dpiX;
                double topDip = Math.Min(topPx, bottomPx) * 96.0 / dpiY;
                double bottomDip = Math.Max(topPx, bottomPx) * 96.0 / dpiY;

                Left = leftDip - (HandleWidthDip / 2);
                Top = topDip - VerticalPaddingDip;
                Width = Math.Max(HandleWidthDip * 2, rightDip - leftDip + HandleWidthDip);
                Height = Math.Max(20, bottomDip - topDip + (VerticalPaddingDip * 2));

                LeftHandle.Height = Height;
                RightHandle.Height = Height;
                Canvas.SetLeft(RightHandle, Width - HandleWidthDip);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttResizeHandlesWindow.UpdateOverlayBounds");
                try { Close(); }
                catch (InvalidOperationException)
                {
                    // The overlay is already closing after the positioning failure.
                }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(closeEx, "GanttResizeHandlesWindow.UpdateOverlayBounds.Close");
                }
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            // This overlay's drag gesture belongs exclusively to its resize thumbs.
            e.Handled = true;
        }
    }
}
