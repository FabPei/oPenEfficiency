using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Features.Utilities;
using oPenEfficiency.Services;
using PowerPointApplication = Microsoft.Office.Interop.PowerPoint.Application;

namespace oPenEfficiency.UI.Toolbars
{
    /// <summary>Non-persistent display layer for the current slide's private review note.</summary>
    public partial class PrivateCommentOverlayWindow : Window
    {
        private readonly PowerPointManager _manager;
        private readonly PowerPointApplication _application;
        private PrivateComment _displayedComment;

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

        public PrivateCommentOverlayWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            _application = manager.GetApplication();
            try { _application.WindowSelectionChange += Application_WindowSelectionChange; } catch { /* Overlay remains usable for its initial shape when event hookup is unavailable. */ }
            Closed += PrivateCommentOverlayWindow_Closed;
            RefreshForCurrentSlide();
            PositionNearPowerPoint();
        }

        private void Application_WindowSelectionChange(Selection selection)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                RefreshForCurrentSlide();
                PositionNearPowerPoint();
            }));
        }

        private void RefreshForCurrentSlide()
        {
            try
            {
                var slide = _application?.ActiveWindow?.View?.Slide as Slide;
                var comments = slide == null ? new System.Collections.Generic.List<PrivateComment>() : PrivateCommentsService.GetCommentsForSlide(_application.ActivePresentation, slide);
                var comment = comments.FirstOrDefault(item => !item.IsDone) ?? comments.FirstOrDefault();
                _displayedComment = comment;
                if (comment == null)
                {
                    MetadataText.Visibility = Visibility.Collapsed;
                    DoneButton.Visibility = Visibility.Collapsed;
                    CommentText.Text = "No private review note on this slide.";
                    return;
                }

                MetadataText.Visibility = Visibility.Visible;
                MetadataText.Text = $"{comment.Author} · {comment.UpdatedAtUtc.ToLocalTime():g}" +
                    (comment.IsDone ? " · Done" : $" · Open · {comments.Count(item => !item.IsDone)} open");
                DoneButton.Visibility = comment.IsDone ? Visibility.Collapsed : Visibility.Visible;
                CommentText.Text = comment.Text;
            }
            catch
            {
                MetadataText.Visibility = Visibility.Collapsed;
                MetadataText.Text = "Private Review Note";
                CommentText.Text = "The comment could not be loaded.";
            }
        }

        private void PositionNearPowerPoint()
        {
            try
            {
                var hwnd = new IntPtr(_application.ActiveWindow.HWND);
                if (!GetWindowRect(hwnd, out var rect)) return;
                using (var graphics = Graphics.FromHwnd(hwnd))
                {
                    double scaleX = 96.0 / graphics.DpiX;
                    double scaleY = 96.0 / graphics.DpiY;
                    Left = (rect.Right * scaleX) - Width - 24;
                    Top = (rect.Top * scaleY) + 90;
                }
            }
            catch { /* A stale selection temporarily hides comment details until the next event. */ }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();

        private void DoneButton_Click(object sender, RoutedEventArgs e)
        {
            var slide = _application?.ActiveWindow?.View?.Slide as Slide;
            if (slide == null || _displayedComment == null) return;
            _displayedComment.Status = PrivateCommentStatus.Done;
            PrivateCommentsService.SaveForSlide(_application.ActivePresentation, slide, _displayedComment);
            RefreshForCurrentSlide();
        }

        private void PrivateCommentOverlayWindow_Closed(object sender, EventArgs e)
        {
            try { _application.WindowSelectionChange -= Application_WindowSelectionChange; } catch { /* Shutdown continues when PowerPoint already disconnected the event source. */ }
            SlideNotesFeature.SetOverlayNull();
        }
    }
}
