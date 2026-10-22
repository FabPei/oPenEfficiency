using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Features.Utilities;

namespace oPenEfficiency.UI.Toolbars
{
    public partial class SlideNotesWindow : Window
    {
        private readonly PowerPointManager _manager;
        private Microsoft.Office.Interop.PowerPoint.Application _app;
        private ToggleButton _toggleBtn;
        private DispatcherTimer _debounceTimer;
        private Slide _currentSlide;
        private Presentation _currentPresentation;
        private bool _isUpdating = false;

        private PrivateComment _currentComment;
        private List<PrivateComment> _comments = new List<PrivateComment>();

        public SlideNotesWindow(PowerPointManager manager, ToggleButton toggleBtn)
        {
            InitializeComponent();
            _manager = manager;
            _toggleBtn = toggleBtn;
            _app = _manager.GetApplication();

            if (_app != null)
            {
                try { _app.WindowSelectionChange += App_WindowSelectionChange; } catch { /* Notes remain usable for the initial slide when event hookup is unavailable. */ }
            }

            _debounceTimer = new DispatcherTimer();
            _debounceTimer.Interval = TimeSpan.FromMilliseconds(500);
            _debounceTimer.Tick += DebounceTimer_Tick;

            this.Closed += SlideNotesWindow_Closed;

            LoadNotesForCurrentSlide();
        }

        private void App_WindowSelectionChange(Selection Sel)
        {
            // Dispatch to UI thread to update TextBox safely
            this.Dispatcher.BeginInvoke(new Action(() =>
            {
                LoadNotesForCurrentSlide();
            }));
        }

        private void LoadNotesForCurrentSlide()
        {
            if (_app == null || _app.ActiveWindow == null) return;
            
            try
            {
                var view = _app.ActiveWindow.View;
                if (view != null && view.Slide != null)
                {
                    var slide = view.Slide as Slide;
                    var presentation = _app.ActivePresentation;
                    if (slide != null)
                    {
                        if (_currentSlide != null && ReferenceEquals(_currentPresentation, presentation) && _currentSlide.SlideID == slide.SlideID) return;

                        // Before switching slides, save any pending changes
                        if (_debounceTimer.IsEnabled)
                        {
                            _debounceTimer.Stop();
                            SaveNotesToCurrentSlide();
                        }

                        _currentSlide = slide;
                        _currentPresentation = presentation;
                        RefreshComments();
                    }
                }
            }
            catch (Exception ex) { Utils.ExceptionLogger.Log(ex, "SlideNotesWindow.LoadNotesForCurrentSlide"); }
        }

        private void NotesTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdating) return;

            // Restart debounce timer
            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        private void CommentListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdating) return;
            SaveNotesToCurrentSlide();
            LoadSelectedComment(CommentListBox.SelectedItem as PrivateComment);
        }

        private void StatusComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdating || _currentComment == null) return;
            _currentComment.Status = StatusComboBox.SelectedIndex == 1 ? PrivateCommentStatus.Done : PrivateCommentStatus.Open;
            SaveNotesToCurrentSlide();
            RefreshComments(_currentComment.Id);
        }

        private void AddCommentButton_Click(object sender, RoutedEventArgs e)
        {
            SaveNotesToCurrentSlide();
            var comment = PrivateCommentsService.CreateComment(_currentSlide, GetSelectedShape());
            if (comment == null) return;
            _currentComment = comment;
            PrivateCommentsService.SaveForSlide(_currentPresentation, _currentSlide, comment);
            RefreshComments(comment.Id);
            NotesTextBox.Focus();
        }

        private void DeleteCommentButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentComment == null) return;
            PrivateCommentsService.DeleteComment(_currentPresentation, _currentComment.Id);
            RefreshComments();
        }

        private void AttachButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentComment == null) return;
            var shape = GetSelectedShape();
            if (shape == null) return;
            _currentComment.AnchorShapeId = shape.Id;
            _currentComment.AnchorShapeName = shape.Name;
            SaveNotesToCurrentSlide();
            RefreshComments(_currentComment.Id);
        }

        private void ShowOverlayButton_Click(object sender, RoutedEventArgs e)
        {
            SlideNotesFeature.ToggleOverlay(_manager);
        }

        private void DebounceTimer_Tick(object sender, EventArgs e)
        {
            _debounceTimer.Stop();
            SaveNotesToCurrentSlide();
        }

        private void SaveNotesToCurrentSlide()
        {
            if (_currentSlide == null || _currentPresentation == null) return;

            try
            {
                if (_currentComment == null && string.IsNullOrWhiteSpace(NotesTextBox.Text)) return;
                if (_currentComment == null)
                {
                    _currentComment = PrivateCommentsService.CreateComment(_currentSlide, GetSelectedShape());
                }
                _currentComment.Text = NotesTextBox.Text;
                _currentComment.Status = StatusComboBox.SelectedIndex == 1 ? PrivateCommentStatus.Done : PrivateCommentStatus.Open;
                PrivateCommentsService.SaveForSlide(_currentPresentation, _currentSlide, _currentComment);
                MetadataText.Text = FormatMetadata(_currentComment);
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "SlideNotesWindow.SaveNotesToCurrentSlide");
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        public void ClearNotesIfSlideMatches(int slideId)
        {
            if (_currentSlide != null && _currentSlide.SlideID == slideId)
            {
                _isUpdating = true;
                NotesTextBox.Text = "";
                _currentComment = null;
                _isUpdating = false;
                _debounceTimer.Stop(); // Don't re-save empty back just in case
            }
        }
        
        public void ForceRefresh()
        {
            _currentSlide = null; // force reload
            _currentPresentation = null;
            LoadNotesForCurrentSlide();
        }

        private void SlideNotesWindow_Closed(object sender, EventArgs e)
        {
            if (_app != null)
            {
                try { _app.WindowSelectionChange -= App_WindowSelectionChange; } catch { /* Shutdown continues when PowerPoint already disconnected the event source. */ }
            }

            _debounceTimer.Stop();
            SaveNotesToCurrentSlide(); // Final save before closing

            if (_toggleBtn != null) _toggleBtn.IsChecked = false;
            SlideNotesFeature.SetWindowNull();
        }

        private void RefreshComments(string selectedId = null)
        {
            if (_currentSlide == null || _currentPresentation == null) return;
            _comments = PrivateCommentsService.GetCommentsForSlide(_currentPresentation, _currentSlide);
            CommentListHeader.Text = $"COMMENTS ON THIS SLIDE · {_comments.Count(comment => !comment.IsDone)} OPEN";
            _isUpdating = true;
            CommentListBox.ItemsSource = _comments;
            var selected = _comments.FirstOrDefault(comment => comment.Id == selectedId) ?? _comments.FirstOrDefault();
            CommentListBox.SelectedItem = selected;
            _isUpdating = false;
            LoadSelectedComment(selected);
        }

        private void LoadSelectedComment(PrivateComment comment)
        {
            _isUpdating = true;
            _currentComment = comment;
            NotesTextBox.Text = comment?.Text ?? string.Empty;
            StatusComboBox.SelectedIndex = comment?.IsDone == true ? 1 : 0;
            MetadataText.Text = comment == null ? "Select + Add to create a comment." : FormatMetadata(comment);
            _isUpdating = false;
        }

        private static string FormatMetadata(PrivateComment comment)
        {
            string anchor = string.IsNullOrWhiteSpace(comment.AnchorShapeName) ? string.Empty : $" · Attached: {comment.AnchorShapeName}";
            return $"{comment.Author} · {comment.UpdatedAtUtc.ToLocalTime():g}{anchor}";
        }

        private Shape GetSelectedShape()
        {
            try
            {
                var selection = _app.ActiveWindow.Selection;
                return selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 0
                    ? selection.ShapeRange[1] : null;
            }
            catch { return null; }
        }
    }
}
