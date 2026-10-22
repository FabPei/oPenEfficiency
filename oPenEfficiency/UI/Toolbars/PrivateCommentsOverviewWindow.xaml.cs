using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI.Toolbars
{
    public partial class PrivateCommentsOverviewWindow : Window
    {
        private readonly PowerPointManager _manager;

        public PrivateCommentsOverviewWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            Refresh();
        }

        private void Refresh()
        {
            var app = _manager.GetApplication();
            var comments = PrivateCommentsService.GetAllComments(app.ActivePresentation);
            if (OpenOnlyCheckBox.IsChecked == true) comments = comments.Where(comment => !comment.IsDone).ToList();
            CommentListBox.ItemsSource = comments.Select(comment => new OverviewItem
            {
                Comment = comment,
                Header = $"Slide {GetSlideNumber(app.ActivePresentation, comment.SlideId)} · {(comment.IsDone ? "Done" : "Open")} · {comment.Author}",
                Text = comment.Text
            }).ToList();
        }

        private static int GetSlideNumber(Presentation presentation, int slideId)
        {
            foreach (Slide slide in presentation.Slides) if (slide.SlideID == slideId) return slide.SlideIndex;
            return 0;
        }

        private void FilterChanged(object sender, RoutedEventArgs e) => Refresh();

        private void CommentListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = CommentListBox.SelectedItem as OverviewItem;
            if (item == null) return;
            var app = _manager.GetApplication();
            foreach (Slide slide in app.ActivePresentation.Slides)
            {
                if (slide.SlideID == item.Comment.SlideId)
                {
                    app.ActiveWindow.View.GotoSlide(slide.SlideIndex);
                    return;
                }
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e) { if (e.ChangedButton == MouseButton.Left) DragMove(); }
        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();

        private class OverviewItem
        {
            public PrivateComment Comment { get; set; }
            public string Header { get; set; }
            public string Text { get; set; }
        }
    }
}
