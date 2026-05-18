using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using oPenEfficiency.Features;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class StorylineDialog : Window
    {
        private readonly PowerPointManager _manager;

        public StorylineDialog(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            this.Topmost = true;
            LoadStoryline();
        }

        private void LoadStoryline()
        {
            try
            {
                var titles = StorylineFeature.ExtractStoryline(_manager);
                TitlesList.ItemsSource = titles;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading storyline: " + ex.Message);
            }
        }

        private void Item_Click(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is Button || e.OriginalSource is TextBox || IsAncestorOf<TextBox>(e.OriginalSource as DependencyObject))
                return;

            if (sender is FrameworkElement fe && fe.DataContext is StorylineFeature.StorylineInfo info)
            {
                StorylineFeature.GoToSlide(_manager, info.SlideNumber);
            }
        }

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.TemplatedParent == null) // Manual check for button inside template
            {
                var parent = VisualTreeHelper.GetParent(btn);
                while (parent != null && !(parent is Grid)) parent = VisualTreeHelper.GetParent(parent);
                
                if (parent is Grid grid)
                {
                    var txtRead = grid.FindName("TxtRead") as UIElement;
                    var txtEdit = grid.FindName("TxtEdit") as TextBox;
                    
                    if (txtRead != null && txtEdit != null)
                    {
                        txtRead.Visibility = Visibility.Collapsed;
                        txtEdit.Visibility = Visibility.Visible;
                        txtEdit.Focus();
                        txtEdit.SelectAll();
                    }
                }
            }
        }

        private void TxtEdit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                CommitEdit(sender as TextBox);
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                CancelEdit(sender as TextBox);
                e.Handled = true;
            }
        }

        private void TxtEdit_LostFocus(object sender, RoutedEventArgs e)
        {
            CommitEdit(sender as TextBox);
        }

        private void CommitEdit(TextBox tb)
        {
            if (tb == null || tb.Visibility != Visibility.Visible) return;

            if (tb.DataContext is StorylineFeature.StorylineInfo info)
            {
                string newTitle = tb.Text.Trim();
                if (newTitle != info.TitleText)
                {
                    if (StorylineFeature.UpdateTitle(_manager, info.SlideNumber, newTitle))
                    {
                        info.TitleText = newTitle;
                    }
                }
            }

            var grid = VisualTreeHelper.GetParent(tb) as Grid;
            if (grid != null)
            {
                var txtRead = grid.FindName("TxtRead") as UIElement;
                if (txtRead != null) txtRead.Visibility = Visibility.Visible;
                tb.Visibility = Visibility.Collapsed;
                
                // Refresh list display binding
                var temp = TitlesList.ItemsSource;
                TitlesList.ItemsSource = null;
                TitlesList.ItemsSource = temp;
            }
        }

        private void CancelEdit(TextBox tb)
        {
            if (tb == null) return;
            
            if (tb.DataContext is StorylineFeature.StorylineInfo info)
            {
                tb.Text = info.TitleText;
            }

            var grid = VisualTreeHelper.GetParent(tb) as Grid;
            if (grid != null)
            {
                var txtRead = grid.FindName("TxtRead") as UIElement;
                if (txtRead != null) txtRead.Visibility = Visibility.Visible;
                tb.Visibility = Visibility.Collapsed;
            }
        }

        private bool IsAncestorOf<T>(DependencyObject node) where T : DependencyObject
        {
            while (node != null)
            {
                if (node is T) return true;
                node = VisualTreeHelper.GetParent(node);
            }
            return false;
        }

        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var titles = (IEnumerable<StorylineFeature.StorylineInfo>)TitlesList.ItemsSource;
                var sb = new System.Text.StringBuilder();
                foreach (var item in titles)
                {
                    sb.AppendLine($"{item.SlideNumber}: {item.TitleText}");
                }
                Clipboard.SetText(sb.ToString());
                MessageBox.Show("Storyline copied to clipboard.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error copying to clipboard: " + ex.Message);
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && !(e.OriginalSource is TextBox))
                this.DragMove();
        }
    }
}
