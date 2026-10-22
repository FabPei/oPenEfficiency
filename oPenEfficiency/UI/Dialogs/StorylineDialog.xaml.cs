using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Win32;
using oPenEfficiency.Features;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class StorylineDialog : Window
    {
        private readonly PowerPointManager _manager;
        private List<StorylineFeature.StorylineInfo> _titles = new List<StorylineFeature.StorylineInfo>();
        private ICollectionView _view;
        private string _currentSection;
        private bool _initializing;
        private readonly DispatcherTimer _refreshTimer;

        public StorylineDialog(PowerPointManager manager, bool currentSectionOnly = false, bool issuesOnly = false)
        {
            InitializeComponent();
            _manager = manager;
            SetInitialView(currentSectionOnly, issuesOnly);
            RefreshStoryline();
            _refreshTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(5) };
            _refreshTimer.Tick += AutoRefreshTimer_Tick;
            _refreshTimer.Start();
            Closed += (s, e) => _refreshTimer.Stop();
        }

        public void SetInitialView(bool currentSectionOnly, bool issuesOnly)
        {
            _initializing = true;
            if (ScopeCombo != null) ScopeCombo.SelectedIndex = currentSectionOnly ? 1 : 0;
            if (FilterCombo != null) FilterCombo.SelectedIndex = issuesOnly ? 1 : 0;
            _initializing = false;
            _view?.Refresh();
            UpdateSummary();
        }

        public void SetFilter(string filter)
        {
            if (FilterCombo == null) return;
            _initializing = true;
            foreach (ComboBoxItem item in FilterCombo.Items)
            {
                if (string.Equals(item.Tag?.ToString(), filter, StringComparison.OrdinalIgnoreCase))
                {
                    FilterCombo.SelectedItem = item;
                    break;
                }
            }
            _initializing = false;
            _view?.Refresh();
            UpdateSummary();
        }

        public void RefreshStoryline()
        {
            try
            {
                int activeSlideId = GetActiveSlideId();
                _titles = StorylineFeature.ExtractStoryline(_manager);
                _currentSection = _titles.FirstOrDefault(item => item.SlideId == activeSlideId)?.SectionName;
                _view = CollectionViewSource.GetDefaultView(_titles);
                _view.GroupDescriptions.Clear();
                _view.GroupDescriptions.Add(new PropertyGroupDescription(nameof(StorylineFeature.StorylineInfo.SectionName)));
                _view.Filter = MatchesFilter;
                TitlesList.ItemsSource = _view;
                UpdateSummary();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineDialog.RefreshStoryline");
                MessageBox.Show("The storyline could not be refreshed.", "Storyline", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool MatchesFilter(object value)
        {
            var item = value as StorylineFeature.StorylineInfo;
            if (item == null) return false;

            string scope = GetSelectedTag(ScopeCombo);
            if (scope == "Section" && !string.Equals(item.SectionName, _currentSection, StringComparison.Ordinal))
                return false;

            string filter = GetSelectedTag(FilterCombo);
            if (filter == "Issues" && (!item.HasIssues || item.IsExcluded)) return false;
            if (filter == "Missing" && !item.IsMissing) return false;
            if (filter == "Review" && (item.IsExcluded || string.Equals(item.ReviewStatus, "Final", StringComparison.OrdinalIgnoreCase))) return false;
            if (filter == "Excluded" && !item.IsExcluded) return false;
            if (filter == "All" && item.IsExcluded) return false;

            string query = SearchBox?.Text?.Trim();
            if (string.IsNullOrEmpty(query)) return true;
            return Contains(item.TitleText, query) ||
                   Contains(item.SectionName, query) ||
                   Contains(item.IssueSummary, query) ||
                   Contains(item.FrameworkRole, query) ||
                   Contains(item.ReviewStatus, query);
        }

        private void Filter_Changed(object sender, EventArgs e)
        {
            if (_initializing || _view == null) return;
            _view.Refresh();
            UpdateSummary();
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshStoryline();
        }

        private void AutoRefreshTimer_Tick(object sender, EventArgs e)
        {
            if (AutoRefreshCheckBox?.IsChecked != true || !IsActive || HasActiveEditor()) return;
            RefreshStoryline();
        }

        private bool HasActiveEditor()
        {
            return Keyboard.FocusedElement is TextBox textBox && textBox.Name == "TxtEdit";
        }

        private void Item_Click(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is Button || e.OriginalSource is TextBox ||
                IsAncestorOf<TextBox>(e.OriginalSource as DependencyObject)) return;
            var item = (sender as FrameworkElement)?.DataContext as StorylineFeature.StorylineInfo;
            if (item != null) StorylineFeature.GoToSlide(_manager, item.SlideId);
        }

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            BeginEdit(FindItemContainer(sender as DependencyObject));
            e.Handled = true;
        }

        private void EditMenu_Click(object sender, RoutedEventArgs e)
        {
            var item = GetMenuItem(sender);
            if (item == null || !item.CanEdit) return;
            var container = TitlesList.ItemContainerGenerator.ContainerFromItem(item);
            BeginEdit(container);
        }

        private void BeginEdit(DependencyObject node)
        {
            TextBlock txtRead = FindVisualChild<TextBlock>(node, "TxtRead");
            TextBox txtEdit = FindVisualChild<TextBox>(node, "TxtEdit");
            Border border = null;
            if (txtRead != null && txtEdit != null)
            {
                txtRead.Visibility = Visibility.Collapsed;
                txtEdit.Visibility = Visibility.Visible;
                txtEdit.Focus();
                txtEdit.SelectAll();
                return;
            }

            border = FindAncestor<Border>(node);
            while (border != null)
            {
                txtRead = FindVisualChild<TextBlock>(border, "TxtRead");
                txtEdit = FindVisualChild<TextBox>(border, "TxtEdit");
                if (txtRead != null && txtEdit != null) break;
                border = FindAncestor<Border>(VisualTreeHelper.GetParent(border));
            }
            if (txtRead == null || txtEdit == null) return;
            txtRead.Visibility = Visibility.Collapsed;
            txtEdit.Visibility = Visibility.Visible;
            txtEdit.Focus();
            txtEdit.SelectAll();
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

        private void CommitEdit(TextBox textBox)
        {
            if (textBox == null || textBox.Visibility != Visibility.Visible) return;
            var item = textBox.DataContext as StorylineFeature.StorylineInfo;
            string newTitle = textBox.Text.Trim();
            if (item != null && newTitle != item.TitleText)
            {
                if (!StorylineFeature.UpdateTitle(_manager, item.SlideId, item.TitleShapeId, newTitle))
                {
                    MessageBox.Show("The detected headline could not be updated.", "Storyline",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    textBox.Text = item.TitleText;
                }
                else
                {
                    RefreshStoryline();
                    return;
                }
            }
            EndEdit(textBox);
        }

        private void CancelEdit(TextBox textBox)
        {
            var item = textBox?.DataContext as StorylineFeature.StorylineInfo;
            if (textBox != null && item != null) textBox.Text = item.TitleText;
            EndEdit(textBox);
        }

        private static void EndEdit(TextBox textBox)
        {
            if (textBox == null) return;
            var parent = FindAncestor<Border>(textBox);
            var txtRead = parent == null ? null : FindVisualChild<TextBlock>(parent, "TxtRead");
            if (txtRead != null) txtRead.Visibility = Visibility.Visible;
            textBox.Visibility = Visibility.Collapsed;
        }

        private void StorylineItemContextMenu_Opened(object sender, RoutedEventArgs e)
        {
            var menu = sender as ContextMenu;
            if (menu == null) return;
            var sectionsMenu = menu.Items.OfType<MenuItem>().FirstOrDefault(item => Equals(item.Tag, "Sections"));
            if (sectionsMenu == null) return;
            sectionsMenu.Items.Clear();
            foreach (string section in StorylineFeature.GetSectionNames(_manager))
            {
                var sectionItem = new MenuItem { Header = section, Tag = section, DataContext = menu.DataContext };
                sectionItem.Click += AssignSection_Click;
                sectionsMenu.Items.Add(sectionItem);
            }
            sectionsMenu.IsEnabled = sectionsMenu.Items.Count > 0;
        }

        private void GoToSlide_Click(object sender, RoutedEventArgs e)
        {
            var item = GetMenuItem(sender);
            if (item != null) StorylineFeature.GoToSlide(_manager, item.SlideId);
        }

        private void CopyHeadline_Click(object sender, RoutedEventArgs e)
        {
            var item = GetMenuItem(sender);
            if (item != null) Clipboard.SetText(item.TitleText ?? string.Empty);
        }

        private void CopySuggestion_Click(object sender, RoutedEventArgs e)
        {
            var item = GetMenuItem(sender);
            if (item != null) Clipboard.SetText(item.SuggestedTitle ?? string.Empty);
        }

        private void ApplySuggestion_Click(object sender, RoutedEventArgs e)
        {
            var item = GetMenuItem(sender);
            if (item == null || !item.CanEdit || string.IsNullOrWhiteSpace(item.SuggestedTitle)) return;
            if (StorylineFeature.UpdateTitle(_manager, item.SlideId, item.TitleShapeId, item.SuggestedTitle))
                RefreshStoryline();
        }

        private void SetStatus_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            var item = GetMenuItem(sender);
            if (item != null && menuItem != null &&
                StorylineFeature.SetReviewStatus(_manager, item.SlideId, menuItem.Header?.ToString()))
                RefreshStoryline();
        }

        private void AssignSection_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            var item = menuItem?.DataContext as StorylineFeature.StorylineInfo;
            if (item != null && StorylineFeature.AssignToSection(_manager, item.SlideId, menuItem.Tag?.ToString()))
                RefreshStoryline();
        }

        private void InsertSlideAfter_Click(object sender, RoutedEventArgs e)
        {
            var item = GetMenuItem(sender);
            if (item != null && StorylineFeature.InsertSlideAfter(_manager, item.SlideId)) RefreshStoryline();
        }

        private void MoveSlideUp_Click(object sender, RoutedEventArgs e) { MoveSlide(sender, -1); }
        private void MoveSlideDown_Click(object sender, RoutedEventArgs e) { MoveSlide(sender, 1); }

        private void MoveSlide(object sender, int delta)
        {
            var item = GetMenuItem(sender);
            if (item != null && StorylineFeature.MoveSlide(_manager, item.SlideId, delta)) RefreshStoryline();
        }

        private void ToggleHidden_Click(object sender, RoutedEventArgs e)
        {
            var item = GetMenuItem(sender);
            if (item != null && StorylineFeature.ToggleHidden(_manager, item.SlideId)) RefreshStoryline();
        }

        private void ToggleExcluded_Click(object sender, RoutedEventArgs e)
        {
            var item = GetMenuItem(sender);
            if (item != null && StorylineFeature.ToggleExcluded(_manager, item.SlideId)) RefreshStoryline();
        }

        private StorylineFeature.StorylineInfo GetMenuItem(object sender)
        {
            var menuItem = sender as MenuItem;
            return menuItem?.DataContext as StorylineFeature.StorylineInfo;
        }

        private void CopyViewButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Clipboard.SetText(StorylineFeature.BuildOutline(VisibleItems(), true, true, false));
                MessageBox.Show("The visible storyline was copied to the clipboard.", "Storyline");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineDialog.CopyViewButton_Click");
                MessageBox.Show("The storyline could not be copied.", "Storyline", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dialog = new SaveFileDialog
                {
                    Title = "Export Storyline",
                    Filter = "Text file (*.txt)|*.txt",
                    FileName = "storyline.txt",
                    AddExtension = true
                };
                if (dialog.ShowDialog(this) != true) return;
                File.WriteAllText(dialog.FileName, StorylineFeature.BuildOutline(VisibleItems(), true, true, false));
                MessageBox.Show("The storyline was exported.", "Storyline");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "StorylineDialog.ExportButton_Click");
                MessageBox.Show("The storyline could not be exported.", "Storyline", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private IEnumerable<StorylineFeature.StorylineInfo> VisibleItems()
        {
            return _view?.Cast<StorylineFeature.StorylineInfo>() ?? Enumerable.Empty<StorylineFeature.StorylineInfo>();
        }

        private void UpdateSummary()
        {
            if (SummaryText == null) return;
            var visible = VisibleItems().ToList();
            SummaryText.Text = $"{visible.Count} headlines · {visible.Count(item => item.HasIssues)} with issues";
        }

        private int GetActiveSlideId()
        {
            try
            {
                var slide = _manager.GetApplication().ActiveWindow.View.Slide as Microsoft.Office.Interop.PowerPoint.Slide;
                return slide?.SlideID ?? 0;
            }
            catch { return 0; }
        }

        private static string GetSelectedTag(ComboBox comboBox)
        {
            return (comboBox?.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "All";
        }

        private static bool Contains(string source, string query)
        {
            return (source ?? string.Empty).IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static DependencyObject FindItemContainer(DependencyObject node)
        {
            return FindAncestor<Border>(node);
        }

        private static T FindAncestor<T>(DependencyObject node) where T : DependencyObject
        {
            while (node != null)
            {
                if (node is T typed) return typed;
                if (node is ContextMenu) return node as T;
                node = node is Visual || node is System.Windows.Media.Media3D.Visual3D
                    ? VisualTreeHelper.GetParent(node)
                    : LogicalTreeHelper.GetParent(node);
            }
            return null;
        }

        private static T FindVisualChild<T>(DependencyObject parent, string name) where T : FrameworkElement
        {
            if (parent == null) return null;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typed && typed.Name == name) return typed;
                var nested = FindVisualChild<T>(child, name);
                if (nested != null) return nested;
            }
            return null;
        }

        private static bool IsAncestorOf<T>(DependencyObject node) where T : DependencyObject
        {
            while (node != null)
            {
                if (node is T) return true;
                node = VisualTreeHelper.GetParent(node);
            }
            return false;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e) { Close(); }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && !(e.OriginalSource is TextBox) &&
                !(e.OriginalSource is ComboBox) && !(e.OriginalSource is Button))
                DragMove();
        }
    }
}
