using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shell;
using System.Windows.Threading;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class FeatureExplorerWindow : Window, System.ComponentModel.INotifyPropertyChanged
    {
        private List<FeatureViewModel> _allFeatures;
        private string _currentCategory = "All";
        private bool _categoryNavigationUpdateScheduled;
        private bool _categoryNavigationRevealPending;
        private bool _isClosing;

        private IEnumerable<FeatureViewModel> _displayedFeatures;
        public IEnumerable<FeatureViewModel> DisplayedFeatures
        {
            get => _displayedFeatures;
            set { _displayedFeatures = value; PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(DisplayedFeatures))); }
        }

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        public FeatureExplorerWindow(string initialCategory = null, bool focusSearch = false)
        {
            InitializeComponent();
            DataContext = this;
            SourceInitialized += (sender, args) => FitToOwnerWorkArea();
            Closed += (sender, args) => _isClosing = true;
            Loaded += (sender, args) => UpdateCaptionHeight();
            HeaderDragRegion.SizeChanged += (sender, args) => UpdateCaptionHeight();
            SelectCategory(initialCategory);
            LoadFeatures();
            if (focusSearch)
            {
                Loaded += (s, e) =>
                {
                    SearchBox.Focus();
                    Keyboard.Focus(SearchBox);
                    SearchBox.SelectAll();
                };
            }
        }

        private void SelectCategory(string category)
        {
            string target = string.IsNullOrWhiteSpace(category) ? "All" : category;
            _currentCategory = target;
            if (CategoryFilters == null) return;
            foreach (var child in CategoryFilters.Children)
            {
                if (child is RadioButton radio &&
                    string.Equals(radio.Tag?.ToString(), target, StringComparison.Ordinal))
                {
                    radio.IsChecked = true;
                    return;
                }
            }
        }

        private void UpdateCaptionHeight()
        {
            if (HeaderDragRegion == null || WindowChromeControl == null || HeaderDragRegion.ActualHeight <= 0) return;
            System.Windows.Point bottom = HeaderDragRegion.TranslatePoint(
                new System.Windows.Point(0, HeaderDragRegion.ActualHeight), this);
            WindowChromeControl.CaptionHeight = Math.Max(0, bottom.Y);
        }

        private void FitToOwnerWorkArea()
        {
            var helper = new System.Windows.Interop.WindowInteropHelper(this);
            var screen = System.Windows.Forms.Screen.FromHandle(helper.Owner);
            var dpi = VisualTreeHelper.GetDpi(this);
            double availableWidth = Math.Max(320, (screen.WorkingArea.Width / dpi.DpiScaleX) - 32);
            double availableHeight = Math.Max(320, (screen.WorkingArea.Height / dpi.DpiScaleY) - 32);

            MinWidth = Math.Min(MinWidth, availableWidth);
            MinHeight = Math.Min(MinHeight, availableHeight);
            MaxWidth = availableWidth;
            MaxHeight = availableHeight;
            Width = Math.Min(Width, availableWidth);
            Height = Math.Min(Height, availableHeight);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                e.Handled = true;
                Close();
            }
            else if (e.Key == Key.F && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                e.Handled = true;
                SearchBox.Focus();
                SearchBox.SelectAll();
            }
        }

        private void LoadFeatures()
        {
            string iconColorMode = IconColorModes.Normalize(
                JsonService.Load<SidebarConfig>("sidebar_config.json")?.IconColorMode);
            _allFeatures = FeatureLibrary.Catalog
                .Where(record => record.UserVisible && !record.IsObsolete)
                .Select(record => new FeatureViewModel
                {
                    Id = record.Id,
                    Name = record.Name,
                    Description = record.Description,
                    DetailedHelpText = FeatureHelpFormatter.ToReadablePlainText(
                        string.IsNullOrEmpty(record.DetailedHelpText)
                            ? record.Description
                            : record.DetailedHelpText),
                    IconColor = ParseColor(record.Color),
                    IconColorMode = iconColorMode,
                    HelpImagePath = record.HelpImagePath,
                    Category = GetCategoryName(record.Category),
                    Maturity = record.Maturity,
                    CatalogRecord = record
                })
                .ToList();

            _allFeatures = _allFeatures.OrderBy(f => f.Name).ToList();
            FilterFeatures();
            
            if (_allFeatures.Count > 0)
            {
                FeatureList.SelectedIndex = 0;
            }
        }

        private static Brush ParseColor(string color)
        {
            try
            {
                return (Brush)new BrushConverter().ConvertFrom(
                    string.IsNullOrWhiteSpace(color) ? "#888888" : color);
            }
            catch
            {
                return Brushes.Gray;
            }
        }

        private static string GetCategoryName(FeatureCategory category)
        {
            switch (category)
            {
                case FeatureCategory.AlignAndSize: return "Align & Size";
                case FeatureCategory.ArrangeAndOrganize: return "Arrange & Organize";
                case FeatureCategory.TextAndFormat: return "Text & Format";
                case FeatureCategory.Tables: return "Tables";
                case FeatureCategory.VisualElements: return "Visual Elements";
                case FeatureCategory.StorylineAndReview: return "Storyline & Review";
                case FeatureCategory.DataAndAutomation: return "Data & Automation";
                case FeatureCategory.Libraries: return "Libraries";
                case FeatureCategory.SettingsAndHelp: return "Settings & Help";
                case FeatureCategory.Labs: return "Labs";
                default: return "Settings & Help";
            }
        }

        private void FilterFeatures()
        {
            if (_allFeatures == null) return;

            var query = SearchBox?.Text ?? "";

            var filtered = _allFeatures.Where(f =>
                (_currentCategory == "All" || f.Category == _currentCategory) &&
                FeatureCatalogSearch.Matches(f.CatalogRecord, query)
            ).ToList();
            DisplayedFeatures = filtered;

            if (FeatureList != null && !filtered.Contains(FeatureList.SelectedItem as FeatureViewModel))
                FeatureList.SelectedItem = filtered.FirstOrDefault();
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            FilterFeatures();
        }

        private void Category_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton rb)
            {
                _currentCategory = rb.Tag?.ToString() ?? "All";
                FilterFeatures();
                ScheduleCategoryNavigationUpdate(revealCategory: true);
            }
        }

        private void Category_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (sender is RadioButton category)
            {
                ScheduleCategoryNavigationUpdate(revealCategory: true);
            }
        }

        private void CategoryScrollViewer_Loaded(object sender, RoutedEventArgs e)
        {
            ScheduleCategoryNavigationUpdate(revealCategory: true);
        }

        private void CategoryScrollViewer_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            ScheduleCategoryNavigationUpdate(revealCategory: true);
        }

        private void CategoryScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            UpdateCategoryNavigation(revealCategory: false);
        }

        private void CategoryScrollLeftButton_Click(object sender, RoutedEventArgs e)
        {
            ScrollCategories(-1);
        }

        private void CategoryScrollRightButton_Click(object sender, RoutedEventArgs e)
        {
            ScrollCategories(1);
        }

        private void ScrollCategories(int direction)
        {
            if (!CanUpdateCategoryNavigation()) return;

            FeatureExplorerNavigationState state = FeatureExplorerNavigationPlanner.Create(
                CategoryScrollViewer.ExtentWidth,
                CategoryNavigationHost.ActualWidth,
                CategoryScrollViewer.ViewportWidth,
                CategoryScrollViewer.HorizontalOffset);
            if (!state.RequiresNavigation) return;

            double step = Math.Max(80, state.ViewportWidth * 0.75);
            CategoryScrollViewer.ScrollToHorizontalOffset(
                state.HorizontalOffset + (direction * step));
            UpdateCategoryNavigation(revealCategory: false);
        }

        private void ScheduleCategoryNavigationUpdate(bool revealCategory)
        {
            if (_isClosing || Dispatcher.HasShutdownStarted || Dispatcher.HasShutdownFinished) return;

            _categoryNavigationRevealPending |= revealCategory;
            if (_categoryNavigationUpdateScheduled) return;

            _categoryNavigationUpdateScheduled = true;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                _categoryNavigationUpdateScheduled = false;
                bool shouldRevealCategory = _categoryNavigationRevealPending;
                _categoryNavigationRevealPending = false;
                UpdateCategoryNavigation(shouldRevealCategory);
            }), DispatcherPriority.Loaded);
        }

        private bool CanUpdateCategoryNavigation()
        {
            return !_isClosing && IsLoaded && CategoryNavigationHost != null && CategoryFilters != null &&
                CategoryScrollViewer != null && CategoryScrollLeftButton != null && CategoryScrollRightButton != null &&
                CategoryNavigationHost.IsLoaded && CategoryScrollViewer.IsLoaded &&
                CategoryNavigationHost.ActualWidth > 0 &&
                PresentationSource.FromVisual(CategoryNavigationHost) != null &&
                PresentationSource.FromVisual(CategoryScrollViewer) != null;
        }

        private void UpdateCategoryNavigation(bool revealCategory)
        {
            if (!CanUpdateCategoryNavigation()) return;

            FeatureExplorerNavigationState state = FeatureExplorerNavigationPlanner.Create(
                CategoryScrollViewer.ExtentWidth,
                CategoryNavigationHost.ActualWidth,
                CategoryScrollViewer.ViewportWidth,
                CategoryScrollViewer.HorizontalOffset);

            Visibility navigationVisibility = state.RequiresNavigation ? Visibility.Visible : Visibility.Collapsed;
            bool navigationChromeChanged = CategoryScrollLeftButton.Visibility != navigationVisibility ||
                CategoryScrollRightButton.Visibility != navigationVisibility;
            CategoryScrollLeftButton.Visibility = navigationVisibility;
            CategoryScrollRightButton.Visibility = navigationVisibility;
            CategoryScrollLeftButton.IsEnabled = state.CanScrollLeft;
            CategoryScrollRightButton.IsEnabled = state.CanScrollRight;

            if (Math.Abs(CategoryScrollViewer.HorizontalOffset - state.HorizontalOffset) > FeatureExplorerNavigationPlanner.Epsilon)
            {
                CategoryScrollViewer.ScrollToHorizontalOffset(state.HorizontalOffset);
            }

            if (navigationChromeChanged)
            {
                ScheduleCategoryNavigationUpdate(revealCategory: true);
                return;
            }

            if (revealCategory)
            {
                RevealCurrentCategory();
            }
        }

        private void RevealCurrentCategory()
        {
            RadioButton category = GetFocusedOrCheckedCategory();
            if (category == null || !category.IsLoaded || CategoryScrollViewer.ViewportWidth <= 0 ||
                PresentationSource.FromVisual(category) == null)
            {
                return;
            }

            Rect bounds = category.TransformToAncestor(CategoryScrollViewer)
                .TransformBounds(new Rect(category.RenderSize));
            double targetOffset = FeatureExplorerNavigationPlanner.CalculateRevealOffset(
                CategoryScrollViewer.ExtentWidth,
                CategoryScrollViewer.ViewportWidth,
                CategoryScrollViewer.HorizontalOffset,
                CategoryScrollViewer.HorizontalOffset + bounds.Left,
                bounds.Width);
            if (Math.Abs(CategoryScrollViewer.HorizontalOffset - targetOffset) > FeatureExplorerNavigationPlanner.Epsilon)
            {
                CategoryScrollViewer.ScrollToHorizontalOffset(targetOffset);
                UpdateCategoryNavigation(revealCategory: false);
            }
        }

        private RadioButton GetFocusedOrCheckedCategory()
        {
            RadioButton focusedCategory = Keyboard.FocusedElement as RadioButton;
            if (IsCategoryFilter(focusedCategory)) return focusedCategory;

            foreach (object child in CategoryFilters.Children)
            {
                RadioButton category = child as RadioButton;
                if (category != null && category.IsChecked == true) return category;
            }

            return null;
        }

        private bool IsCategoryFilter(RadioButton category)
        {
            return category != null && CategoryFilters.Children.Contains(category);
        }

        private void FeatureList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FeatureList.SelectedItem is FeatureViewModel feature)
            {
                DetailsPanel.Visibility = Visibility.Visible;
                DetailName.Text = feature.Name;
                DetailDescription.Text = feature.Description;
                DetailDetailedHelp.Text = feature.DetailedHelpText;
                DetailMaturity.Text = feature.Maturity == FeatureMaturity.Experimental
                    ? "LABS · EXPERIMENTAL"
                    : "PREVIEW";
                DetailMaturity.Visibility = feature.Maturity == FeatureMaturity.Stable
                    ? Visibility.Collapsed
                    : Visibility.Visible;

                try
                {
                    DetailIcon.FeatureId = feature.Id;
                    DetailIcon.IconColor = feature.IconColor;
                    DetailIcon.ColorMode = feature.IconColorMode;
                }
                catch { DetailIcon.FeatureId = string.Empty; }

                var assignment = ShortcutManager.GetAssignment(feature.Id);
                if (assignment != null && !string.IsNullOrEmpty(assignment.ShortcutText))
                {
                    ShortcutPanel.Visibility = Visibility.Visible;
                    DetailShortcut.Text = assignment.ShortcutText;
                    ShortcutPanel.ToolTip = assignment.NeedsAttention
                        ? "Requires attention: " + assignment.AttentionReason
                        : null;
                }
                else
                {
                    ShortcutPanel.Visibility = Visibility.Collapsed;
                }

                if (!string.IsNullOrEmpty(feature.HelpImagePath))
                {
                    try
                    {
                        DetailImage.Source = new BitmapImage(new Uri(feature.HelpImagePath, UriKind.RelativeOrAbsolute));
                        DetailImageBorder.Visibility = Visibility.Visible;
                    }
                    catch
                    {
                        DetailImageBorder.Visibility = Visibility.Collapsed;
                    }
                }
                else
                {
                    DetailImageBorder.Visibility = Visibility.Collapsed;
                }
            }
            else
            {
                DetailsPanel.Visibility = Visibility.Collapsed;
                DetailsRegion.Visibility = Visibility.Collapsed;
                DetailsDivider.Visibility = Visibility.Collapsed;
                DetailsColumn.Width = new GridLength(0);
            }

            if (FeatureList.SelectedItem != null)
            {
                DetailsRegion.Visibility = Visibility.Visible;
                DetailsDivider.Visibility = Visibility.Visible;
                DetailsColumn.Width = new GridLength(280);
            }
        }

        public class FeatureViewModel
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public string Description { get; set; }
            public string DetailedHelpText { get; set; }
            public Brush IconColor { get; set; }
            public string IconColorMode { get; set; }
            public string HelpImagePath { get; set; }
            public string Category { get; set; }
            public FeatureMaturity Maturity { get; set; }
            public FeatureCatalogRecord CatalogRecord { get; set; }

        }
    }
}
