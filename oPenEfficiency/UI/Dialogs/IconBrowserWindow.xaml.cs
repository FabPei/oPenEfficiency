using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI.Dialogs
{
    public enum IconBrowserStartView
    {
        Default,
        Search,
        Svg,
        PowerPointShape,
        Favorites
    }

    public partial class IconBrowserWindow : Window
    {
        private const string LightThemeAccentColor = "#2563EB";
        private const string DarkThemeAccentColor = "#3B82F6";
        private readonly PowerPointManager _manager;
        private readonly IconCatalogService _catalogService;
        private const int InitialVisibleIconCount = 72;
        private const int VisibleIconPageSize = 72;
        private const int GeometryLoadBatchSize = 12;
        private readonly List<IconCatalogItem> _allIcons = new List<IconCatalogItem>();
        private List<IconCatalogItem> _filteredIcons = new List<IconCatalogItem>();
        private readonly Dictionary<string, IconCardViewModel> _viewModelCache = new Dictionary<string, IconCardViewModel>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _collectionLabels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly DispatcherTimer _searchDebounceTimer;
        private readonly DispatcherTimer _statusResetTimer;
        private Brush _selectedColorBrush = Brushes.Black;
        private bool _isUpdatingColor;
        private bool _isCatalogLoaded;
        private bool _isDetailsVisible;
        private int _visibleIconCount;
        private int _renderGeneration;

        public IconBrowserWindow(
            PowerPointManager manager,
            IconBrowserStartView startView = IconBrowserStartView.Default)
        {
            _manager = manager;
            _catalogService = new IconCatalogService();
            _searchDebounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(180) };
            _searchDebounceTimer.Tick += SearchDebounceTimer_Tick;
            _statusResetTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2.2) };
            _statusResetTimer.Tick += StatusResetTimer_Tick;

            InitializeComponent();
            InitializeWindowState();
            if (startView == IconBrowserStartView.PowerPointShape)
                InsertModeComboBox.SelectedIndex = 1;
            else if (startView == IconBrowserStartView.Svg)
                InsertModeComboBox.SelectedIndex = 0;
            else if (startView == IconBrowserStartView.Favorites)
                FavoritesOnlyCheckBox.IsChecked = true;
            if (startView == IconBrowserStartView.Search)
            {
                Loaded += (s, e) =>
                {
                    SearchBox.Focus();
                    Keyboard.Focus(SearchBox);
                    SearchBox.SelectAll();
                };
            }
            Loaded += IconBrowserWindow_Loaded;
            Closed += (sender, args) => _renderGeneration++;
        }

        private void InitializeWindowState()
        {
            SetSelectedColor(GetDefaultColorForCurrentTheme());
            PopulateThemeColorSwatches();
            PopulateInsertModes();
            InsertButton.IsEnabled = false;
            StatusLabel.Text = "";
            InsertFeedbackText.Text = "Select one or more icons, choose SVG or PowerPoint shape, then press Enter to insert them. Press Ctrl+F to search.";
            SetFilterControlsEnabled(false);
            ClearSelectionPreview();
            UpdateDetailsVisibility();
        }

        private async void IconBrowserWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Loaded -= IconBrowserWindow_Loaded;
            PopulateThemeColorSwatches();
            await LoadCatalogAsync();
        }

        private async Task LoadCatalogAsync()
        {
            SetStatus("Loading icons");

            try
            {
                var manifest = await Task.Run(() => _catalogService.LoadManifest());

                PopulateFilters(manifest);
                _allIcons.Clear();
                _allIcons.AddRange(manifest.Icons);
                _viewModelCache.Clear();
                _isCatalogLoaded = true;
                SetFilterControlsEnabled(true);
                ApplyFilters();
                SetStatus("Ready");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconBrowserWindow.LoadCatalogAsync");
                MessageBox.Show("Could not load the icon catalog.", "Icon Library", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void PopulateFilters(IconCollectionManifest manifest)
        {
            CollectionFilter.Items.Clear();
            _collectionLabels.Clear();
            CollectionFilter.Items.Add(FilterOption.All());
            foreach (var collection in manifest.Collections.OrderBy(c => c.Name))
            {
                _collectionLabels[collection.Id ?? string.Empty] = string.IsNullOrWhiteSpace(collection.Name) ? ToDisplayLabel(collection.Id) : collection.Name;
                CollectionFilter.Items.Add(new FilterOption(collection.Id, _collectionLabels[collection.Id ?? string.Empty]));
            }
            CollectionFilter.SelectedIndex = 0;

            ThemeFilter.Items.Clear();
            ThemeFilter.Items.Add(FilterOption.All());
            foreach (var theme in _catalogService.GetAvailableThemes(manifest.Icons))
            {
                ThemeFilter.Items.Add(new FilterOption(theme, ToDisplayLabel(theme)));
            }
            ThemeFilter.SelectedIndex = 0;

            StyleFilter.Items.Clear();
            StyleFilter.Items.Add(FilterOption.All());
            foreach (var style in _catalogService.GetAvailableStyles(manifest.Icons))
            {
                StyleFilter.Items.Add(new FilterOption(style, ToDisplayLabel(style)));
            }
            SelectDefaultStyle();
        }

        private void ApplyFilters()
        {
            if (!_isCatalogLoaded)
            {
                return;
            }

            string search = SearchBox.Text;
            string collection = GetSelectedFilterValue(CollectionFilter);
            string theme = GetSelectedFilterValue(ThemeFilter);
            string style = GetSelectedFilterValue(StyleFilter);
            bool favoritesOnly = FavoritesOnlyCheckBox.IsChecked == true;

            _filteredIcons = _catalogService.FilterIcons(
                _allIcons,
                search,
                NormalizeFilter(collection),
                NormalizeFilter(theme),
                NormalizeFilter(style),
                favoritesOnly)
                .ToList();
            _visibleIconCount = Math.Min(InitialVisibleIconCount, _filteredIcons.Count);
            RenderVisibleIcons();
        }

        private static string NormalizeFilter(string value)
        {
            return string.Equals(value, "All", StringComparison.OrdinalIgnoreCase) ? null : value;
        }

        private static string GetSelectedFilterValue(ComboBox comboBox)
        {
            if (comboBox?.SelectedItem is FilterOption option)
            {
                return option.Value;
            }

            return comboBox?.SelectedItem as string;
        }

        private void RefreshColorBrushes()
        {
            foreach (var icon in _viewModelCache.Values)
            {
                icon.SetColorBrush(_selectedColorBrush);
            }

            if (IconList.ItemsSource != null)
            {
                IconList.Items.Refresh();
            }

            RefreshSelectionPreview(GetSelectedIcons());
        }

        private void RefreshSelectionPreview(IReadOnlyList<IconCardViewModel> selectedIcons)
        {
            if (selectedIcons == null || selectedIcons.Count == 0)
            {
                ClearSelectionPreview();
                return;
            }

            var primarySelection = selectedIcons[0];
            InsertButton.IsEnabled = true;
            InsertButton.Content = selectedIcons.Count > 1 ? $"Insert {selectedIcons.Count} Icons" : "Insert Selected";

            if (selectedIcons.Count == 1)
            {
                SelectedNameText.Text = primarySelection.Name;
                SelectedMetaText.Text = $"{primarySelection.CollectionLabel}  |  Theme: {primarySelection.Theme}  |  Style: {primarySelection.StyleLabel}";
                SelectedTagsText.Text = string.IsNullOrWhiteSpace(primarySelection.TagsLabel) ? "No tags available for this icon." : primarySelection.TagsLabel;
                SelectedPathText.Text = primarySelection.RelativePath;
            }
            else
            {
                SelectedNameText.Text = $"{selectedIcons.Count} icons selected";
                SelectedMetaText.Text = $"Primary preview: {primarySelection.Name}  |  Batch insert will place the selected icons in a centered grid.";
                SelectedTagsText.Text = string.Join(", ", selectedIcons
                    .SelectMany(icon => icon.Item.Tags ?? new List<string>())
                    .Where(tag => !string.IsNullOrWhiteSpace(tag))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Take(18));
                if (string.IsNullOrWhiteSpace(SelectedTagsText.Text))
                {
                    SelectedTagsText.Text = "No shared tags available for this selection.";
                }

                SelectedPathText.Text = string.Join(Environment.NewLine, selectedIcons
                    .Take(6)
                    .Select(icon => icon.RelativePath));

                if (selectedIcons.Count > 6)
                {
                    SelectedPathText.Text += $"{Environment.NewLine}... and {selectedIcons.Count - 6} more";
                }
            }

            SelectedIconPath.Data = primarySelection.Geometry;
            if (primarySelection.IsFilled)
            {
                SelectedIconPath.Fill = _selectedColorBrush;
                SelectedIconPath.Stroke = null;
                SelectedIconPath.StrokeThickness = 0;
            }
            else
            {
                SelectedIconPath.Fill = Brushes.Transparent;
                SelectedIconPath.Stroke = _selectedColorBrush;
                SelectedIconPath.StrokeThickness = primarySelection.StrokeThickness;
            }
        }

        private void ClearSelectionPreview()
        {
            InsertButton.IsEnabled = false;
            InsertButton.Content = "Insert Selected";
            SelectedNameText.Text = "No icon selected";
            SelectedMetaText.Text = "Choose one or more icons from the catalog.";
            SelectedTagsText.Text = "";
            SelectedPathText.Text = "";
            SelectedIconPath.Data = null;
            SelectedIconPath.Fill = Brushes.Transparent;
            SelectedIconPath.Stroke = null;
            SelectedIconPath.StrokeThickness = 0;
        }

        private void SetSelectedColor(string hexColor)
        {
            try
            {
                var color = (Color)ColorConverter.ConvertFromString(hexColor);
                _selectedColorBrush = new SolidColorBrush(color);
                if (_selectedColorBrush.CanFreeze) _selectedColorBrush.Freeze();

                _isUpdatingColor = true;
                ColorHexTextBox.Text = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
                ColorPreviewBorder.Background = new SolidColorBrush(color);
                _isUpdatingColor = false;

                RefreshColorBrushes();
            }
            catch
            {
                _isUpdatingColor = false;
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                DragMove();
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Filter_Changed(object sender, EventArgs e)
        {
            if (sender == SearchBox)
            {
                _searchDebounceTimer.Stop();
                _searchDebounceTimer.Start();
                return;
            }

            if (sender == CollectionFilter)
            {
                SyncFiltersForSelectedCollection();
            }

            ApplyFilters();
        }

        private void IconList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshSelectionPreview(GetSelectedIcons());
        }

        private void ColorHexTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingColor)
            {
                return;
            }

            string text = ColorHexTextBox.Text?.Trim();
            if (!string.IsNullOrWhiteSpace(text) && text.StartsWith("#", StringComparison.Ordinal) && text.Length == 7)
            {
                SetSelectedColor(text);
            }
        }

        private void PickColorButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new oPenEfficiency.UI.WpfColorPickerDialog(ColorHexTextBox.Text) { Owner = this };
            if (dialog.ShowDialog() == true)
            {
                SetSelectedColor($"#{dialog.SelectedColor.R:X2}{dialog.SelectedColor.G:X2}{dialog.SelectedColor.B:X2}");
            }
        }

        private void PresetColorButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string hexColor)
            {
                SetSelectedColor(hexColor);
            }
        }

        private void FavoriteButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement element && element.Tag is IconCardViewModel viewModel)
            {
                bool newValue = !viewModel.Item.IsFavorite;
                _catalogService.SetFavorite(viewModel.Item.Id, newValue);
                viewModel.Item.IsFavorite = newValue;
                viewModel.RefreshFavoriteState();
                IconList.Items.Refresh();

                if (FavoritesOnlyCheckBox.IsChecked == true && !newValue)
                {
                    ApplyFilters();
                }
            }
        }

        private void InsertButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedIcons = GetSelectedIcons();
            if (selectedIcons.Count == 0)
            {
                return;
            }

            IconInsertMode insertMode = GetSelectedInsertMode();

            bool inserted = _catalogService.InsertIconsIntoActiveSlide(
                _manager,
                selectedIcons.Select(icon => icon.Item),
                ColorHexTextBox.Text,
                insertMode,
                48f);
            if (inserted)
            {
                SetStatus("Inserted");
                string modeLabel = insertMode == IconInsertMode.PowerPointShape ? "PowerPoint shape" : "SVG";
                InsertFeedbackText.Text = selectedIcons.Count == 1
                    ? $"Inserted '{selectedIcons[0].Name}' as a {modeLabel} into the active slide."
                    : $"Inserted {selectedIcons.Count} icons as {modeLabel}s into the active slide.";
            }
            else
            {
                SetStatus("Insert failed");
                InsertFeedbackText.Text = "Insertion failed. Make sure a presentation and a normal slide view are active. For complex icons, SVG mode is the safest option.";
                MessageBox.Show("Could not insert the selected icon selection into the active slide.", "Icon Library", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ResetFiltersButton_Click(object sender, RoutedEventArgs e)
        {
            SearchBox.Text = "";
            CollectionFilter.SelectedIndex = 0;
            ThemeFilter.SelectedIndex = 0;
            SelectDefaultStyle();
            FavoritesOnlyCheckBox.IsChecked = false;
            ApplyFilters();
            SearchBox.Focus();
            SetStatus("Filters reset");
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Enter && InsertButton.IsEnabled && !IsTypingControlFocused(Keyboard.FocusedElement as DependencyObject))
            {
                InsertButton_Click(InsertButton, new RoutedEventArgs());
                e.Handled = true;
                return;
            }

            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.F)
            {
                SearchBox.Focus();
                SearchBox.SelectAll();
                e.Handled = true;
                return;
            }

            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.D)
            {
                ToggleDetailsVisibility();
                e.Handled = true;
            }
        }

        private void IconList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (GetSelectedIcons().Count > 0 && InsertButton.IsEnabled)
            {
                InsertButton_Click(InsertButton, new RoutedEventArgs());
            }
        }

        private void LoadMoreButton_Click(object sender, RoutedEventArgs e)
        {
            if (_visibleIconCount >= _filteredIcons.Count)
            {
                return;
            }

            _visibleIconCount = Math.Min(_visibleIconCount + VisibleIconPageSize, _filteredIcons.Count);
            RenderVisibleIcons();
            SetStatus("More icons loaded");
        }

        private void SearchDebounceTimer_Tick(object sender, EventArgs e)
        {
            _searchDebounceTimer.Stop();
            ApplyFilters();
        }

        private void SetStatus(string message)
        {
            StatusLabel.Text = message;
            _statusResetTimer.Stop();
            _statusResetTimer.Start();
        }

        private void StatusResetTimer_Tick(object sender, EventArgs e)
        {
            _statusResetTimer.Stop();
            StatusLabel.Text = "";
        }

        private void RenderVisibleIcons()
        {
            var visibleItems = _filteredIcons.Take(_visibleIconCount).ToList();
            var visibleIcons = visibleItems.Select(GetOrCreateViewModel).ToList();
            int renderGeneration = ++_renderGeneration;
            var existingSelectionIds = new HashSet<string>(
                GetSelectedIcons()
                    .Select(icon => icon.Item.Id ?? string.Empty),
                StringComparer.OrdinalIgnoreCase);

            IconList.ItemsSource = visibleIcons;
            EmptyStatePanel.Visibility = _filteredIcons.Count == 0 ? Visibility.Visible : Visibility.Collapsed;

            bool canLoadMore = _filteredIcons.Count > _visibleIconCount;
            LoadMorePanel.Visibility = _filteredIcons.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            LoadMoreButton.Visibility = canLoadMore ? Visibility.Visible : Visibility.Collapsed;
            VisibleCountLabel.Margin = canLoadMore ? new Thickness(0, 0, 12, 0) : new Thickness(0);
            VisibleCountLabel.Text = canLoadMore
                ? $"{_filteredIcons.Count:N0} icons found · Showing {_visibleIconCount:N0} of {_filteredIcons.Count:N0}"
                : $"{_filteredIcons.Count:N0} icons found · Showing all";

            if (visibleIcons.Count > 0)
            {
                if (existingSelectionIds.Count > 0)
                {
                    IconList.SelectedItems.Clear();

                    foreach (var visibleIcon in visibleIcons.Where(icon => existingSelectionIds.Contains(icon.Item.Id ?? string.Empty)))
                    {
                        IconList.SelectedItems.Add(visibleIcon);
                    }

                    if (IconList.SelectedItems.Count > 0)
                    {
                        RefreshSelectionPreview(GetSelectedIcons());
                    }
                    else
                    {
                        IconList.SelectedIndex = 0;
                    }
                }
                else
                {
                    IconList.SelectedIndex = 0;
                }
            }
            else
            {
                IconList.SelectedItem = null;
                ClearSelectionPreview();
            }

            LoadVisibleGeometriesAsync(visibleIcons, renderGeneration);
        }

        private async void LoadVisibleGeometriesAsync(IReadOnlyList<IconCardViewModel> icons, int renderGeneration)
        {
            try
            {
                for (int index = 0; index < icons.Count; index += GeometryLoadBatchSize)
                {
                    if (renderGeneration != _renderGeneration || !IsLoaded)
                    {
                        return;
                    }

                    var batch = icons.Skip(index).Take(GeometryLoadBatchSize).ToList();
                    var results = await Task.Run(() => batch.Select(icon => icon.ParseGeometry()).ToList());

                    if (renderGeneration != _renderGeneration || !IsLoaded)
                    {
                        return;
                    }

                    foreach (var result in results)
                    {
                        result.Icon.ApplyGeometry(result.Geometry, result.IsFilled);
                    }

                    var selectedIcons = GetSelectedIcons();
                    if (selectedIcons.Count > 0 && results.Any(result => ReferenceEquals(result.Icon, selectedIcons[0])))
                    {
                        RefreshSelectionPreview(selectedIcons);
                    }

                    await Dispatcher.Yield(DispatcherPriority.Background);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconBrowserWindow.LoadVisibleGeometriesAsync");
            }
        }

        private void SelectDefaultStyle()
        {
            StyleFilter.SelectedItem = StyleFilter.Items
                .Cast<object>()
                .OfType<FilterOption>()
                .FirstOrDefault(i => string.Equals(i.Value, "regular", StringComparison.OrdinalIgnoreCase))
                ?? StyleFilter.Items.Cast<object>().FirstOrDefault();
        }

        private void PopulateInsertModes()
        {
            InsertModeComboBox.Items.Clear();
            InsertModeComboBox.Items.Add(new FilterOption("svg", "SVG image"));
            InsertModeComboBox.Items.Add(new FilterOption("shape", "PowerPoint shape"));
            InsertModeComboBox.SelectedIndex = 0;
        }

        private IconInsertMode GetSelectedInsertMode()
        {
            if (InsertModeComboBox.SelectedItem is FilterOption option &&
                string.Equals(option.Value, "shape", StringComparison.OrdinalIgnoreCase))
            {
                return IconInsertMode.PowerPointShape;
            }

            return IconInsertMode.Svg;
        }

        private void SyncFiltersForSelectedCollection()
        {
            string selectedCollection = NormalizeFilter(GetSelectedFilterValue(CollectionFilter));
            if (string.IsNullOrWhiteSpace(selectedCollection))
            {
                return;
            }

            var collectionIcons = _allIcons
                .Where(i => string.Equals(i.CollectionId, selectedCollection, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (collectionIcons.Count == 0)
            {
                return;
            }

            EnsureFilterValueAvailable(ThemeFilter, collectionIcons.Select(i => i.Theme));
            EnsureStyleValueAvailable(collectionIcons.Select(i => i.Style));
        }

        private void EnsureFilterValueAvailable(ComboBox comboBox, IEnumerable<string> validValues)
        {
            string selectedValue = NormalizeFilter(GetSelectedFilterValue(comboBox));
            if (string.IsNullOrWhiteSpace(selectedValue))
            {
                return;
            }

            var validSet = new HashSet<string>(
                validValues.Where(v => !string.IsNullOrWhiteSpace(v)),
                StringComparer.OrdinalIgnoreCase);

            if (validSet.Contains(selectedValue))
            {
                return;
            }

            comboBox.SelectedIndex = 0;
        }

        private void EnsureStyleValueAvailable(IEnumerable<string> validStyles)
        {
            string selectedStyle = NormalizeFilter(GetSelectedFilterValue(StyleFilter));
            var validSet = new HashSet<string>(
                validStyles.Where(v => !string.IsNullOrWhiteSpace(v)),
                StringComparer.OrdinalIgnoreCase);

            if (!string.IsNullOrWhiteSpace(selectedStyle) && validSet.Contains(selectedStyle))
            {
                return;
            }

            var defaultOption = StyleFilter.Items
                .Cast<object>()
                .OfType<FilterOption>()
                .FirstOrDefault(i => validSet.Contains(i.Value) && string.Equals(i.Value, "regular", StringComparison.OrdinalIgnoreCase))
                ?? StyleFilter.Items
                    .Cast<object>()
                    .OfType<FilterOption>()
                    .FirstOrDefault(i => validSet.Contains(i.Value) && string.Equals(i.Value, "outline", StringComparison.OrdinalIgnoreCase))
                ?? StyleFilter.Items
                    .Cast<object>()
                    .OfType<FilterOption>()
                    .FirstOrDefault(i => validSet.Contains(i.Value))
                ?? StyleFilter.Items
                    .Cast<object>()
                    .OfType<FilterOption>()
                    .FirstOrDefault(i => string.Equals(i.Value, "All", StringComparison.OrdinalIgnoreCase));

            if (defaultOption != null)
            {
                StyleFilter.SelectedItem = defaultOption;
            }
        }

        private void SetFilterControlsEnabled(bool isEnabled)
        {
            SearchBox.IsEnabled = isEnabled;
            CollectionFilter.IsEnabled = isEnabled;
            ThemeFilter.IsEnabled = isEnabled;
            StyleFilter.IsEnabled = isEnabled;
            FavoritesOnlyCheckBox.IsEnabled = isEnabled;
            LoadMoreButton.IsEnabled = isEnabled;
        }

        private IconCardViewModel GetOrCreateViewModel(IconCatalogItem item)
        {
            if (item == null)
            {
                return null;
            }

            string key = item.Id ?? string.Empty;
            if (_viewModelCache.TryGetValue(key, out IconCardViewModel existing))
            {
                existing.SetColorBrush(_selectedColorBrush);
                return existing;
            }

            var created = new IconCardViewModel(
                item,
                _catalogService.ResolveIconPath(item),
                _selectedColorBrush,
                GetCollectionLabel(item.CollectionId));
            _viewModelCache[key] = created;
            return created;
        }

        private static bool IsTypingControlFocused(DependencyObject focusedElement)
        {
            if (focusedElement == null)
            {
                return false;
            }

            if (focusedElement is TextBox || focusedElement is RichTextBox)
            {
                return true;
            }

            if (focusedElement is ComboBox comboBox && comboBox.IsEditable)
            {
                return true;
            }

            return IsTypingControlFocused(System.Windows.Media.VisualTreeHelper.GetParent(focusedElement));
        }

        private List<IconCardViewModel> GetSelectedIcons()
        {
            return IconList.SelectedItems
                .Cast<object>()
                .OfType<IconCardViewModel>()
                .ToList();
        }

        private static string GetDefaultColorForCurrentTheme()
        {
            try
            {
                var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
                bool isDarkTheme = string.Equals(config?.AppTheme, "Dark", StringComparison.OrdinalIgnoreCase);
                return isDarkTheme ? LightThemeAccentColor : DarkThemeAccentColor;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconBrowserWindow.GetDefaultColorForCurrentTheme");
                return LightThemeAccentColor;
            }
        }

        private string GetCollectionLabel(string collectionId)
        {
            if (string.IsNullOrWhiteSpace(collectionId))
            {
                return "Unknown";
            }

            return _collectionLabels.TryGetValue(collectionId, out string label) ? label : ToDisplayLabel(collectionId);
        }

        private static string ToDisplayLabel(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "Unknown";
            }

            var parts = value.Replace("-", " ").Replace("_", " ").Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(" ", parts.Select(p => char.ToUpperInvariant(p[0]) + p.Substring(1)));
        }

        private void ToggleDetailsButton_Click(object sender, RoutedEventArgs e)
        {
            ToggleDetailsVisibility();
        }

        private void ToggleDetailsVisibility()
        {
            _isDetailsVisible = !_isDetailsVisible;
            UpdateDetailsVisibility();
        }

        private void UpdateDetailsVisibility()
        {
            if (DetailsPanel == null || DetailsColumn == null || ToggleDetailsButton == null)
            {
                return;
            }

            DetailsPanel.Visibility = _isDetailsVisible ? Visibility.Visible : Visibility.Collapsed;
            DetailsColumn.Width = _isDetailsVisible ? new GridLength(360) : new GridLength(0);
            ToggleDetailsButton.Content = _isDetailsVisible ? "Hide" : "Details";
            ToggleDetailsButton.ToolTip = _isDetailsVisible ? "Hide details (Ctrl+D)" : "Show details (Ctrl+D)";
        }

        private void PopulateThemeColorSwatches()
        {
            ThemeColorSwatchPanel.Children.Clear();

            foreach (var swatch in GetPresentationThemeColors())
            {
                ThemeColorSwatchPanel.Children.Add(CreateThemeColorButton(swatch.Hex, swatch.Label));
            }
        }

        private IEnumerable<(string Hex, string Label)> GetPresentationThemeColors()
        {
            try
            {
                var app = _manager?.GetApplication();
                var pres = app?.ActivePresentation;
                if (pres?.SlideMaster?.Theme?.ThemeColorScheme == null)
                {
                    return GetFallbackThemeColors();
                }

                var scheme = pres.SlideMaster.Theme.ThemeColorScheme;
                var indices = new[]
                {
                    (Office.MsoThemeColorSchemeIndex.msoThemeDark1, "Dark 1"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeLight1, "Light 1"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeAccent1, "Accent 1"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeAccent2, "Accent 2"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeAccent3, "Accent 3"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeAccent4, "Accent 4"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeAccent5, "Accent 5"),
                    (Office.MsoThemeColorSchemeIndex.msoThemeAccent6, "Accent 6")
                };

                return indices
                    .Select(entry => (Hex: ColorConstants.ToHex(scheme.Colors(entry.Item1).RGB), Label: entry.Item2))
                    .Distinct()
                    .ToList();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconBrowserWindow.GetPresentationThemeColors");
                return GetFallbackThemeColors();
            }
        }

        private static IEnumerable<(string Hex, string Label)> GetFallbackThemeColors()
        {
            return new[]
            {
                ("#111111", "Dark 1"),
                ("#FFFFFF", "Light 1"),
                ("#2563EB", "Accent 1"),
                ("#10B981", "Accent 2"),
                ("#F59E0B", "Accent 3"),
                ("#DC2626", "Accent 4")
            };
        }

        private Button CreateThemeColorButton(string hexColor, string label)
        {
            var button = new Button
            {
                Width = 28,
                Height = 28,
                Margin = new Thickness(0, 0, 8, 8),
                Tag = hexColor,
                ToolTip = label,
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hexColor)),
                BorderBrush = (Brush)FindResource("BorderBrush"),
                BorderThickness = new Thickness(1),
                Style = null
            };

            button.Click += PresetColorButton_Click;
            return button;
        }

        private sealed class FilterOption
        {
            public string Value { get; }
            public string Label { get; }

            public FilterOption(string value, string label)
            {
                Value = value;
                Label = label;
            }

            public static FilterOption All()
            {
                return new FilterOption("All", "All");
            }

            public override string ToString()
            {
                return Label;
            }
        }

        private sealed class IconCardViewModel : INotifyPropertyChanged
        {
            public IconCatalogItem Item { get; }
            public string Name => Item.Name;
            public string Theme => ToDisplayLabel(Item.Theme);
            public string StyleLabel => ToDisplayLabel(Item.Style);
            public string CollectionLabel { get; }
            public string RelativePath => Item.RelativePath;
            public string TagsLabel => string.Join(", ", Item.Tags ?? new List<string>());
            public Geometry Geometry => _geometry;
            public bool IsFilled { get; private set; }
            public double StrokeThickness => IsFilled ? 0 : 2;
            public Brush FillBrush { get; private set; }
            public Brush StrokeBrush { get; private set; }
            public string LoadingGlyph => _geometry == null ? "..." : string.Empty;
            public string FavoriteGlyph => Item.IsFavorite ? "*" : "+";
            public Brush FavoriteBrush => Item.IsFavorite ? Brushes.Gold : Brushes.Gray;
            private readonly string _filePath;
            private Geometry _geometry;
            private Brush _currentBrush;

            public event PropertyChangedEventHandler PropertyChanged;

            public IconCardViewModel(IconCatalogItem item, string filePath, Brush selectedBrush, string collectionLabel)
            {
                Item = item;
                _filePath = filePath;
                CollectionLabel = collectionLabel;
                SetColorBrush(selectedBrush);
                RefreshFavoriteState();
            }

            public void SetColorBrush(Brush brush)
            {
                _currentBrush = brush;

                if (IsFilled)
                {
                    FillBrush = brush;
                    StrokeBrush = null;
                }
                else
                {
                    FillBrush = Brushes.Transparent;
                    StrokeBrush = brush;
                }

                NotifyPropertyChanged(nameof(FillBrush));
                NotifyPropertyChanged(nameof(StrokeBrush));
            }

            public void RefreshFavoriteState()
            {
            }

            public IconGeometryLoadResult ParseGeometry()
            {
                try
                {
                    var geometry = SvgParser.ParseLucideSvg(_filePath, out bool isFilled);
                    if (geometry != null && geometry.Children.Count > 0)
                    {
                        if (geometry.CanFreeze)
                        {
                            geometry.Freeze();
                        }

                        return new IconGeometryLoadResult(this, geometry, isFilled);
                    }
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "IconBrowserWindow.IconCardViewModel.ParseGeometry");
                }

                var fallbackGeometry = Geometry.Parse("M4,4H20V20H4Z");
                if (fallbackGeometry.CanFreeze)
                {
                    fallbackGeometry.Freeze();
                }

                return new IconGeometryLoadResult(this, fallbackGeometry, true);
            }

            public void ApplyGeometry(Geometry geometry, bool isFilled)
            {
                _geometry = geometry;
                IsFilled = isFilled;
                SetColorBrush(_currentBrush ?? Brushes.Black);
                NotifyPropertyChanged(nameof(Geometry));
                NotifyPropertyChanged(nameof(IsFilled));
                NotifyPropertyChanged(nameof(StrokeThickness));
                NotifyPropertyChanged(nameof(LoadingGlyph));
            }

            private void NotifyPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        private sealed class IconGeometryLoadResult
        {
            public IconCardViewModel Icon { get; }
            public Geometry Geometry { get; }
            public bool IsFilled { get; }

            public IconGeometryLoadResult(IconCardViewModel icon, Geometry geometry, bool isFilled)
            {
                Icon = icon;
                Geometry = geometry;
                IsFilled = isFilled;
            }
        }
    }
}
