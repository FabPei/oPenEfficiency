using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Threading;
using oPenEfficiency.Features;
using oPenEfficiency.Features.Tables;
using oPenEfficiency.Features.Text;
using oPenEfficiency.Features.Utilities;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.UI.Dialogs;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency
{
    public partial class MainSidebar : UserControl, INotifyPropertyChanged
    {
        private PowerPointManager _powerPointManager;
        private UI.ShapeSizePanel _sizePanel;
        private UI.FontPanel _fontPanel;
        private UI.SearchBarControl _searchBar;
        private UI.SpecialCharactersWindow _specialCharsWindow;
        private PowerPoint.Selection _currentSelection;

        private SidebarConfig _config = new SidebarConfig();
        public SidebarConfig Config
        {
            get => _config;
            set
            {
                _config = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

                private System.Collections.Generic.Dictionary<string, Window> _floatingWindows = new System.Collections.Generic.Dictionary<string, Window>();

        private void ToggleFloatingWindow<T>(string key, Func<T> createAction, System.Windows.Controls.Primitives.ToggleButton toggleBtn) where T : Window
        {
            if (toggleBtn != null && toggleBtn.IsChecked == false)
            {
                if (_floatingWindows.ContainsKey(key))
                {
                    _floatingWindows[key].Close();
                    _floatingWindows.Remove(key);
                }
                return;
            }

            if (!_floatingWindows.ContainsKey(key) || !_floatingWindows[key].IsLoaded)
            {
                var window = createAction();
                _floatingWindows[key] = window;
                if (toggleBtn != null)
                {
                    toggleBtn.IsChecked = true;
                    window.Closed += (s, e) => { 
                        toggleBtn.IsChecked = false; 
                        if (_floatingWindows.ContainsKey(key)) _floatingWindows.Remove(key);
                    };
                }
                window.Show();
            }
            else
            {
                if (toggleBtn != null) toggleBtn.IsChecked = true;
                _floatingWindows[key].Activate();
            }
        }


        public MainSidebar()
        {
            InitializeComponent();
            RegisterShortcuts();

            _resizeTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
            _resizeTimer.Tick += ResizeTimer_Tick;
            
            // Initial Rebuild from Config
            this.Loaded += (s, e) => RebuildUI();

            // Listen for theme changes to refresh local resources
            oPenEfficiency.Services.ThemeManager.ThemeChanged += (s, themeName) => RefreshThemeResources();
        }

        private void RefreshThemeResources()
        {
            try
            {
                // Find and refresh Controls.xaml in this UserControl's resources
                var controlsDict = this.Resources.MergedDictionaries
                    .FirstOrDefault(d => d.Source != null && d.Source.OriginalString.Contains("UI/Styles/Controls.xaml"));

                if (controlsDict != null)
                {
                    this.Resources.MergedDictionaries.Remove(controlsDict);
                    this.Resources.MergedDictionaries.Add(new ResourceDictionary
                    {
                        Source = new Uri("/oPenEfficiency;component/UI/Styles/Controls.xaml", UriKind.RelativeOrAbsolute)
                    });
                }
                
                // Also trigger a rebuild if necessary, or just invalidate
                this.InvalidateVisual();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error refreshing sidebar theme: {ex.Message}");
            }
        }

        private int _lastColumnCount = 0;
        private DispatcherTimer _resizeTimer;
        private string _lastSectionLayoutHash = "";
        private bool _isCurrentlyHorizontal = false;

        /// <summary>
        /// Sets the dock position of the task pane to adjust layout accordingly.
        /// </summary>
        public void SetDockPosition(Office.MsoCTPDockPosition position)
        {
            // Kept for compatibility, but layout is now size-based
        }

        private void UserControl_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // Stop and restart the timer to debounce the rebuild
            _resizeTimer.Stop();
            _resizeTimer.Start();

            // Force invalidation to keep hit-testing in sync with host during resize
            this.InvalidateVisual();
        }

        private int CalculateColumnCount(double width, double iconSize = 16)
        {
            // Calculate columns based on available width
            // Each button needs iconSize + 14 pixels (10 for button width, 4 for spacing)
            int columns = (int)(width / (iconSize + 14));

            // Ensure at least 1 column, and cap at reasonable maximum
            return Math.Max(1, Math.Min(12, columns));
        }

        private void ResizeTimer_Tick(object sender, EventArgs e)
        {
            _resizeTimer.Stop();

            // Only rebuild if the layout would actually change
            bool isWideWindow = this.ActualWidth > (this.ActualHeight * 1.5);

            // Check if layout mode changed (horizontal <-> vertical)
            if (isWideWindow != _isCurrentlyHorizontal)
            {
                _isCurrentlyHorizontal = isWideWindow;
                RebuildUI();
                return;
            }

            if (isWideWindow)
            {
                // For horizontal layout, check if the matrix configuration would change
                string newLayoutHash = ComputeHorizontalLayoutHash();
                if (newLayoutHash != _lastSectionLayoutHash)
                {
                    _lastSectionLayoutHash = newLayoutHash;
                    RebuildUI();
                }
                else
                {
                    // Even if we don't rebuild, ensure layout is updated and synced
                    this.UpdateLayout();
                    this.InvalidateVisual();
                }
            }
            else
            {
                // For vertical layout, check if column count would change
                int newCols = CalculateColumnCount(this.ActualWidth);
                if (newCols != _lastColumnCount)
                {
                    _lastColumnCount = newCols;
                    RebuildUI();
                }
                else
                {
                    // Even if we don't rebuild, ensure layout is updated and synced
                    this.UpdateLayout();
                    this.InvalidateVisual();
                }
            }
        }

        /// <summary>
        /// Computes a hash representing the current horizontal layout configuration.
        /// Changes when section column counts or row counts would change.
        /// </summary>
        private string ComputeHorizontalLayoutHash()
        {
            var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
            if (config == null || config.Sections == null) return "";

            double availableHeight = this.ActualHeight - 40;
            int buttonHeight = 28;
            int maxRows = Math.Max(1, (int)(availableHeight / buttonHeight));
            var widgetIds = new System.Collections.Generic.HashSet<string> { "BtnSizePanel", "BtnFontPanel" };

            var hashParts = new System.Collections.Generic.List<string>();
            foreach (var section in config.Sections)
            {
                if (!section.IsVisible) continue;
                var buttonFeatures = section.Features.Where(f => f != null && !widgetIds.Contains(f.Id)).ToList();
                if (buttonFeatures.Count == 0) continue;

                // Always use minimum 4 columns per section (unless section has fewer buttons)
                int columns = 4;

                // If section has fewer than 4 buttons, use that many columns
                if (buttonFeatures.Count < 4)
                {
                    columns = Math.Max(1, buttonFeatures.Count);
                }
                // If buttons would overflow with 4 columns, increase to 5 or 6
                else
                {
                    int rowsNeededWith4Cols = (buttonFeatures.Count + 4 - 1) / 4;
                    if (rowsNeededWith4Cols > maxRows)
                    {
                        // Try 5 columns
                        int rowsNeededWith5Cols = (buttonFeatures.Count + 5 - 1) / 5;
                        if (rowsNeededWith5Cols <= maxRows)
                        {
                            columns = 5;
                        }
                        else
                        {
                            // Try 6 columns as last resort
                            columns = 6;
                        }
                    }
                }

                int rowsNeeded = (buttonFeatures.Count + columns - 1) / columns;
                hashParts.Add($"{section.Name}:{columns}x{rowsNeeded}");
            }

            return string.Join("|", hashParts);
        }

        /// <summary>
        /// Called from ThisAddIn after the sidebar is created, to give it a
        /// PowerPointManager for the size panel and selection-change wiring.
        /// </summary>
        public void SetManager(PowerPointManager manager)
        {
            _powerPointManager = manager;

            if (_sizePanel == null)
            {
                _sizePanel = new UI.ShapeSizePanel(manager);
            }
            if (_fontPanel == null)
            {
                _fontPanel = new UI.FontPanel(manager);
            }

            // If the sidebar has already rendered (Loaded fired), rebuild so the
            // widgets appear in their configured section position.
            if (DynamicContentPanel != null && DynamicContentPanel.Children.Count > 0)
                RebuildUI();
        }

        /// <summary>Called by ThisAddIn whenever PowerPoint selection changes.</summary>
        public void OnSelectionChange(PowerPoint.Selection sel)
        {
            _currentSelection = sel;
            _sizePanel?.OnSelectionChange(sel);
            _fontPanel?.OnSelectionChange(sel);
            UpdateButtonStates();
        }

        /// <summary>Called by ThisAddIn to update size panel while editing text inside a shape.</summary>
        public void UpdateSizePanel(PowerPoint.Shape shape)
        {
            _sizePanel?.UpdateSize(shape);
        }

        // Current button background brush, set by RebuildUI and used by CreateFeatureButton
        private Brush _currentBtnBgBrush = Brushes.Transparent;

        public void RebuildUI(string searchTerm = "")
        {
            if (DynamicContentPanel == null) return;

            DynamicContentPanel.Children.Clear();

            // Update the current layout mode
            _isCurrentlyHorizontal = this.ActualWidth > (this.ActualHeight * 1.5);

            SidebarConfig config;
            bool isSearchMode = !string.IsNullOrWhiteSpace(searchTerm);

            if (isSearchMode)
            {
                // In search mode, we always show the search bar at the top,
                // followed by matching results.
                var searchBar = GetWidgetForId("BtnSearchBar") as UI.SearchBarControl;
                if (searchBar != null)
                {
                    searchBar.SetSearchText(searchTerm);
                    DynamicContentPanel.Children.Add(searchBar);
                }

                config = new SidebarConfig();
                var searchSection = new SidebarSection { Name = "Search Results", Color = "#10B981" };
                
                string query = searchTerm.ToLower();
                var matches = UI.FeatureLibrary.AllFeatures
                    .Where(f => f.Name.ToLower().Contains(query) || f.Id.ToLower().Contains(query))
                    .Where(f => f.Id != "BtnSearchBar") // Don't match the search bar itself
                    .ToList();

                foreach (var m in matches) searchSection.Features.Add(m);
                config.Sections.Add(searchSection);
            }
            else
            {
                config = JsonService.Load<SidebarConfig>("sidebar_config.json");
                if (config == null || config.Sections == null || config.Sections.Count == 0)
                    config = UI.FeatureLibrary.GetEmergencyDefault();
                else
                    FilterObsoleteFeatures(config);
            }

            // Apply persisted background colour only if it's not the default "theme-tracking" value
            try
            {
                if (!string.IsNullOrWhiteSpace(config.BackgroundColor) && config.BackgroundColor != "Transparent" && config.BackgroundColor != "#252525")
                {
                    this.Background = (SolidColorBrush)new BrushConverter().ConvertFrom(config.BackgroundColor);
                }
                else
                {
                    // Fall back to the DynamicResource defined in XAML
                    this.SetResourceReference(Control.BackgroundProperty, "WindowBackgroundBrush");
                }
            }
            catch { this.SetResourceReference(Control.BackgroundProperty, "WindowBackgroundBrush"); }

            // Resolve section header font colour with safe fallback
            SolidColorBrush sectionFontBrush;
            try
            {
                var fc = string.IsNullOrWhiteSpace(config.SectionFontColor) ? "#888888" : config.SectionFontColor;
                sectionFontBrush = (SolidColorBrush)new BrushConverter().ConvertFrom(fc);
            }
            catch { sectionFontBrush = new SolidColorBrush(Color.FromRgb(0x88, 0x88, 0x88)); }

            // Resolve section header font size
            double sectionFontSize = config.SectionFontSize > 0 ? config.SectionFontSize : 9;

            // Resolve icon size (button size)
            double iconSize = config.IconSize > 0 ? config.IconSize : 16;

            // Resolve button background brush
            try
            {
                var btnBgStr = config.ButtonBackgroundColor;
                _currentBtnBgBrush = string.IsNullOrWhiteSpace(btnBgStr) || btnBgStr == "Transparent"
                    ? Brushes.Transparent
                    : (SolidColorBrush)new BrushConverter().ConvertFrom(btnBgStr);
            }
            catch { _currentBtnBgBrush = Brushes.Transparent; }

            // Feature IDs that render as full-width widgets instead of icon buttons
            var widgetIds = new System.Collections.Generic.HashSet<string> { "BtnSizePanel", "BtnFontPanel", "BtnSearchBar" };

            // Determine layout orientation based on window aspect ratio
            // Wide window (width significantly > height) = horizontal ribbon-like arrangement
            // Tall/square window (height >= width or similar) = vertical stack arrangement
            bool isWideWindow = this.ActualWidth > (this.ActualHeight * 1.5);

            // Update the layout hash for horizontal mode
            if (isWideWindow)
            {
                _lastSectionLayoutHash = ComputeHorizontalLayoutHash();
            }

            if (isWideWindow)
            {
                // Horizontal ribbon-like layout: sections side-by-side, buttons in matrix
                // WrapPanel for all sections including Settings
                var sectionsWrapPanel = new WrapPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Left,
                    VerticalAlignment = VerticalAlignment.Top
                };

                foreach (var section in config.Sections)
                {
                    if (!section.IsVisible) continue;

                    // Separate widget features from button features
                    var buttonFeatures = section.Features
                        .Where(f => f != null && !widgetIds.Contains(f.Id))
                        .ToList();

                    if (buttonFeatures.Count == 0) continue;

                    // Calculate available height for buttons (total height minus header and margins)
                    double availableHeight = this.ActualHeight - 50; // Account for header, section margins, wrapper
                    int buttonHeight = 28; // 26px button + 2px margin
                    int maxRows = Math.Max(1, (int)(availableHeight / buttonHeight));

                    // Always use minimum 3 columns per section (unless section has fewer buttons)
                    int columns = 3;

                    // If section has fewer than 3 buttons, use that many columns
                    if (buttonFeatures.Count < 3)
                    {
                        columns = Math.Max(1, buttonFeatures.Count);
                    }
                    else
                    {
                        // Calculate rows needed for each column count
                        int rowsWith3 = (buttonFeatures.Count + 2) / 3;
                        int rowsWith4 = (buttonFeatures.Count + 3) / 4;
                        int rowsWith5 = (buttonFeatures.Count + 4) / 5;
                        int rowsWith6 = (buttonFeatures.Count + 5) / 6;

                        // Only increase columns if 3 columns would exceed available rows
                        if (rowsWith3 > maxRows)
                        {
                            // Try 4 columns - only if it fits within maxRows
                            if (rowsWith4 <= maxRows)
                            {
                                columns = 4;
                            }
                            // Try 5 columns
                            else if (rowsWith5 <= maxRows)
                            {
                                columns = 5;
                            }
                            // Try 6 columns
                            else if (rowsWith6 <= maxRows)
                            {
                                columns = 6;
                            }
                            // If even 6 columns don't fit, stay with 6 (will scroll)
                            else
                            {
                                columns = 6;
                            }
                        }
                        // Otherwise stay with 3 columns (default)
                    }

                    // Create a section container with border (right side only to avoid double lines)
                    // Use Auto width so section only takes needed space
                    var sectionBorder = new Border
                    {
                        BorderBrush = new SolidColorBrush(Color.FromRgb(85, 85, 85)),
                        BorderThickness = new Thickness(0, 0, 1, 0), // Right border only
                        Margin = new Thickness(4, 4, 0, 4),
                        Padding = new Thickness(8, 0, 8, 0), // Left and right padding for buttons
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalAlignment = VerticalAlignment.Top
                    };

                    var sectionStack = new StackPanel
                    {
                        Orientation = Orientation.Vertical,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Top
                    };

                    // Section header (centered, small)
                    var header = new TextBlock
                    {
                        Text = section.Name.ToUpper(),
                        Style = (Style)FindResource("SectionHeader"),
                        Foreground = sectionFontBrush,
                        FontSize = sectionFontSize,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        Margin = new Thickness(0, 0, 0, 4)
                    };
                    sectionStack.Children.Add(header);

                    // Button matrix grid for this section
                    var buttonGrid = new Grid
                    {
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        VerticalAlignment = VerticalAlignment.Top
                    };

                    // Define columns with equal Star width
                    for (int c = 0; c < columns; c++)
                    {
                        buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    }

                    // Define rows with Auto height (no gap - rows size to content)
                    int rowsNeeded = (buttonFeatures.Count + columns - 1) / columns;
                    for (int r = 0; r < rowsNeeded; r++)
                    {
                        buttonGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                    }

                    int index = 0;
                    foreach (var feature in buttonFeatures)
                    {
                        // Use ToggleButton for features that need toggle state
                        FrameworkElement btn;
                        var featureInfo = UI.FeatureLibrary.GetFeatureInfo(feature.Id);
                        if (featureInfo.IsToggle)
                        {
                            var toggleBtn = CreateToggleButton(feature.Id, isSearchMode ? null : section.Color, iconSize);
                            btn = toggleBtn;
                        }
                        else
                        {
                            btn = CreateFeatureButton(feature.Id, isSearchMode ? null : section.Color, iconSize);
                        }

                        if (btn != null)
                        {
                            btn.Width = iconSize + 10;
                            btn.Height = iconSize + 10;
                            btn.Margin = new Thickness(0);
                            btn.HorizontalAlignment = HorizontalAlignment.Center;
                            btn.VerticalAlignment = VerticalAlignment.Center;

                            int row = index / columns;
                            int col = index % columns;
                            Grid.SetRow(btn, row);
                            Grid.SetColumn(btn, col);
                            buttonGrid.Children.Add(btn);
                            index++;
                        }
                    }
                    sectionStack.Children.Add(buttonGrid);

                    // Set explicit height on section border to fill available height
                    // This makes the separator line stretch from top to bottom
                    double sectionHeight = Math.Max(0, this.ActualHeight - 16);
                    if (sectionHeight > 0)
                    {
                        sectionBorder.Height = sectionHeight;
                    }

                    sectionBorder.Child = sectionStack;
                    sectionsWrapPanel.Children.Add(sectionBorder);
                }

                // Add Settings section as the last item in the WrapPanel (right side)
                var settingsBorder = new Border
                {
                    BorderBrush = new SolidColorBrush(Color.FromRgb(85, 85, 85)),
                    BorderThickness = new Thickness(0, 0, 1, 0), // Right border only
                    Margin = new Thickness(4, 4, 0, 4),
                    Padding = new Thickness(8, 0, 8, 0), // Left and right padding for buttons
                    HorizontalAlignment = HorizontalAlignment.Left,
                    VerticalAlignment = VerticalAlignment.Top
                };

                var settingsStack = new StackPanel
                {
                    Orientation = Orientation.Vertical,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Top
                };

                var settingsHeader = new TextBlock
                {
                    Text = "SETTINGS",
                    Style = (Style)FindResource("SectionHeader"),
                    Foreground = (SolidColorBrush)new BrushConverter().ConvertFrom("#9CA3AF"),
                    FontSize = 9,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 4)
                };
                settingsStack.Children.Add(settingsHeader);

                var settingsBtn = CreateFeatureButton("BtnSettings", null, iconSize);
                if (settingsBtn != null)
                {
                    settingsBtn.Width = iconSize + 10;
                    settingsBtn.Height = iconSize + 10;
                    settingsBtn.Margin = new Thickness(1);
                    settingsBtn.HorizontalAlignment = HorizontalAlignment.Center;
                    settingsStack.Children.Add(settingsBtn);
                }

                // Set explicit height on settings border to fill available height
                double settingsHeight = Math.Max(0, this.ActualHeight - 16);
                if (settingsHeight > 0)
                {
                    settingsBorder.Height = settingsHeight;
                }

                settingsBorder.Child = settingsStack;
                sectionsWrapPanel.Children.Add(settingsBorder);

                // Set explicit height on WrapPanel to prevent parent StackPanel from stretching it
                if (this.ActualHeight > 0)
                {
                    sectionsWrapPanel.Height = this.ActualHeight - 16;
                }

                DynamicContentPanel.Children.Add(sectionsWrapPanel);
                UpdateButtonStates();
                return;
            }

            // Vertical layout (original behavior): sections stacked
            foreach (var section in config.Sections)
            {
                if (!section.IsVisible) continue;

                var header = new TextBlock
                {
                    Text       = section.Name.ToUpper(),
                    Style      = (Style)FindResource("SectionHeader"),
                    Foreground = sectionFontBrush,
                    FontSize   = sectionFontSize
                };
                DynamicContentPanel.Children.Add(header);

                // Separate widget features (full-width) from ordinary button features
                var buttonFeatures = new System.Collections.Generic.List<SidebarFeature>();

                foreach (var feature in section.Features)
                {
                    if (feature == null) continue;

                    if (widgetIds.Contains(feature.Id))
                    {
                        // Render as a full-width control
                        var widget = GetWidgetForId(feature.Id);
                        if (widget != null)
                            DynamicContentPanel.Children.Add(widget);
                    }
                    else
                    {
                        buttonFeatures.Add(feature);
                    }
                }

                // Build the icon button grid for non-widget features
                // Use a fixed column count (5) that only reduces when sidebar gets very narrow
                int columnCount = CalculateColumnCount(this.ActualWidth, iconSize);

                _lastColumnCount = columnCount;

                if (buttonFeatures.Count > 0)
                {
                    var grid = new Grid { Margin = new Thickness(0, 0, 0, 2) };
                    for (int i = 0; i < columnCount; i++)
                        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                    int row = 0, col = 0;
                    foreach (var feature in buttonFeatures)
                    {
                        if (col >= columnCount) { col = 0; row++; }
                        while (grid.RowDefinitions.Count <= row)
                            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(iconSize + 12) });

                        // Use ToggleButton for features that need toggle state
                        FrameworkElement btn;
                        var featureInfo = UI.FeatureLibrary.GetFeatureInfo(feature.Id);
                        if (featureInfo.IsToggle)
                        {
                            var toggleBtn = CreateToggleButton(feature.Id, isSearchMode ? null : section.Color, iconSize);
                            btn = toggleBtn;
                        }
                        else
                        {
                            btn = CreateFeatureButton(feature.Id, isSearchMode ? null : section.Color, iconSize);
                        }

                        if (btn != null)
                        {
                            btn.Width = iconSize + 10;
                            btn.Height = iconSize + 10;
                            Grid.SetRow(btn, row);
                            Grid.SetColumn(btn, col);
                            grid.Children.Add(btn);
                            col++;
                        }
                    }
                    DynamicContentPanel.Children.Add(grid);
                }

                // Separator
                DynamicContentPanel.Children.Add(new Rectangle
                {
                    Height = 1,
                    Fill = new SolidColorBrush(Color.FromRgb(85, 85, 85)),
                    Margin = new Thickness(0, 1, 0, 2)
                });
            }
            AddSettingsButton();
            UpdateButtonStates();
        }

        /// <summary>
        /// Returns the full-width WPF control for a widget feature ID, or null if unknown.
        /// </summary>
        private UIElement GetWidgetForId(string id)
        {
            switch (id)
            {
                case "BtnSizePanel":
                    // Create on demand if SetManager hasn't been called yet
                    if (_sizePanel == null && _powerPointManager != null)
                        _sizePanel = new UI.ShapeSizePanel(_powerPointManager);
                    return _sizePanel;
                case "BtnFontPanel":
                    if (_fontPanel == null && _powerPointManager != null)
                        _fontPanel = new UI.FontPanel(_powerPointManager);
                    return _fontPanel;
                case "BtnSearchBar":
                    if (_searchBar == null)
                    {
                        _searchBar = new UI.SearchBarControl();
                        _searchBar.SearchChanged += (s, term) => RebuildUI(term);
                    }
                    return _searchBar;
                default:
                    return null;
            }
        }

        /// <summary>
        /// Updates the enabled/disabled state of all feature buttons based on current selection.
        /// </summary>
        private void UpdateButtonStates()
        {
            // Update all buttons recursively in all children
            foreach (var child in DynamicContentPanel.Children)
            {
                UpdateButtonsInContainer(child as UIElement);
            }
        }

        /// <summary>
        /// Recursively updates button states in a container element.
        /// </summary>
        private void UpdateButtonsInContainer(UIElement container)
        {
            if (container == null) return;

            if (container is ToggleButton toggleBtn)
            {
                UpdateButtonState(toggleBtn);
                return;
            }

            if (container is Button btn)
            {
                UpdateButtonState(btn);
                return;
            }

            // Handle Panel types (StackPanel, WrapPanel, Grid, etc.)
            if (container is Panel panel)
            {
                foreach (var child in panel.Children)
                {
                    UpdateButtonsInContainer(child as UIElement);
                }
                return;
            }

            // Handle Border (has single Child property)
            if (container is Border border)
            {
                UpdateButtonsInContainer(border.Child as UIElement);
                return;
            }

            // Handle ScrollViewer
            if (container is ScrollViewer scrollViewer)
            {
                UpdateButtonsInContainer(scrollViewer.Content as UIElement);
                return;
            }
        }

        /// <summary>
        /// Updates a single button's enabled state based on feature requirements and current selection.
        /// </summary>
        private void UpdateButtonState(Button btn)
        {
            var info = UI.FeatureLibrary.GetFeatureInfo(btn.Name);
            bool isEnabled = IsFeatureEnabled(info, _currentSelection);
            btn.IsEnabled = isEnabled;
        }

        private void UpdateButtonState(ToggleButton btn)
        {
            var info = UI.FeatureLibrary.GetFeatureInfo(btn.Name);
            bool isEnabled = IsFeatureEnabled(info, _currentSelection);
            btn.IsEnabled = isEnabled;

            if (btn.Name == "BtnShapeLock" || btn.Name == "BtnLockDimensions")
            {
                if (!isEnabled || _currentSelection == null || _currentSelection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    btn.IsChecked = false;
                    return;
                }

                try
                {
                    int lockedCount = 0;
                    int totalCount = _currentSelection.ShapeRange.Count;

                    for (int i = 1; i <= totalCount; i++)
                    {
                        bool isLocked = btn.Name == "BtnShapeLock" 
                            ? ShapeLockingFeature.IsShapeLocked(_currentSelection.ShapeRange[i])
                            : DimensionLockFeature.IsDimensionLocked(_currentSelection.ShapeRange[i]);

                        if (isLocked) lockedCount++;
                    }

                    // For toggle button we can set IsThreeState to true to support null
                    btn.IsThreeState = true;

                    if (lockedCount == totalCount)
                        btn.IsChecked = true;
                    else if (lockedCount == 0)
                        btn.IsChecked = false;
                    else
                        btn.IsChecked = null; // Mixed state
                }
                catch
                {
                    btn.IsChecked = false;
                }
            }
        }

        /// <summary>
        /// Determines if a feature should be enabled based on the current selection.
        /// Checks selection type, shape count, table presence, and hidden shapes requirements.
        /// </summary>
        private bool IsFeatureEnabled(UI.FeatureDisplayInfo info, PowerPoint.Selection selection)
        {
            // Check hidden shapes requirement first (for Show Hidden feature)
            // This check is independent of selection - it checks the entire slide
            if (info.RequiresHiddenShapes)
            {
                bool hasHiddenShapes = false;
                try
                {
                    var app = Globals.ThisAddIn.Application;
                    if (app != null && app.ActiveWindow != null && app.ActiveWindow.View != null)
                    {
                        var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                        if (slide != null)
                        {
                            for (int i = 1; i <= slide.Shapes.Count; i++)
                            {
                                try
                                {
                                    if (slide.Shapes[i].Visible == Microsoft.Office.Core.MsoTriState.msoFalse)
                                    {
                                        hasHiddenShapes = true;
                                        break;
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                }
                catch { }
                if (!hasHiddenShapes) return false;
            }

            // No selection captured yet? Allow features that don't need one.
            if (selection == null)
            {
                return info.MinSelection == 0 && !info.RequiresTable && !info.RequiresHiddenShapes;
            }

            // Check selection type requirement (only if explicitly set to ppSelectionShapes)
            if (info.RequiredType == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes && 
                    selection.Type != PowerPoint.PpSelectionType.ppSelectionText)
                    return false;
            }

            // Check table requirement
            if (info.RequiresTable)
            {
                bool hasTable = false;
                try
                {
                    if (selection.ShapeRange != null)
                    {
                        for (int i = 1; i <= selection.ShapeRange.Count; i++)
                        {
                            if (selection.ShapeRange[i].Type == Microsoft.Office.Core.MsoShapeType.msoTable)
                            {
                                hasTable = true;
                                break;
                            }
                        }
                    }
                }
                catch { }
                if (!hasTable) return false;
            }

            // Check shape count requirements (for shape and text selections)
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes || 
                selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                int count = 0;
                try 
                { 
                    if (selection.ShapeRange != null)
                        count = selection.ShapeRange.Count; 
                } 
                catch { }
                
                // When in text edit mode, there is always effectively 1 shape context
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText && count == 0)
                {
                    count = 1;
                }

                if (info.MinSelection > 0 && count < info.MinSelection)
                    return false;
                if (info.MaxSelection > 0 && count > info.MaxSelection)
                    return false;
            }

            return true;
        }

        private void AddSettingsButton()
        {
            // Vertical layout only - horizontal layout adds Settings inline in RebuildUI
            var header = new TextBlock { Text = "SETTINGS", Style = (Style)FindResource("SectionHeader") };
            DynamicContentPanel.Children.Add(header);

            // Load icon size from config
            var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
            double iconSize = config?.IconSize > 0 ? config.IconSize : 16;

            var stack = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
            
            var btn = CreateFeatureButton("BtnSettings", null, iconSize);
            btn.Width = iconSize + 10;
            btn.Height = iconSize + 10;
            btn.Margin = new Thickness(0, 0, 4, 0);
            stack.Children.Add(btn);

            var exploreBtn = CreateFeatureButton("BtnExploreFeatures", null, iconSize);
            exploreBtn.Width = iconSize + 10;
            exploreBtn.Height = iconSize + 10;
            stack.Children.Add(exploreBtn);

            DynamicContentPanel.Children.Add(stack);
        }

        private void FilterObsoleteFeatures(SidebarConfig config)
        {
            if (config?.Sections == null) return;
            
            // Features to completely remove
            var obsoleteIds = new System.Collections.Generic.HashSet<string> { 
                "BtnSwapHorizontal", "BtnSwapVertical"
            };
            
            bool changed = false;

            foreach (var section in config.Sections)
            {
                if (section.Features == null) continue;
                
                for (int i = 0; i < section.Features.Count; i++)
                {
                    var f = section.Features[i];
                    if (f != null)
                    {
                        // Map legacy IDs to new IDs
                        if (f.Id == "BtnAlignToTable")
                        {
                            f.Id = "BtnAlignToTableCell";
                            f.Name = "Align to Table Cell";
                            changed = true;
                        }
                        
                        // Update names for existing IDs to reflect new functionality
                        if (f.Id == "BtnReplaceFont" && f.Name != "Bulk Format Replacer")
                        {
                            f.Name = "Bulk Format Replacer";
                            changed = true;
                        }
                    }
                }
                
                int removedCount = section.Features.RemoveAll(f => f == null || obsoleteIds.Contains(f.Id));
                if (removedCount > 0) changed = true;
            }

            if (changed)
            {
                JsonService.Save(config, "sidebar_config.json");
            }
        }

        private Button CreateFeatureButton(string id, string overrideColor = null, double iconSize = 22)
        {
            var info = UI.FeatureLibrary.GetFeatureInfo(id);
            var btn = new Button
            {
                Name = id,
                Style = (Style)FindResource("GridIconButton")
            };

            // Enhanced Tooltip with description after delay
            var tooltipPanel = new StackPanel { MaxWidth = 250, Margin = new Thickness(2) };
            tooltipPanel.Children.Add(new TextBlock
            {
                Text = info.Tooltip,
                FontWeight = FontWeights.Bold,
                FontSize = 12,
                Foreground = (Brush)FindResource("TextPrimaryBrush"),
                Margin = new Thickness(0, 0, 0, 4)
            });

            if (!string.IsNullOrEmpty(info.Description))
            {
                var descBlock = new TextBlock
                {
                    Text = info.Description,
                    TextWrapping = TextWrapping.Wrap,
                    FontSize = 11,
                    Foreground = (Brush)FindResource("TextSecondaryBrush"),
                    LineHeight = 14,
                    Visibility = Visibility.Collapsed
                };

                // Add a small spacer/separator that toggles with description
                var separator = new Border
                {
                    Height = 1,
                    Background = (Brush)FindResource("BorderBrush"),
                    Margin = new Thickness(0, 2, 0, 4),
                    Visibility = Visibility.Collapsed
                };

                tooltipPanel.Children.Add(separator);
                tooltipPanel.Children.Add(descBlock);

                // Show description and separator after a longer delay
                var timer = new System.Windows.Threading.DispatcherTimer();
                timer.Interval = TimeSpan.FromSeconds(0.5);
                timer.Tick += (s, e) =>
                {
                    descBlock.Visibility = Visibility.Visible;
                    separator.Visibility = Visibility.Visible;
                    timer.Stop();
                };

                btn.MouseEnter += (s, e) => timer.Start();
                btn.MouseLeave += (s, e) =>
                {
                    timer.Stop();
                    descBlock.Visibility = Visibility.Collapsed;
                    separator.Visibility = Visibility.Collapsed;
                };
            }

            btn.ToolTip = tooltipPanel;
            
            var containerGrid = new Grid
            {
                Width = iconSize + 4,
                Height = iconSize + 4,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };

            try
            {
                var hexColor = !string.IsNullOrEmpty(overrideColor) ? overrideColor : info.Color;
                var brush = (Brush)new BrushConverter().ConvertFrom(hexColor);

                // 1. Get Base Icon
                bool isFilled;
                var baseGeom = IconService.GetIconGeometry(id, out isFilled);
                
                // 2. Try to get Hover Icon
                bool isHoverFilled;
                var hoverGeom = IconService.GetIconGeometry(id, true, out isHoverFilled);

                if (baseGeom != null)
                {
                    var baseViewbox = CreateIconViewbox(baseGeom, brush, isFilled, iconSize);
                    containerGrid.Children.Add(baseViewbox);

                    if (hoverGeom != null)
                    {
                        var hoverViewbox = CreateIconViewbox(hoverGeom, brush, isHoverFilled, iconSize);
                        hoverViewbox.Visibility = Visibility.Hidden;
                        containerGrid.Children.Add(hoverViewbox);

                        // Swap icons on hover using a simple event-based approach
                        btn.MouseEnter += (s, e) =>
                        {
                            baseViewbox.Visibility = Visibility.Hidden;
                            hoverViewbox.Visibility = Visibility.Visible;
                        };
                        btn.MouseLeave += (s, e) =>
                        {
                            baseViewbox.Visibility = Visibility.Visible;
                            hoverViewbox.Visibility = Visibility.Hidden;
                        };
                    }
                }
                else
                {
                    // Fallback for missing icon
                    containerGrid.Children.Add(new TextBlock { Text = "?", Foreground = brush, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center });
                }
            }
            catch
            {
                containerGrid.Children.Add(new TextBlock { Text = "!" });
            }

            btn.Background = _currentBtnBgBrush;
            btn.Content = containerGrid;

            btn.Click += (s, e) => ExecuteFeature(id, btn);
            AttachStandardContextMenu(btn);

            return btn;
        }

        private Viewbox CreateIconViewbox(Geometry geometry, Brush brush, bool isFilled, double iconSize)
        {
            var path = new Path
            {
                Data = geometry,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Stretch = Stretch.Uniform
            };

            if (isFilled)
            {
                path.Fill = brush;
            }
            else
            {
                path.Stroke = brush;
                path.StrokeThickness = 2;
                path.StrokeLineJoin = PenLineJoin.Round;
                path.StrokeStartLineCap = PenLineCap.Round;
                path.StrokeEndLineCap = PenLineCap.Round;
                path.Fill = Brushes.Transparent;
            }

            return new Viewbox
            {
                Child = path,
                MaxHeight = iconSize + 4,
                MaxWidth = iconSize + 4,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Stretch = Stretch.Uniform
            };
        }

        private ToggleButton CreateToggleButton(string id, string overrideColor = null, double iconSize = 22)
        {
            var info = UI.FeatureLibrary.GetFeatureInfo(id);
            var toggleBtn = new ToggleButton
            {
                Name = id,
                Style = (Style)FindResource("GridIconButtonToggleButton")
            };

            // Enhanced Tooltip with description after delay
            var tooltipPanel = new StackPanel { MaxWidth = 250, Margin = new Thickness(2) };
            tooltipPanel.Children.Add(new TextBlock
            {
                Text = info.Tooltip,
                FontWeight = FontWeights.Bold,
                FontSize = 12,
                Foreground = (Brush)FindResource("TextPrimaryBrush"),
                Margin = new Thickness(0, 0, 0, 4)
            });

            if (!string.IsNullOrEmpty(info.Description))
            {
                var descBlock = new TextBlock
                {
                    Text = info.Description,
                    TextWrapping = TextWrapping.Wrap,
                    FontSize = 11,
                    Foreground = (Brush)FindResource("TextSecondaryBrush"),
                    LineHeight = 14,
                    Visibility = Visibility.Collapsed
                };

                var separator = new Border
                {
                    Height = 1,
                    Background = (Brush)FindResource("BorderBrush"),
                    Margin = new Thickness(0, 2, 0, 4),
                    Visibility = Visibility.Collapsed
                };

                tooltipPanel.Children.Add(separator);
                tooltipPanel.Children.Add(descBlock);

                var timer = new System.Windows.Threading.DispatcherTimer();
                timer.Interval = TimeSpan.FromSeconds(0.5);
                timer.Tick += (s, e) =>
                {
                    descBlock.Visibility = Visibility.Visible;
                    separator.Visibility = Visibility.Visible;
                    timer.Stop();
                };

                toggleBtn.MouseEnter += (s, e) => timer.Start();
                toggleBtn.MouseLeave += (s, e) =>
                {
                    timer.Stop();
                    descBlock.Visibility = Visibility.Collapsed;
                    separator.Visibility = Visibility.Collapsed;
                };
            }

            toggleBtn.ToolTip = tooltipPanel;

            var containerGrid = new Grid
            {
                Width = iconSize + 4,
                Height = iconSize + 4,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };

            try
            {
                var hexColor = !string.IsNullOrEmpty(overrideColor) ? overrideColor : info.Color;
                var brush = (Brush)new BrushConverter().ConvertFrom(hexColor);

                // 1. Get Base Icon
                bool isFilled;
                var baseGeom = IconService.GetIconGeometry(id, out isFilled);
                
                // 2. Try to get Hover Icon
                bool isHoverFilled;
                var hoverGeom = IconService.GetIconGeometry(id, true, out isHoverFilled);

                if (baseGeom != null)
                {
                    var baseViewbox = CreateIconViewbox(baseGeom, brush, isFilled, iconSize);
                    containerGrid.Children.Add(baseViewbox);

                    if (hoverGeom != null)
                    {
                        var hoverViewbox = CreateIconViewbox(hoverGeom, brush, isHoverFilled, iconSize);
                        hoverViewbox.Visibility = Visibility.Hidden;
                        containerGrid.Children.Add(hoverViewbox);

                        // Swap icons on hover using a simple event-based approach
                        toggleBtn.MouseEnter += (s, e) =>
                        {
                            baseViewbox.Visibility = Visibility.Hidden;
                            hoverViewbox.Visibility = Visibility.Visible;
                        };
                        toggleBtn.MouseLeave += (s, e) =>
                        {
                            baseViewbox.Visibility = Visibility.Visible;
                            hoverViewbox.Visibility = Visibility.Hidden;
                        };
                    }
                }
                else
                {
                    containerGrid.Children.Add(new TextBlock { Text = "?", Foreground = brush, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center });
                }
            }
            catch
            {
                containerGrid.Children.Add(new TextBlock { Text = "!" });
            }

            toggleBtn.Background = _currentBtnBgBrush;
            toggleBtn.Content = containerGrid;

            // Sync toggle state with SnapManager on click
            if (id == "BtnSnapToObjects" || id == "BtnSnapToGrid")
            {
                toggleBtn.Checked += (s, e) =>
                {
                    if (Globals.ThisAddIn.SnapManager != null && !Globals.ThisAddIn.SnapManager.IsEnabled)
                    {
                        Globals.ThisAddIn.SnapManager.Toggle();
                    }
                };
                toggleBtn.Unchecked += (s, e) =>
                {
                    if (Globals.ThisAddIn.SnapManager != null && Globals.ThisAddIn.SnapManager.IsEnabled)
                    {
                        Globals.ThisAddIn.SnapManager.Toggle();
                    }
                };
            }
            else
            {
                toggleBtn.Click += (s, e) => 
                {
                    ExecuteFeature(id, toggleBtn);
                    UpdateButtonStates();
                };
            }
            AttachStandardContextMenu(toggleBtn);

            return toggleBtn;
        }

        private void ExecuteFeature(string id, FrameworkElement targetBtn = null)
        {
            var manager = GetManager();
            if (manager == null)
            {
                System.Diagnostics.Debug.WriteLine($"ExecuteFeature: Manager is null for feature {id}");
                return;
            }

            // Try auto-discovered features first (for features with [FeatureMetadata] attribute)
            var feature = FeatureDiscovery.GetFeature(id);
            if (feature != null)
            {
                var toggleBtn = targetBtn as System.Windows.Controls.Primitives.ToggleButton;
                bool handled = feature.Execute(manager, toggleBtn);
                if (handled) return;
            }

            // Fallback to manual switch for non-migrated features or features requiring UI context
            switch (id)
            {
                case "BtnAlignLeft": AlignToFirstLeftFeature.Execute(manager); break;
                case "BtnAlignCenter": AlignShapesFeature.Execute(manager, "center"); break;
                case "BtnAlignRight": AlignToFirstRightFeature.Execute(manager); break;
                case "BtnAlignTop": AlignToFirstTopFeature.Execute(manager); break;
                case "BtnAlignBottom": AlignToFirstBottomFeature.Execute(manager); break;
                case "BtnAlignCenterVertical": AlignToFirstCenterVerticalFeature.Execute(manager); break;
                case "BtnMatchSize": MatchSizeFeature.Execute(manager); break;
                case "BtnMatchWidth": MatchWidthToFirstFeature.Execute(manager); break;
                case "BtnMatchHeight": MatchHeightToFirstFeature.Execute(manager); break;
                case "BtnGlassHide": GlassHideFeature.Execute(manager); break;
                case "BtnMultiSwap": MultiSwapFeature.Execute(manager); break;
                case "BtnSwapPositions": SwapPositionsFeature.Execute(manager); break;
                case "BtnStretchLeft": StretchToFirstLeftFeature.Execute(manager); break;
                case "BtnStretchRight": StretchToFirstRightFeature.Execute(manager); break;
                case "BtnStretchTop": StretchToFirstTopFeature.Execute(manager); break;
                case "BtnStretchBottom": StretchToFirstBottomFeature.Execute(manager); break;
                case "BtnSplitShape": ToggleFloatingWindow("BtnSplitShape", () => new UI.SplitShapeWindow(manager), targetBtn as System.Windows.Controls.Primitives.ToggleButton); break;
                case "BtnSnapShape": ToggleFloatingWindow("BtnSnapShape", () => new UI.ObjectConnectorWindow(manager), targetBtn as System.Windows.Controls.Primitives.ToggleButton); break;
                case "BtnStretchLeftEdge": StretchToFirstLeftEdgeFeature.Execute(manager); break;
                case "BtnStretchRightEdge": StretchToFirstRightEdgeFeature.Execute(manager); break;
                case "BtnStretchTopEdge": StretchToFirstTopEdgeFeature.Execute(manager); break;
                case "BtnStretchBottomEdge": StretchToFirstBottomEdgeFeature.Execute(manager); break;
                case "BtnBringToFront": BringToFrontFeature.Execute(manager); break;
                case "BtnSendToBack": SendToBackFeature.Execute(manager); break;
                case "BtnBringForward": BringForwardFeature.Execute(manager); break;
                case "BtnSendBackward": SendBackwardFeature.Execute(manager); break;
                case "BtnPickColor": PickColorFeature.Execute(manager); break;
                case "BtnPickFillColor": PickFillColorFeature.Execute(manager); break;
                case "BtnPickLineColor": PickLineColorFeature.Execute(manager); break;
                case "BtnPickTextColor": PickTextColorFeature.Execute(manager); break;
                case "BtnDockLeft": DockToFirstLeftFeature.Execute(manager); break;
                case "BtnDockRight": DockToFirstRightFeature.Execute(manager); break;
                case "BtnDockTop": DockToFirstTopFeature.Execute(manager); break;
                case "BtnDockBottom": DockToFirstBottomFeature.Execute(manager); break;
                case "BtnRectifyRotation": RectifyRotationFeature.Execute(manager); break;
                case "BtnArrangeCircle": ArrangeInShapeFeature.Execute(manager, "circle"); break;
                case "BtnArrangePro": ToggleFloatingWindow("BtnArrangePro", () => new UI.ArrangeProWindow(manager), targetBtn as System.Windows.Controls.Primitives.ToggleButton); break;
                case "BtnShapeLock": ShapeLockingFeature.SetSelectionShapeLock(manager, null); break;
                case "BtnLockDimensions": DimensionLockFeature.SetSelectionDimensionLock(manager, null); break;
                case "BtnShapeToggleAll": ShapeLockingFeature.SetGlobalShapeLock(manager, null); break;
                case "BtnShapeLockAll": ShapeLockingFeature.SetGlobalShapeLock(manager, true); break;
                case "BtnShapeUnlockAll": ShapeLockingFeature.SetGlobalShapeLock(manager, false); break;
                case "BtnAgendaWizard": AgendaWizardFeature.Execute(manager); break;
                case "BtnSaveAgendaLayout": SaveAgendaLayoutFeature.Execute(manager); break;
                case "BtnSeriesGenerator": oPenEfficiency.Features.Utilities.SeriesGeneratorFeature.Execute(manager); break;
                case "BtnDiagonalAlign": oPenEfficiency.Features.Alignment.DiagonalAlignFeature.Execute(manager); break;
                case "BtnStickyNote": StickyNoteFeature.Execute(manager); break;
                case "BtnSlideNotes": 
                    var toggleSlideNotes = targetBtn as System.Windows.Controls.Primitives.ToggleButton;
                    oPenEfficiency.Features.Utilities.SlideNotesFeature.Execute(manager, toggleSlideNotes);
                    break;
                case "BtnManageStickyNotes": StickyNoteManagerFeature.Execute(manager, StickyNoteManagerFeature.ActionType.MoveOff, StickyNoteManagerFeature.Scope.ThisSlide); break;
                case "BtnExportWizard": ToggleFloatingWindow("BtnExportWizard", () => new UI.ExportWizard(), targetBtn as System.Windows.Controls.Primitives.ToggleButton); break;
                case "BtnQRCode": ToggleFloatingWindow("BtnQRCode", () => new UI.QRCodeWindow(manager), targetBtn as System.Windows.Controls.Primitives.ToggleButton); break;

                case "BtnThemeColor": ThemeColorPickerFeature.Execute(manager); break;
                
                case "BtnAddSticker": AddStickerFeature.Execute(manager); break;
                case "BtnMagicResizer": ToggleFloatingWindow("BtnMagicResizer", () => new UI.MagicResizerWindow(), targetBtn as System.Windows.Controls.Primitives.ToggleButton); break;
                case "BtnSpecialChars": 
                    OpenSpecialCharactersWindow(null);
                    break;
                case "BtnThermometer": ThermometerFeature.Execute(manager); break;
                case "BtnSelectSameType": SelectSameTypeFeature.Execute(manager); break;
                case "BtnSnapToObjects":
                    var toggleBtn = targetBtn as System.Windows.Controls.Primitives.ToggleButton;
                    SnapToObjectsFeature.Execute(manager, toggleBtn);
                    break;
                case "BtnSnapToGrid":
                    var gridToggleBtn = targetBtn as System.Windows.Controls.Primitives.ToggleButton;
                    SnapToGridFeature.Execute(manager, gridToggleBtn);
                    break;
                case "BtnReplaceFont": BulkFormatReplacerFeature.Execute(manager); break;
                case "BtnFormatNumbers": oPenEfficiency.Features.FormatNumbersFeature.Execute(manager); break;
                case "BtnStyleCheck": oPenEfficiency.Features.StyleCheckFeature.ExecuteDialog(manager, targetBtn as System.Windows.Controls.Primitives.ToggleButton); break;
                case "BtnTransparentColor": oPenEfficiency.Features.TransparentColorFeature.Execute(manager); break;
                case "BtnTagInspector": oPenEfficiency.Features.Utilities.TagInspectorFeature.Execute(manager); break;
                case "BtnAlignToTableCell": 
                case "BtnAlignToTable": AlignToTableCellFeature.Execute(manager); break;
                case "BtnRepeatShape": RepeatShapeFeature.Execute(manager, RepeatShapeFeature.RepeatDirection.Right); break;
                
                case "BtnRemoveHorizontalGaps": RemoveHorizontalGapsFeature.Execute(manager); break;
                case "BtnRemoveVerticalGaps": RemoveVerticalGapsFeature.Execute(manager); break;
                case "BtnIncreaseHorizontalSpacing": AdjustHorizontalSpacingFeature.Execute(manager, GetSpacingIncrement()); break;
                case "BtnDecreaseHorizontalSpacing": AdjustHorizontalSpacingFeature.Execute(manager, -GetSpacingIncrement()); break;
                case "BtnIncreaseVerticalSpacing": AdjustVerticalSpacingFeature.Execute(manager, GetSpacingIncrement()); break;
                case "BtnDecreaseVerticalSpacing": AdjustVerticalSpacingFeature.Execute(manager, -GetSpacingIncrement()); break;
                case "BtnUpdateExcelCharts":
                case "BtnSlideGuidelines":
                case "BtnAdjustSpacing":
                case "BtnTableSort": 
                case "BtnAlignText":
                //case "BtnSmartCorners":
                    // Open context menu on click
                    if (targetBtn != null && targetBtn.ContextMenu != null) targetBtn.ContextMenu.IsOpen = true;
                    break;
                case "BtnSlidePaste":
                    ToggleFloatingWindow("BtnSlidePaste", () => new UI.SlidePasteWindow(manager), targetBtn as System.Windows.Controls.Primitives.ToggleButton);
                    break;
                case "BtnTableBranding":
                    var brandingDialog = new TableBrandingDialog();
                    if (brandingDialog.ShowDialog() == true)
                    {
                        if (brandingDialog.Action == "Save" && !string.IsNullOrEmpty(brandingDialog.StyleToSave))
                        {
                            TableBrandingFeature.SaveTableStyle(manager, brandingDialog.StyleToSave);
                        }
                        else if (brandingDialog.Action == "Apply" && brandingDialog.SelectedStyle != null)
                        {
                            TableBrandingFeature.ApplyTableStyle(manager, brandingDialog.SelectedStyle);
                        }
                    }
                    break;
                case "BtnSyncObjects": SyncObjectsFeature.Execute(manager); break;
                case "BtnWinnerPicker": oPenEfficiency.Features.LotteryFeature.ExecuteWithToggle(manager, targetBtn as System.Windows.Controls.Primitives.ToggleButton); break;
                case "BtnNumeration": oPenEfficiency.Features.NumerationFeature.Create(manager); break;
                case "BtnProgressSeries": oPenEfficiency.Features.ProgressSeriesFeature.Create(manager); break;
                case "BtnCleaner": // By default, maybe do nothing or show context menu? 
                case "BtnSettings": OpenSettings(); break;
                case "BtnConvertTableToShapes": ConvertTableToShapesFeature.Execute(manager); break;
                case "BtnInsertColumnLeft": InsertTableColumnLeftFeature.Execute(manager, false); break;
                case "BtnInsertColumnRight": InsertTableColumnRightFeature.Execute(manager, false); break;
                case "BtnInsertRowTop": InsertTableRowTopFeature.Execute(manager, false); break;
                case "BtnInsertRowBottom": InsertTableRowBottomFeature.Execute(manager, false); break;
                case "BtnTableTranspose": TableTransposeFeature.Execute(manager); break;
                case "BtnTranslate": TranslationFeatures.Execute(manager); break;
                case "BtnTableSplit": TableSplitFeature.Execute(manager, true); break;
                case "BtnTableSum": TableSumFeature.Execute(manager, true, false); break;
                case "BtnToggleSidebar": Globals.ThisAddIn.ToggleSidebar(); break;
                case "BtnTableDimensions":
                    if (targetBtn != null && targetBtn.ContextMenu != null) targetBtn.ContextMenu.IsOpen = true;
                    break;
                case "BtnApplyTextTool":
                    ToggleFloatingWindow("BtnApplyTextTool", () => new UI.ApplyTextWindow(manager), targetBtn as System.Windows.Controls.Primitives.ToggleButton);
                    break;
            }
        }



        private void OpenSettings()
        {
            var win = new UI.SettingsWindow(GetManager());
            win.SettingsApplied += (s, args) => RebuildUI();
            win.ShowDialog();
        }

        private float GetSpacingIncrement()
        {
            var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
            return (float)(config?.SpacingIncrement ?? 10.0);
        }

        private PowerPointManager GetManager()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null) 
                {
                    System.Diagnostics.Debug.WriteLine("GetManager: Globals.ThisAddIn.Application is null");
                    return null;
                }
                return _powerPointManager ?? (_powerPointManager = new PowerPointManager(app));
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetManager COMException: {ex.Message}");
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetManager Exception: {ex.Message}");
                return null;
            }
        }

        private void RegisterShortcuts()
        {
            foreach (var feature in UI.FeatureLibrary.AllFeatures)
            {
                string id = feature.Id;
                ShortcutManager.RegisterHandler(id, () => ExecuteFeature(id, null));
            }
            ShortcutManager.LoadShortcuts();
        }

        /// <summary>
        /// Attach context menu to a feature button. Works for both Button and ToggleButton.
        /// Only BtnSnapToObjects and BtnSnapToGrid use ToggleButton; all others use Button.
        /// </summary>
        private void AttachStandardContextMenu(System.Windows.Controls.Primitives.ButtonBase btn)
        {
            btn.ContextMenu = new ContextMenu();
            var item = new MenuItem { Header = "Configure Shortcuts" };
            item.Click += (s, e) => {
                var win = new ShortcutConfigWindow(btn.Name);
                win.ShowDialog();
            };
            btn.ContextMenu.Items.Add(item);

            if (btn.Name == "BtnAlignLeft" || btn.Name == "BtnAlignRight" || btn.Name == "BtnAlignTop" || btn.Name == "BtnAlignBottom")
            {
                btn.ContextMenu.Items.Add(new Separator());

                string direction, label;
                switch (btn.Name)
                {
                    case "BtnAlignLeft":   direction = "left";   label = "Move Left Until Object";  break;
                    case "BtnAlignRight":  direction = "right";  label = "Move Right Until Object"; break;
                    case "BtnAlignTop":    direction = "top";    label = "Move Up Until Object";    break;
                    default:               direction = "bottom"; label = "Move Down Until Object";  break;
                }

                var miCollision = new MenuItem { Header = label };
                var capturedDir = direction;
                miCollision.Click += (s, e) =>
                {
                    var mgr = GetManager();
                    var app = mgr?.GetApplication();
                    if (app?.ActiveWindow?.Selection?.Type != Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes) return;
                    var sr = app.ActiveWindow.Selection.ShapeRange;
                    if (sr == null || sr.Count == 0) return;
                    for (int i = 1; i <= sr.Count; i++)
                        Features.AlignmentHelpers.MoveUntilCollision(app, sr[i], capturedDir);
                };
                btn.ContextMenu.Items.Add(miCollision);
            }
            else if (btn.Name == "BtnSwapPositions")
            {
                var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
                if (config != null && !config.EnableSwapPositionsContextMenu) return;

                btn.ContextMenu.Items.Add(new Separator());

                var anchors = new[] {
                    ("Top Left", SwapPositionsFeature.SwapAnchor.TopLeft),
                    ("Top Center", SwapPositionsFeature.SwapAnchor.Top),
                    ("Top Right", SwapPositionsFeature.SwapAnchor.TopRight),
                    ("Left Center", SwapPositionsFeature.SwapAnchor.Left),
                    ("Center", SwapPositionsFeature.SwapAnchor.Center),
                    ("Right Center", SwapPositionsFeature.SwapAnchor.Right),
                    ("Bottom Left", SwapPositionsFeature.SwapAnchor.BottomLeft),
                    ("Bottom Center", SwapPositionsFeature.SwapAnchor.Bottom),
                    ("Bottom Right", SwapPositionsFeature.SwapAnchor.BottomRight)
                };

                var miList = new System.Collections.Generic.List<MenuItem>();
                foreach (var anchorInfo in anchors)
                {
                    var mi = new MenuItem { Header = anchorInfo.Item1, IsCheckable = true };
                    var anchorVal = anchorInfo.Item2;
                    mi.Tag = anchorVal;
                    
                    btn.ContextMenu.Opened += (s, e) => {
                        mi.IsChecked = SwapPositionsFeature.DefaultAnchor == (SwapPositionsFeature.SwapAnchor)mi.Tag;
                    };

                    mi.Click += (s, e) => {
                        SwapPositionsFeature.DefaultAnchor = anchorVal;
                        foreach (var m in miList) m.IsChecked = ((SwapPositionsFeature.SwapAnchor)m.Tag == anchorVal);
                        SwapPositionsFeature.Execute(GetManager(), anchorVal);
                    };
                    miList.Add(mi);
                    btn.ContextMenu.Items.Add(mi);
                }
                btn.ContextMenu.Items.Add(new Separator());
            }
            else if (btn.Name == "BtnStoryline")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var miCopy = new MenuItem { Header = "Copy Storyline to Clipboard" };
                miCopy.Click += (s, e) => {
                    var mgr = GetManager();
                    var titles = StorylineFeature.ExtractStoryline(mgr);
                    var sb = new System.Text.StringBuilder();
                    foreach (var storyLineItem in titles) sb.AppendLine($"{storyLineItem.SlideNumber}: {storyLineItem.TitleText}");
                    System.Windows.Clipboard.SetText(sb.ToString());
                    System.Windows.MessageBox.Show("Storyline copied to clipboard.");
                };
                btn.ContextMenu.Items.Add(miCopy);
            }
            else if (btn.Name == "BtnPasteXY")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var miCoords = new MenuItem { Header = "Paste Coordinates (X/Y)" };
                miCoords.Click += (s, e) => PasteCoordinatesFeature.Execute(GetManager(), PasteCoordinatesFeature.PasteMode.Coordinates);
                btn.ContextMenu.Items.Add(miCoords);

                var miSize = new MenuItem { Header = "Paste Size (Width/Height)" };
                miSize.Click += (s, e) => PasteCoordinatesFeature.Execute(GetManager(), PasteCoordinatesFeature.PasteMode.Size);
                btn.ContextMenu.Items.Add(miSize);

                var miBoth = new MenuItem { Header = "Paste Both (Coordinates & Size)" };
                miBoth.Click += (s, e) => PasteCoordinatesFeature.Execute(GetManager(), PasteCoordinatesFeature.PasteMode.Both);
                btn.ContextMenu.Items.Add(miBoth);
            }
            else if (btn.Name == "BtnSlideGuidelines")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var mi1 = new MenuItem { Header = "3-Column Layout" };
                mi1.Click += (s, e) => SlideGuidelinesFeature.Execute(GetManager(), SlideGuidelinesFeature.GuidelineLayout.ThreeColumn);
                btn.ContextMenu.Items.Add(mi1);

                var mi2 = new MenuItem { Header = "4-Column Layout" };
                mi2.Click += (s, e) => SlideGuidelinesFeature.Execute(GetManager(), SlideGuidelinesFeature.GuidelineLayout.FourColumn);
                btn.ContextMenu.Items.Add(mi2);

                var mi3 = new MenuItem { Header = "5-Column Layout" };
                mi3.Click += (s, e) => SlideGuidelinesFeature.Execute(GetManager(), SlideGuidelinesFeature.GuidelineLayout.FiveColumn);
                btn.ContextMenu.Items.Add(mi3);

                var mi4 = new MenuItem { Header = "1:2 Column Layout (Asymmetrical)" };
                mi4.Click += (s, e) => SlideGuidelinesFeature.Execute(GetManager(), SlideGuidelinesFeature.GuidelineLayout.OneTwoColumn);
                btn.ContextMenu.Items.Add(mi4);

                btn.ContextMenu.Items.Add(new Separator());
                var miThirds = new MenuItem { Header = "Rule of Thirds (3x3 Composition Grid)" };
                miThirds.Click += (s, e) => SlideGuidelinesFeature.Execute(GetManager(), SlideGuidelinesFeature.GuidelineLayout.RuleOfThirds);
                btn.ContextMenu.Items.Add(miThirds);

                btn.ContextMenu.Items.Add(new Separator());
                var miDefault = new MenuItem { Header = "2-Column Layout" };
                miDefault.Click += (s, e) => SlideGuidelinesFeature.Execute(GetManager(), SlideGuidelinesFeature.GuidelineLayout.Default);
                btn.ContextMenu.Items.Add(miDefault);
            }
            else if (btn.Name == "BtnCleaner")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var miSlides = new MenuItem { Header = "Remove Hidden Slides" };
                miSlides.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveHiddenSlides);
                btn.ContextMenu.Items.Add(miSlides);

                var miAnim = new MenuItem { Header = "Remove Animations" };
                var miAnimAll = new MenuItem { Header = "From All Slides" };
                miAnimAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveAnimationsAll);
                var miAnimSel = new MenuItem { Header = "From Selected Slides" };
                miAnimSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveAnimationsSelected);
                miAnim.Items.Add(miAnimAll);
                miAnim.Items.Add(miAnimSel);
                btn.ContextMenu.Items.Add(miAnim);

                var miTrans = new MenuItem { Header = "Remove Transitions" };
                var miTransAll = new MenuItem { Header = "From All Slides" };
                miTransAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveTransitionsAll);
                var miTransSel = new MenuItem { Header = "From Selected Slides" };
                miTransSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveTransitionsSelected);
                miTrans.Items.Add(miTransAll);
                miTrans.Items.Add(miTransSel);
                btn.ContextMenu.Items.Add(miTrans);

                var miNotes = new MenuItem { Header = "Remove Speaker Notes" };
                var miNotesAll = new MenuItem { Header = "From All Slides" };
                miNotesAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveSpeakerNotesAll);
                var miNotesSel = new MenuItem { Header = "From Selected Slides" };
                miNotesSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveSpeakerNotesSelected);
                miNotes.Items.Add(miNotesAll);
                miNotes.Items.Add(miNotesSel);
                btn.ContextMenu.Items.Add(miNotes);

                var miDevNotes = new MenuItem { Header = "Remove Developer Notes" };
                var miDevNotesAll = new MenuItem { Header = "From All Slides" };
                miDevNotesAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveDeveloperNotesAll);
                var miDevNotesSel = new MenuItem { Header = "From Selected Slides" };
                miDevNotesSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveDeveloperNotesSelected);
                miDevNotes.Items.Add(miDevNotesAll);
                miDevNotes.Items.Add(miDevNotesSel);
                btn.ContextMenu.Items.Add(miDevNotes);

                var miComm = new MenuItem { Header = "Remove Comments" };
                var miCommAll = new MenuItem { Header = "From All Slides" };
                miCommAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveCommentsAll);
                var miCommSel = new MenuItem { Header = "From Selected Slides" };
                miCommSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveCommentsSelected);
                miComm.Items.Add(miCommAll);
                miComm.Items.Add(miCommSel);
                btn.ContextMenu.Items.Add(miComm);

                var miMasters = new MenuItem { Header = "Remove Unused Masters" };
                miMasters.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveUnusedMasters);
                btn.ContextMenu.Items.Add(miMasters);

                var miSticky = new MenuItem { Header = "Delete All Sticky Notes" };
                miSticky.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.DeleteStickyNotes);
                btn.ContextMenu.Items.Add(miSticky);

                var miAnon = new MenuItem { Header = "Anonymize Selected Slides (Lorem Ipsum)" };
                miAnon.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.AnonymizeSelected);
                btn.ContextMenu.Items.Add(miAnon);

                var miFoot = new MenuItem { Header = "Remove Unedited Footers" };
                var miFootAll = new MenuItem { Header = "From All Slides" };
                miFootAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveUneditedFootersAll);
                var miFootSel = new MenuItem { Header = "From Selected Slides" };
                miFootSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveUneditedFootersSelected);
                miFoot.Items.Add(miFootAll);
                miFoot.Items.Add(miFootSel);
                btn.ContextMenu.Items.Add(miFoot);

                var miDoubleSpace = new MenuItem { Header = "Remove Double Spaces" };
                var miDoubleSpaceAll = new MenuItem { Header = "From All Slides" };
                miDoubleSpaceAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveDoubleSpacesAll);
                var miDoubleSpaceSel = new MenuItem { Header = "From Selected Slides" };
                miDoubleSpaceSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.RemoveDoubleSpacesSelected);
                miDoubleSpace.Items.Add(miDoubleSpaceAll);
                miDoubleSpace.Items.Add(miDoubleSpaceSel);
                btn.ContextMenu.Items.Add(miDoubleSpace);

                btn.ContextMenu.Items.Add(new Separator());
                var miConvert = new MenuItem { Header = "Convert Slides to Images" };
                var miConvertAll = new MenuItem { Header = "All Slides" };
                miConvertAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.ConvertSlidesToImagesAll);
                var miConvertSel = new MenuItem { Header = "Selected Slides" };
                miConvertSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.ConvertSlidesToImagesSelected);
                miConvert.Items.Add(miConvertAll);
                miConvert.Items.Add(miConvertSel);
                btn.ContextMenu.Items.Add(miConvert);

                btn.ContextMenu.Items.Add(new Separator());
                var miSuper = new MenuItem { Header = "Detect Superscript Numbers Without Footer" };
                miSuper.Click += (s, e) => CleanerFeature.Execute(GetManager(), CleanerFeature.CleanAction.DetectSuperscriptWithoutFooter);
                btn.ContextMenu.Items.Add(miSuper);
            }
            else if (btn.Name == "BtnArrangeCircle")
            {
                btn.ContextMenu.Items.Add(new Separator());
                string[] shapes = { "Circle", "Triangle", "Square", "Star", "Octagon", "Pentagon" };
                foreach (var sh in shapes) {
                    var mi = new MenuItem { Header = $"Arrange in {sh}" };
                    mi.Click += (s, e) => ArrangeInShapeFeature.Execute(GetManager(), sh.ToLower());
                    btn.ContextMenu.Items.Add(mi);
                }
                btn.ContextMenu.Items.Add(new Separator());
                var miPro = new MenuItem { Header = "Pro..." };
                miPro.Click += (s, e) => new UI.ArrangeProWindow(GetManager()).Show();
                btn.ContextMenu.Items.Add(miPro);
            }
            else if (btn.Name == "BtnShapeLock")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var miT = new MenuItem { Header = "Toggle Lock All Shapes (All Slides)" };
                miT.Click += (s, e) => ShapeLockingFeature.SetGlobalShapeLock(GetManager(), null);
                btn.ContextMenu.Items.Add(miT);
                var miL = new MenuItem { Header = "Lock All Shapes (All Slides)" };
                miL.Click += (s, e) => ShapeLockingFeature.SetGlobalShapeLock(GetManager(), true);
                btn.ContextMenu.Items.Add(miL);
                var miU = new MenuItem { Header = "Unlock All Shapes (All Slides)" };
                miU.Click += (s, e) => ShapeLockingFeature.SetGlobalShapeLock(GetManager(), false);
                btn.ContextMenu.Items.Add(miU);
            }
            else if (btn.Name == "BtnLockDimensions")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var miT = new MenuItem { Header = "Toggle Dimension Lock (All Slides)" };
                miT.Click += (s, e) => DimensionLockFeature.SetGlobalDimensionLock(GetManager(), null);
                btn.ContextMenu.Items.Add(miT);
                var miL = new MenuItem { Header = "Lock Dimensions (All Slides)" };
                miL.Click += (s, e) => DimensionLockFeature.SetGlobalDimensionLock(GetManager(), true);
                btn.ContextMenu.Items.Add(miL);
                var miU = new MenuItem { Header = "Unlock Dimensions (All Slides)" };
                miU.Click += (s, e) => DimensionLockFeature.SetGlobalDimensionLock(GetManager(), false);
                btn.ContextMenu.Items.Add(miU);
            }
            else if (btn.Name == "BtnIllustrativeSticker")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
                string defaultLabels = "Work in progress, Updated, For discussion, Backup, Confidential, Strictly confidential";
                string rawLabels = config?.IllustrativeStickerTexts ?? defaultLabels;
                var labels = rawLabels.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(l => l.Trim()).ToArray();
                if (labels.Length == 0) labels = defaultLabels.Split(',').Select(l => l.Trim()).ToArray();

                foreach (var l in labels) {
                    var mi = new MenuItem { Header = l };
                    mi.Click += (s, e) => IllustrativeStickerFeature.Execute(GetManager(), l);
                    btn.ContextMenu.Items.Add(mi);
                }
            }
            else if (btn.Name == "BtnHarveyBall")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var levels = new[] { 0f, 0.25f, 0.5f, 0.75f, 1f };
                foreach (var l in levels) {
                    var mi = new MenuItem { Header = $"Set {l * 100}%" };
                    mi.Click += (s, e) => HarveyBallFeature.Execute(GetManager(), l);
                    btn.ContextMenu.Items.Add(mi);
                }
            }
            else if (btn.Name == "BtnThermometer")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var levels = new[] { 0f, 0.25f, 0.5f, 0.75f, 1f };
                foreach (var l in levels) {
                    var mi = new MenuItem { Header = $"Set {l * 100}%" };
                    mi.Click += (s, e) => ThermometerFeature.Execute(GetManager(), l);
                    btn.ContextMenu.Items.Add(mi);
                }
            }
            else if (btn.Name == "BtnThemeColor")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Dock to Sidebar" };
                mi.Click += (s, e) => Globals.ThisAddIn.ToggleThemeColorPane();
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnSpecialChars")
            {
                btn.ContextMenu.Items.Clear();

                var symbolsItem = new MenuItem { Header = "Symbols" };
                foreach (var category in UI.SpecialCharactersWindow.SymbolData)
                {
                    var subItem = new MenuItem { Header = category.Key };
                    subItem.Click += (s, e) => OpenSpecialCharactersWindow(category.Key);
                    symbolsItem.Items.Add(subItem);
                }
                btn.ContextMenu.Items.Add(symbolsItem);

                var languagesItem = new MenuItem { Header = "Languages" };
                foreach (var lang in UI.SpecialCharactersWindow.LanguageData)
                {
                    var subItem = new MenuItem { Header = lang.Key };
                    subItem.Click += (s, e) => OpenSpecialCharactersWindow(lang.Key);
                    languagesItem.Items.Add(subItem);
                }
                btn.ContextMenu.Items.Add(languagesItem);
            }
            else if (btn.Name == "BtnSlideNotes")
            {
                btn.ContextMenu.Items.Add(new Separator());
                
                var mi1 = new MenuItem { Header = "Delete notes on this slide" };
                mi1.Click += (s, e) => oPenEfficiency.Features.Utilities.SlideNotesFeature.DeleteNotesOnCurrentSlide(GetManager());
                btn.ContextMenu.Items.Add(mi1);

                var mi2 = new MenuItem { Header = "Delete all notes in presentation" };
                mi2.Click += (s, e) => oPenEfficiency.Features.Utilities.SlideNotesFeature.DeleteAllNotes(GetManager());
                btn.ContextMenu.Items.Add(mi2);
            }
            else if (btn.Name == "BtnManageStickyNotes")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var mi1 = new MenuItem { Header = "Move notes off this slide" };
                mi1.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.MoveOff, StickyNoteManagerFeature.Scope.ThisSlide);
                btn.ContextMenu.Items.Add(mi1);

                var mi2 = new MenuItem { Header = "Move notes back on this slide" };
                mi2.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.MoveOn, StickyNoteManagerFeature.Scope.ThisSlide);
                btn.ContextMenu.Items.Add(mi2);

                var mi3 = new MenuItem { Header = "Delete notes on this slide" };
                mi3.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.Delete, StickyNoteManagerFeature.Scope.ThisSlide);
                btn.ContextMenu.Items.Add(mi3);

                var mi4 = new MenuItem { Header = "Convert comments on this slide to notes" };
                mi4.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.ConvertComments, StickyNoteManagerFeature.Scope.ThisSlide);
                btn.ContextMenu.Items.Add(mi4);

                btn.ContextMenu.Items.Add(new Separator());

                var mi5 = new MenuItem { Header = "Move notes off selected slides" };
                mi5.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.MoveOff, StickyNoteManagerFeature.Scope.SelectedSlides);
                btn.ContextMenu.Items.Add(mi5);

                var mi6 = new MenuItem { Header = "Move notes back on selected slides" };
                mi6.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.MoveOn, StickyNoteManagerFeature.Scope.SelectedSlides);
                btn.ContextMenu.Items.Add(mi6);

                var mi7 = new MenuItem { Header = "Delete notes on selected slides" };
                mi7.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.Delete, StickyNoteManagerFeature.Scope.SelectedSlides);
                btn.ContextMenu.Items.Add(mi7);

                var mi8 = new MenuItem { Header = "Convert comments on selected slides to notes" };
                mi8.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.ConvertComments, StickyNoteManagerFeature.Scope.SelectedSlides);
                btn.ContextMenu.Items.Add(mi8);

                btn.ContextMenu.Items.Add(new Separator());

                var mi9 = new MenuItem { Header = "Convert selected shapes to StickyNotes (Add prefix)" };
                mi9.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.PrefixName, StickyNoteManagerFeature.Scope.ThisSlide);
                btn.ContextMenu.Items.Add(mi9);

                btn.ContextMenu.Items.Add(new Separator());
                var miT = new MenuItem { Header = "Convert Sticky Notes to Tags (Selected Slides)" };
                miT.Click += (s, e) => {
                    // Similar logic as in TagInspectorWindow
                    try {
                        var app = GetManager().GetApplication();
                        int totalCount = 0;
                        PowerPoint.SlideRange slideRange = null;
                        try { slideRange = app.ActiveWindow.Selection.SlideRange; } catch { }
                        
                        var slides = new List<PowerPoint.Slide>();
                        if (slideRange != null && slideRange.Count > 0)
                        {
                            for (int x = 1; x <= slideRange.Count; x++) slides.Add(slideRange[x]);
                        }
                        else if (app.ActiveWindow.View.Slide is PowerPoint.Slide slide) slides.Add(slide);

                        foreach (var slideItem in slides)
                        {
                            int count = 0;
                            for (int i = 1; i <= slideItem.Shapes.Count; i++)
                            {
                                var shape = slideItem.Shapes[i];
                                if (shape.Name.StartsWith("StickyNote", StringComparison.OrdinalIgnoreCase))
                                {
                                    string text = "";
                                    try { text = shape.TextFrame.TextRange.Text; } catch { }
                                    if (!string.IsNullOrEmpty(text))
                                    {
                                        count++;
                                        slideItem.Tags.Add($"OE_STICKY_NOTE_{count}", text);
                                    }
                                }
                            }
                            totalCount += count;
                        }
                        MessageBox.Show($"Converted {totalCount} sticky note(s) to tags.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    } catch (Exception ex) {
                        MessageBox.Show("Error converting sticky notes: " + ex.Message);
                    }
                };
                btn.ContextMenu.Items.Add(miT);

                var miM = new MenuItem { Header = "Manage Tags..." };
                miM.Click += (s, e) => oPenEfficiency.Features.Utilities.TagInspectorFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(miM);
            }
            else if (btn.Name == "BtnTableSort")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var mi1 = new MenuItem { Header = "Sort from A to Z" };
                mi1.Click += (s, e) => TableSortFeature.Execute(GetManager(), TableSortFeature.SortType.AZ_Selection);
                btn.ContextMenu.Items.Add(mi1);

                var mi2 = new MenuItem { Header = "Sort from Z to A" };
                mi2.Click += (s, e) => TableSortFeature.Execute(GetManager(), TableSortFeature.SortType.ZA_Selection);
                btn.ContextMenu.Items.Add(mi2);

                var mi3 = new MenuItem { Header = "Sort from A to Z (whole row)" };
                mi3.Click += (s, e) => TableSortFeature.Execute(GetManager(), TableSortFeature.SortType.AZ_WholeRow);
                btn.ContextMenu.Items.Add(mi3);

                var mi4 = new MenuItem { Header = "Sort from Z to A (whole row)" };
                mi4.Click += (s, e) => TableSortFeature.Execute(GetManager(), TableSortFeature.SortType.ZA_WholeRow);
                btn.ContextMenu.Items.Add(mi4);
            }
            else if (btn.Name == "BtnUpdateExcelCharts")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var miUpdate = new MenuItem { Header = "Update Excel Charts in This Slide" };
                miUpdate.Click += (s, e) => UpdateExcelChartsFeature.Execute(GetManager(), updateAllSlides: false);
                btn.ContextMenu.Items.Add(miUpdate);

                btn.ContextMenu.Items.Add(new Separator());
                var miAll = new MenuItem { Header = "Update All Charts in Presentation" };
                miAll.Click += (s, e) => UpdateExcelChartsFeature.Execute(GetManager(), updateAllSlides: true);
                btn.ContextMenu.Items.Add(miAll);
            }
            else if (btn.Name == "BtnFlightMode")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var miToggle = new MenuItem { Header = "Toggle Flight Mode" };
                miToggle.Click += (s, e) => FlightModeFeature.Toggle(GetManager());
                btn.ContextMenu.Items.Add(miToggle);

                btn.ContextMenu.Items.Add(new Separator());

                var miCover = new MenuItem { Header = "Cover Logos Only" };
                miCover.Click += (s, e) => FlightModeFeature.EnableCoverLogos(GetManager());
                btn.ContextMenu.Items.Add(miCover);

                var miLight = new MenuItem { Header = "Light Overlay (95% Visible)" };
                miLight.Click += (s, e) => FlightModeFeature.EnableLightOverlay(GetManager());
                btn.ContextMenu.Items.Add(miLight);

                var miHeavy = new MenuItem { Header = "Heavy Overlay (85% Visible)" };
                miHeavy.Click += (s, e) => FlightModeFeature.EnableHeavyOverlay(GetManager());
                btn.ContextMenu.Items.Add(miHeavy);

                var miHide = new MenuItem { Header = "Hide Master Background" };
                miHide.Click += (s, e) => FlightModeFeature.EnableHideBackground(GetManager());
                btn.ContextMenu.Items.Add(miHide);

                btn.ContextMenu.Items.Add(new Separator());

                var miDisable = new MenuItem { Header = "Disable Flight Mode" };
                miDisable.Click += (s, e) => FlightModeFeature.Disable(GetManager());
                btn.ContextMenu.Items.Add(miDisable);
            }
            else if (btn.Name == "BtnInsertColumnLeft")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Insert Column Left & Preserve Width" };
                mi.Click += (s, e) => InsertTableColumnLeftFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnInsertColumnRight")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Insert Column Right & Preserve Width" };
                mi.Click += (s, e) => InsertTableColumnRightFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnInsertRowTop")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Insert Row Top & Preserve Height" };
                mi.Click += (s, e) => InsertTableRowTopFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnInsertRowBottom")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Insert Row Bottom & Preserve Height" };
                mi.Click += (s, e) => InsertTableRowBottomFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnRemoveHorizontalGaps")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Remove Horizontal Gaps & Center" };
                mi.Click += (s, e) => RemoveHorizontalGapsFeature.ExecuteCentered(GetManager());
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnRemoveVerticalGaps")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Remove Vertical Gaps & Center" };
                mi.Click += (s, e) => RemoveVerticalGapsFeature.ExecuteCentered(GetManager());
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnTableSplit")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Split Table by Row" };
                mi.Click += (s, e) => TableSplitFeature.Execute(GetManager(), false);
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnTableSum")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = "Sum Column in Current Selected Cells" };
                mi1.Click += (s, e) => TableSumFeature.Execute(GetManager(), false, true);
                btn.ContextMenu.Items.Add(mi1);

                var mi2 = new MenuItem { Header = "Sum Column and Row in Current Selected Cells" };
                mi2.Click += (s, e) => TableSumFeature.Execute(GetManager(), true, true);
                btn.ContextMenu.Items.Add(mi2);
            }
            else if (btn.Name == "BtnAdjustSpacing")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var miIncH = new MenuItem { Header = "Increase Horizontal Spacing" };
                miIncH.Click += (s, e) => AdjustHorizontalSpacingFeature.Execute(GetManager(), GetSpacingIncrement());
                btn.ContextMenu.Items.Add(miIncH);

                var miDecH = new MenuItem { Header = "Decrease Horizontal Spacing" };
                miDecH.Click += (s, e) => AdjustHorizontalSpacingFeature.Execute(GetManager(), -GetSpacingIncrement());
                btn.ContextMenu.Items.Add(miDecH);

                var miIncV = new MenuItem { Header = "Increase Vertical Spacing" };
                miIncV.Click += (s, e) => AdjustVerticalSpacingFeature.Execute(GetManager(), GetSpacingIncrement());
                btn.ContextMenu.Items.Add(miIncV);

                var miDecV = new MenuItem { Header = "Decrease Vertical Spacing" };
                miDecV.Click += (s, e) => AdjustVerticalSpacingFeature.Execute(GetManager(), -GetSpacingIncrement());
                btn.ContextMenu.Items.Add(miDecV);
            }
            else if (btn.Name == "BtnSnapShape")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = "Connect Selected Shapes" };
                mi1.Click += (s, e) => {
                    ToggleFloatingWindow("BtnSnapShape", () => new UI.ObjectConnectorWindow(GetManager()), null);
                };
                btn.ContextMenu.Items.Add(mi1);
            }
            else if (btn.Name == "BtnPickColor")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var miFill = new MenuItem { Header = "Pick Fill Color (Default)" };
                miFill.Click += (s, e) => Features.PickColorFeature.Execute(GetManager(), Features.PickColorFeature.ApplyTarget.Fill);
                btn.ContextMenu.Items.Add(miFill);

                var miLine = new MenuItem { Header = "Pick Line Color" };
                miLine.Click += (s, e) => Features.PickColorFeature.Execute(GetManager(), Features.PickColorFeature.ApplyTarget.Line);
                btn.ContextMenu.Items.Add(miLine);

                var miText = new MenuItem { Header = "Pick Text Color" };
                miText.Click += (s, e) => Features.PickColorFeature.Execute(GetManager(), Features.PickColorFeature.ApplyTarget.Text);
                btn.ContextMenu.Items.Add(miText);
            }
            else if (btn.Name == "BtnRepeatShape")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var miRight = new MenuItem { Header = "Repeat to the Right" };
                miRight.Click += (s, e) => RepeatShapeFeature.Execute(GetManager(), RepeatShapeFeature.RepeatDirection.Right);
                btn.ContextMenu.Items.Add(miRight);

                var miLeft = new MenuItem { Header = "Repeat to the Left" };
                miLeft.Click += (s, e) => RepeatShapeFeature.Execute(GetManager(), RepeatShapeFeature.RepeatDirection.Left);
                btn.ContextMenu.Items.Add(miLeft);

                var miUp = new MenuItem { Header = "Repeat to the Top" };
                miUp.Click += (s, e) => RepeatShapeFeature.Execute(GetManager(), RepeatShapeFeature.RepeatDirection.Up);
                btn.ContextMenu.Items.Add(miUp);

                var miDown = new MenuItem { Header = "Repeat to the Bottom" };
                miDown.Click += (s, e) => RepeatShapeFeature.Execute(GetManager(), RepeatShapeFeature.RepeatDirection.Down);
                btn.ContextMenu.Items.Add(miDown);

                btn.ContextMenu.Items.Add(new Separator());
                var miGrid = new MenuItem { Header = "Arrange in Grid..." };
                miGrid.Click += (s, e) => {
                    // Open grid arrangement dialog if available, or use default
                    RepeatShapeFeature.Execute(GetManager(), RepeatShapeFeature.RepeatDirection.Right);
                };
                btn.ContextMenu.Items.Add(miGrid);
            }
            else if (btn.Name == "BtnTableBranding")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var miSave = new MenuItem { Header = "Save Table Style" };
                miSave.Click += (s, e) => {
                    var dialog = new TableBrandingDialog();
                    dialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                    if (dialog.ShowDialog() == true && dialog.Action == "Save" && !string.IsNullOrEmpty(dialog.StyleToSave))
                    {
                        TableBrandingFeature.SaveTableStyle(GetManager(), dialog.StyleToSave);
                    }
                };
                btn.ContextMenu.Items.Add(miSave);

                var miLoad = new MenuItem { Header = "Load and Apply Table Style" };
                miLoad.Click += (s, e) => {
                    var dialog = new TableBrandingDialog();
                    dialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                    if (dialog.ShowDialog() == true && dialog.Action == "Apply" && dialog.SelectedStyle != null)
                    {
                        TableBrandingFeature.ApplyTableStyle(GetManager(), dialog.SelectedStyle);
                    }
                };
                btn.ContextMenu.Items.Add(miLoad);

                btn.ContextMenu.Items.Add(new Separator());
                var miManage = new MenuItem { Header = "Manage Table Styles..." };
                miManage.Click += (s, e) => {
                    var dialog = new TableBrandingDialog();
                    dialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                    dialog.ShowDialog();
                };
                btn.ContextMenu.Items.Add(miManage);
            }
            else if (btn.Name == "BtnTableDimensions")
            {
                var miCopyW = new MenuItem { Header = "Copy Column Width" };
                miCopyW.Click += (s, e) => TableColumnWidthFeature.Execute(GetManager(), false);
                btn.ContextMenu.Items.Add(miCopyW);

                var miCopyH = new MenuItem { Header = "Copy Row Height" };
                miCopyH.Click += (s, e) => TableRowHeightFeature.Execute(GetManager(), false);
                btn.ContextMenu.Items.Add(miCopyH);

                var miCopyB = new MenuItem { Header = "Copy Both (Width & Height)" };
                miCopyB.Click += (s, e) => TableDimensionsFeature.Execute(GetManager(), false);
                btn.ContextMenu.Items.Add(miCopyB);

                btn.ContextMenu.Items.Add(new Separator());

                var miPasteW = new MenuItem { Header = "Paste Column Width" };
                miPasteW.Click += (s, e) => TableColumnWidthFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(miPasteW);

                var miPasteH = new MenuItem { Header = "Paste Row Height" };
                miPasteH.Click += (s, e) => TableRowHeightFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(miPasteH);

                var miPasteB = new MenuItem { Header = "Paste Both (Width & Height)" };
                miPasteB.Click += (s, e) => TableDimensionsFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(miPasteB);
            }
            else if (btn.Name == "BtnSelectSameType")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var miFill = new MenuItem { Header = "Select shapes with same fill color" };
                miFill.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.FillColor);
                btn.ContextMenu.Items.Add(miFill);

                var miLine = new MenuItem { Header = "Select shapes with same line color" };
                miLine.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.LineColor);
                btn.ContextMenu.Items.Add(miLine);

                var miBoth = new MenuItem { Header = "Select shapes with same fill and line color" };
                miBoth.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.FillAndLineColor);
                btn.ContextMenu.Items.Add(miBoth);

                var miSize = new MenuItem { Header = "Select shapes with same size" };
                miSize.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.Size);
                btn.ContextMenu.Items.Add(miSize);

                var miWidth = new MenuItem { Header = "Select shapes with same width" };
                miWidth.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.Width);
                btn.ContextMenu.Items.Add(miWidth);

                var miHeight = new MenuItem { Header = "Select shapes with same height" };
                miHeight.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.Height);
                btn.ContextMenu.Items.Add(miHeight);

                var miType = new MenuItem { Header = "Select shapes with same type" };
                miType.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.Type);
                btn.ContextMenu.Items.Add(miType);

                var miFormatting = new MenuItem { Header = "Select shapes with same type and formatting" };
                miFormatting.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.TypeAndFormatting);
                btn.ContextMenu.Items.Add(miFormatting);
            }
            else if (btn.Name == "BtnBringToFront" || btn.Name == "BtnSendToBack" || btn.Name == "BtnBringForward" || btn.Name == "BtnSendBackward")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var miRanked = new MenuItem { Header = $"{UI.FeatureLibrary.GetFeatureInfo(btn.Name).Tooltip} (Ranked)" };
                miRanked.Click += (s, e) =>
                {
                    var mgr = GetManager();
                    if (btn.Name == "BtnBringToFront") BringToFrontFeature.ExecuteRanked(mgr);
                    else if (btn.Name == "BtnSendToBack") SendToBackFeature.ExecuteRanked(mgr);
                    else if (btn.Name == "BtnBringForward") BringForwardFeature.ExecuteRanked(mgr);
                    else if (btn.Name == "BtnSendBackward") SendBackwardFeature.ExecuteRanked(mgr);
                };
                btn.ContextMenu.Items.Add(miRanked);
            }
            else if (btn.Name == "BtnRotateShape")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = "Rotate 90° Left" };
                mi1.Click += (s, e) => oPenEfficiency.Features.Visuals.RotateShapeFeature.ExecuteSpecific(GetManager(), -90f);
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = "Rotate 45° Right" };
                mi2.Click += (s, e) => oPenEfficiency.Features.Visuals.RotateShapeFeature.ExecuteSpecific(GetManager(), 45f);
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = "Rotate 45° Left" };
                mi3.Click += (s, e) => oPenEfficiency.Features.Visuals.RotateShapeFeature.ExecuteSpecific(GetManager(), -45f);
                btn.ContextMenu.Items.Add(mi3);
            }
            else if (btn.Name == "BtnChangeCase")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = "Sentence case." };
                mi1.Click += (s, e) => oPenEfficiency.Features.Text.ChangeCaseFeature.ExecuteSpecific(GetManager(), PowerPoint.PpChangeCase.ppCaseSentence);
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = "lowercase" };
                mi2.Click += (s, e) => oPenEfficiency.Features.Text.ChangeCaseFeature.ExecuteSpecific(GetManager(), PowerPoint.PpChangeCase.ppCaseLower);
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = "UPPERCASE" };
                mi3.Click += (s, e) => oPenEfficiency.Features.Text.ChangeCaseFeature.ExecuteSpecific(GetManager(), PowerPoint.PpChangeCase.ppCaseUpper);
                btn.ContextMenu.Items.Add(mi3);
                var mi4 = new MenuItem { Header = "Capitalize Each Word" };
                mi4.Click += (s, e) => oPenEfficiency.Features.Text.ChangeCaseFeature.ExecuteSpecific(GetManager(), PowerPoint.PpChangeCase.ppCaseTitle);
                btn.ContextMenu.Items.Add(mi4);
                var mi5 = new MenuItem { Header = "tOGGLE cASE" };
                mi5.Click += (s, e) => oPenEfficiency.Features.Text.ChangeCaseFeature.ExecuteSpecific(GetManager(), PowerPoint.PpChangeCase.ppCaseToggle);
                btn.ContextMenu.Items.Add(mi5);
            }
            else if (btn.Name == "BtnFormatBold")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Remove Bold Formatting" };
                mi.Click += (s, e) => oPenEfficiency.Features.Text.FormatBoldFeature.Remove(GetManager());
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnFormatItalic")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Remove Italic Formatting" };
                mi.Click += (s, e) => oPenEfficiency.Features.Text.FormatItalicFeature.Remove(GetManager());
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnFormatUnderline")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Remove Underline Formatting" };
                mi.Click += (s, e) => oPenEfficiency.Features.Text.FormatUnderlineFeature.Remove(GetManager());
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnFormatStrikethrough")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Remove Strikethrough Formatting" };
                mi.Click += (s, e) => oPenEfficiency.Features.Text.FormatStrikethroughFeature.Remove(GetManager());
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnFormatSubscript")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Remove Subscript Formatting" };
                mi.Click += (s, e) => oPenEfficiency.Features.Text.FormatSubscriptFeature.Remove(GetManager());
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnFormatSuperscript")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Remove Superscript Formatting" };
                mi.Click += (s, e) => oPenEfficiency.Features.Text.FormatSuperscriptFeature.Remove(GetManager());
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnDistributeHorizontal")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = "Align to Slide" };
                mi1.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeHorizontalFeature.Execute(GetManager(), relativeToSlide: true);
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = "Align Selected Objects" };
                mi2.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeHorizontalFeature.Execute(GetManager(), relativeToSlide: false);
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = "Align Middles & Distribute" };
                mi3.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeHorizontalFeature.Execute(GetManager(), relativeToSlide: false, alignMiddles: true);
                btn.ContextMenu.Items.Add(mi3);
                var mi4 = new MenuItem { Header = "Align Middles & Distribute (to Slide)" };
                mi4.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeHorizontalFeature.Execute(GetManager(), relativeToSlide: true, alignMiddles: true);
                btn.ContextMenu.Items.Add(mi4);
            }
            else if (btn.Name == "BtnDistributeVertical")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = "Align to Slide" };
                mi1.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeVerticalFeature.Execute(GetManager(), relativeToSlide: true);
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = "Align Selected Objects" };
                mi2.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeVerticalFeature.Execute(GetManager(), relativeToSlide: false);
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = "Align Centers & Distribute" };
                mi3.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeVerticalFeature.Execute(GetManager(), relativeToSlide: false, alignCenters: true);
                btn.ContextMenu.Items.Add(mi3);
                var mi4 = new MenuItem { Header = "Align Centers & Distribute (to Slide)" };
                mi4.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeVerticalFeature.Execute(GetManager(), relativeToSlide: true, alignCenters: true);
                btn.ContextMenu.Items.Add(mi4);
            }
            else if (btn.Name == "BtnDiagonalAlign")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = "Align 45° (Anchor: Top/Left-most shape)" };
                mi1.Click += (s, e) => oPenEfficiency.Features.Alignment.DiagonalAlignFeature.ExecuteSpecific(GetManager(), 45f, "First");
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = "Align 135° (Anchor: Top/Left-most shape)" };
                mi2.Click += (s, e) => oPenEfficiency.Features.Alignment.DiagonalAlignFeature.ExecuteSpecific(GetManager(), 135f, "First");
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = "Align -45° (Anchor: Top/Left-most shape)" };
                mi3.Click += (s, e) => oPenEfficiency.Features.Alignment.DiagonalAlignFeature.ExecuteSpecific(GetManager(), -45f, "First");
                btn.ContextMenu.Items.Add(mi3);
                var mi4 = new MenuItem { Header = "Align -135° (Anchor: Top/Left-most shape)" };
                mi4.Click += (s, e) => oPenEfficiency.Features.Alignment.DiagonalAlignFeature.ExecuteSpecific(GetManager(), -135f, "First");
                btn.ContextMenu.Items.Add(mi4);
            }
            else if (btn.Name == "BtnAlignText")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var miL = new MenuItem { Header = "Left" };
                miL.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), null, PowerPoint.PpParagraphAlignment.ppAlignLeft);
                btn.ContextMenu.Items.Add(miL);
                var miC = new MenuItem { Header = "Center" };
                miC.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), null, PowerPoint.PpParagraphAlignment.ppAlignCenter);
                btn.ContextMenu.Items.Add(miC);
                var miR = new MenuItem { Header = "Right" };
                miR.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), null, PowerPoint.PpParagraphAlignment.ppAlignRight);
                btn.ContextMenu.Items.Add(miR);
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = "Top" };
                mi1.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorTop);
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = "Middle" };
                mi2.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle);
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = "Bottom" };
                mi3.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorBottom);
                btn.ContextMenu.Items.Add(mi3);
                btn.ContextMenu.Items.Add(new Separator());
                var mi4 = new MenuItem { Header = "Top Centered" };
                mi4.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorTop, PowerPoint.PpParagraphAlignment.ppAlignCenter);
                btn.ContextMenu.Items.Add(mi4);
                var mi5 = new MenuItem { Header = "Middle Centered" };
                mi5.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle, PowerPoint.PpParagraphAlignment.ppAlignCenter);
                btn.ContextMenu.Items.Add(mi5);
                var mi6 = new MenuItem { Header = "Bottom Centered" };
                mi6.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorBottom, PowerPoint.PpParagraphAlignment.ppAlignCenter);
                btn.ContextMenu.Items.Add(mi6);
            }
            else if (btn.Name == "BtnTextDirection")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = "Horizontal" };
                mi1.Click += (s, e) => oPenEfficiency.Features.Text.TextDirectionFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal);
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = "Rotate all text 90°" };
                mi2.Click += (s, e) => oPenEfficiency.Features.Text.TextDirectionFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationDownward);
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = "Rotate all text 270°" };
                mi3.Click += (s, e) => oPenEfficiency.Features.Text.TextDirectionFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationUpward);
                btn.ContextMenu.Items.Add(mi3);
                var mi4 = new MenuItem { Header = "Stacked" };
                mi4.Click += (s, e) => oPenEfficiency.Features.Text.TextDirectionFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationVerticalFarEast);
                btn.ContextMenu.Items.Add(mi4);
            }
            else if (btn.Name == "BtnCharacterSpacing")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = "Very Tight" };
                mi1.Click += (s, e) => oPenEfficiency.Features.Text.CharacterSpacingFeature.ExecuteSpecific(GetManager(), -1.5f);
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = "Tight" };
                mi2.Click += (s, e) => oPenEfficiency.Features.Text.CharacterSpacingFeature.ExecuteSpecific(GetManager(), -0.75f);
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = "Normal" };
                mi3.Click += (s, e) => oPenEfficiency.Features.Text.CharacterSpacingFeature.ExecuteSpecific(GetManager(), 0f);
                btn.ContextMenu.Items.Add(mi3);
                var mi4 = new MenuItem { Header = "Loose" };
                mi4.Click += (s, e) => oPenEfficiency.Features.Text.CharacterSpacingFeature.ExecuteSpecific(GetManager(), 1.5f);
                btn.ContextMenu.Items.Add(mi4);
                var mi5 = new MenuItem { Header = "Very Loose" };
                mi5.Click += (s, e) => oPenEfficiency.Features.Text.CharacterSpacingFeature.ExecuteSpecific(GetManager(), 3.0f);
                btn.ContextMenu.Items.Add(mi5);
                var mi6 = new MenuItem { Header = "More Spacing..." };
                mi6.Click += (s, e) => oPenEfficiency.Features.Text.CharacterSpacingFeature.ExecuteMoreSpacing(GetManager());
                btn.ContextMenu.Items.Add(mi6);
            }
            else if (btn.Name == "BtnTranslate")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var miAPI = new MenuItem { Header = "Translate via Configured API" };
                miAPI.Click += (s, e) => {
                    var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
                    bool hasKey = false;
                    if (config != null)
                    {
                        if (config.SelectedTranslationTool == "DeepL" && !string.IsNullOrEmpty(config.DeepLApiKey)) hasKey = true;
                        else if (config.SelectedTranslationTool == "OpenAI" && !string.IsNullOrEmpty(config.OpenAIApiKey)) hasKey = true;
                        else if (config.SelectedTranslationTool == "Claude" && !string.IsNullOrEmpty(config.ClaudeApiKey)) hasKey = true;
                        else if (config.SelectedTranslationTool == "LibreTranslate" && !string.IsNullOrEmpty(config.LibreTranslateBaseUrl)) hasKey = true;
                    }
                    if (hasKey) oPenEfficiency.Features.Text.TranslationFeatures.TranslateSelectionSafe(GetManager(), config);
                    else System.Windows.MessageBox.Show("Please configure an API provider in the settings first.", "Translation", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                };
                btn.ContextMenu.Items.Add(miAPI);

                var miFree = new MenuItem { Header = "Translate via Free API (MyMemory)" };
                miFree.Click += (s, e) => {
                    var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
                    string lang = config?.TranslationTargetLanguage ?? "EN-US";
                    oPenEfficiency.Features.Text.TranslationFeatures.TranslateWithFreeApiSafe(GetManager(), lang);
                };
                btn.ContextMenu.Items.Add(miFree);
            }
            else if (btn.Name == "BtnTransparency")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var levels = new[] { 0f, 0.25f, 0.5f, 0.75f, 1f };
                foreach (var l in levels) {
                    var mi = new MenuItem { Header = $"Set {l * 100}%" };
                    mi.Click += (s, e) => oPenEfficiency.Features.Visuals.TransparencyFeature.SetTransparency(GetManager(), l);
                    btn.ContextMenu.Items.Add(mi);
                }
            }
        }

        private void OpenSpecialCharactersWindow(string selection)
        {
            try
            {
                var sidebarWindow = Window.GetWindow(this);

                if (_specialCharsWindow == null)
                {
                    _specialCharsWindow = new UI.SpecialCharactersWindow(GetManager());
                    if (sidebarWindow != null) _specialCharsWindow.Owner = sidebarWindow;
                }

                // Ensure it's not minimized and visible
                if (_specialCharsWindow.WindowState == WindowState.Minimized)
                    _specialCharsWindow.WindowState = WindowState.Normal;

                _specialCharsWindow.SelectLanguage(selection);

                // Positioning near sidebar
                if (sidebarWindow != null)
                {
                    double targetLeft = sidebarWindow.Left - _specialCharsWindow.Width - 10;
                    // If it would be off-screen to the left, try the right side
                    if (targetLeft < 0) targetLeft = sidebarWindow.Left + sidebarWindow.Width + 10;
                    
                    _specialCharsWindow.Left = targetLeft;
                    _specialCharsWindow.Top = sidebarWindow.Top;
                }
                else
                {
                    _specialCharsWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                }

                _specialCharsWindow.Show();
                _specialCharsWindow.Activate();
                _specialCharsWindow.Focus();
                _specialCharsWindow.Topmost = true;
            }
            catch
            {
                _specialCharsWindow?.Show();
            }
        }
    }
}
