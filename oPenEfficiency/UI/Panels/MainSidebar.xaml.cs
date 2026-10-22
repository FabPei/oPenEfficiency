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
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Services.Commands;
using oPenEfficiency.UI.Dialogs;
using oPenEfficiency.Utils;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency
{
    public partial class MainSidebar : UserControl, INotifyPropertyChanged
    {
        private static MainSidebar _activeInstance;
        private PowerPointManager _powerPointManager;
        private UI.Panels.ISidebarHost _host;
        private UI.ShapeSizePanel _sizePanel;
        private UI.FontPanel _fontPanel;
        private UI.SearchBarControl _searchBar;
        private UI.Panels.SidebarNavigationService _navigation;
        private UI.Panels.SidebarContextMenuProvider _contextMenus;
        private UI.Panels.SidebarNavigationService Navigation =>
            _navigation ?? (_navigation = new UI.Panels.SidebarNavigationService(
                this,
                GetManager,
                () => RebuildUI()));
        private UI.Panels.SidebarContextMenuProvider ContextMenus =>
            _contextMenus ?? (_contextMenus = new UI.Panels.SidebarContextMenuProvider(
                TryAttachRemainingFeatureContextMenu,
                AttachCatalogContextMenu));
        private SidebarSelectionSnapshot _selectionSnapshot;
        private bool _themeSubscribed;
        private string _currentIconColorMode = IconColorModes.PowerPoint;
        private string _activeSearchTerm = string.Empty;

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
            _floatingWindows.TryGetValue(key, out var existing);
            if (existing != null && !existing.IsLoaded)
            {
                _floatingWindows.Remove(key);
                existing = null;
            }

            if (toggleBtn != null && toggleBtn.IsChecked == false)
            {
                existing?.Close();
                _floatingWindows.Remove(key);
                return;
            }

            if (existing != null)
            {
                SetPowerPointOwner(existing, GetManager());
                if (!existing.IsVisible) existing.Show();
                if (existing.WindowState == WindowState.Minimized)
                    existing.WindowState = WindowState.Normal;
                existing.Activate();
                existing.Focus();
                if (toggleBtn != null) toggleBtn.IsChecked = true;
                return;
            }

            var window = createAction();
            if (window == null) throw new InvalidOperationException("Window factory returned null.");
            SetPowerPointOwner(window, GetManager());
            window.Closed += (s, e) =>
            {
                if (toggleBtn != null) toggleBtn.IsChecked = false;
                if (_floatingWindows.TryGetValue(key, out var tracked) && ReferenceEquals(tracked, window))
                    _floatingWindows.Remove(key);
            };
            try
            {
                window.Show();
                if (!window.IsLoaded || !window.IsVisible)
                    throw new InvalidOperationException(
                        $"Window for '{key}' closed before it became visible.");
                _floatingWindows[key] = window;
                if (toggleBtn != null) toggleBtn.IsChecked = true;
            }
            catch
            {
                try { window.Close(); }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(closeEx, "MainSidebar.ToggleFloatingWindow.FailureCleanup");
                }
                _floatingWindows.Remove(key);
                if (toggleBtn != null) toggleBtn.IsChecked = false;
                throw;
            }
        }

        internal bool ExerciseFloatingWindowLifecycleForAutomation()
        {
            const string key = "OfficeIntegrationFloatingProbe";
            int initialWindowCount = _floatingWindows.Count;
            Window probe = null;
            try
            {
                ToggleFloatingWindow(
                    key,
                    () => probe = new Window
                    {
                        Width = 1,
                        Height = 1,
                        Left = -32000,
                        Top = -32000,
                        Opacity = 0,
                        ShowActivated = false,
                        ShowInTaskbar = false,
                        WindowStyle = WindowStyle.None
                    },
                    null);
                probe.Dispatcher.Invoke(() => { }, DispatcherPriority.Background);
                bool opened = probe.IsLoaded && _floatingWindows.Count == initialWindowCount + 1;
                probe.Close();
                probe.Dispatcher.Invoke(() => { }, DispatcherPriority.Background);
                return opened && !probe.IsLoaded && _floatingWindows.Count == initialWindowCount;
            }
            finally
            {
                if (probe?.IsLoaded == true) probe.Close();
                _floatingWindows.Remove(key);
            }
        }


        private void ShowOrCenterGanttWindow(
            PowerPointManager manager,
            ToggleButton toggleBtn,
            UI.Panels.GanttStartAction startAction = UI.Panels.GanttStartAction.Open)
        {
            const string key = "BtnGanttChart";

            _floatingWindows.TryGetValue(key, out var existingWindow);
            if (existingWindow != null && !existingWindow.IsLoaded)
            {
                _floatingWindows.Remove(key);
                existingWindow = null;
            }

            if (toggleBtn != null && toggleBtn.IsChecked == false)
            {
                existingWindow?.Close();
                _floatingWindows.Remove(key);
                return;
            }

            if (existingWindow != null)
            {
                (existingWindow as UI.Dialogs.GanttChartWindow)?.ApplyStartAction(startAction);
                if (toggleBtn != null) toggleBtn.IsChecked = true;
                SetPowerPointOwner(existingWindow, manager);
                if (!existingWindow.IsVisible) existingWindow.Show();
                if (existingWindow.WindowState == WindowState.Minimized)
                    existingWindow.WindowState = WindowState.Normal;
                CenterWindowOverSlide(existingWindow, manager);
                existingWindow.Activate();
                existingWindow.Focus();
                return;
            }

            var window = new UI.Dialogs.GanttChartWindow(manager, startAction);
            SetPowerPointOwner(window, manager);
            window.Closed += (s, e) =>
            {
                if (toggleBtn != null) toggleBtn.IsChecked = false;
                if (_floatingWindows.TryGetValue(key, out var tracked) && ReferenceEquals(tracked, window))
                    _floatingWindows.Remove(key);
            };
            try
            {
                window.Show();
                if (!window.IsLoaded || !window.IsVisible)
                    throw new InvalidOperationException(
                        "Gantt window closed before it became visible.");
                _floatingWindows[key] = window;
                if (toggleBtn != null) toggleBtn.IsChecked = true;
                CenterWindowOverSlide(window, manager);
            }
            catch
            {
                try { window.Close(); }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(closeEx, "MainSidebar.ShowOrCenterGanttWindow.FailureCleanup");
                }
                _floatingWindows.Remove(key);
                if (toggleBtn != null) toggleBtn.IsChecked = false;
                throw;
            }
        }

        private static void SetPowerPointOwner(Window window, PowerPointManager manager)
        {
            var app = manager?.GetApplication();
            if (window != null && app != null)
                new System.Windows.Interop.WindowInteropHelper(window).Owner = (IntPtr)app.HWND;
        }

        private void CenterWindowOverSlide(Window window, PowerPointManager manager)
        {
            try
            {
                var app = manager.GetApplication();
                if (app?.ActiveWindow == null) return;

                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;

                int centerXPx = app.ActiveWindow.PointsToScreenPixelsX(slideWidth / 2);
                int centerYPx = app.ActiveWindow.PointsToScreenPixelsY(slideHeight / 2);

                float dpiScaleX, dpiScaleY;
                using (var g = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
                {
                    dpiScaleX = g.DpiX / 96f;
                    dpiScaleY = g.DpiY / 96f;
                }

                double wpfW = window.ActualWidth > 0 ? window.ActualWidth : window.Width;
                double wpfH = window.ActualHeight > 0 ? window.ActualHeight : window.Height;
                double wpfX = (centerXPx / dpiScaleX) - (wpfW / 2);
                double wpfY = (centerYPx / dpiScaleY) - (wpfH / 2);

                wpfX = Math.Max(0, Math.Min(wpfX, SystemParameters.PrimaryScreenWidth - wpfW));
                wpfY = Math.Max(0, Math.Min(wpfY, SystemParameters.PrimaryScreenHeight - wpfH));

                window.Left = wpfX;
                window.Top = wpfY;
            }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "MainSidebar.PositionFloatingWindow"); }
        }

        public MainSidebar()
        {
            InitializeComponent();
            RegisterShortcuts();

            _resizeTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
            _resizeTimer.Tick += ResizeTimer_Tick;

            // Initial Rebuild from Config
            this.Loaded += MainSidebar_Loaded;
            this.Unloaded += MainSidebar_Unloaded;
            this.MouseEnter += (s, e) => RefreshButtonStates();

        }

        private void MainSidebar_Loaded(object sender, RoutedEventArgs e)
        {
            _activeInstance = this;
            if (!_themeSubscribed)
            {
                oPenEfficiency.Services.ThemeManager.ThemeChanged += ThemeManager_ThemeChanged;
                _themeSubscribed = true;
            }
            RebuildUI();
        }

        private void ThemeManager_ThemeChanged(object sender, string themeName)
        {
            RefreshThemeResources();
            RebuildUI(_activeSearchTerm);
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
                ExceptionLogger.Log(ex, "MainSidebar.RefreshThemeResources");
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
            _sizePanel?.OnSelectionChange(sel);
            _fontPanel?.OnSelectionChange(sel);
            UpdateButtonStates();
        }

        public void RefreshButtonStates()
        {
            UpdateButtonStates();
        }

        internal int SelectionShapeCountForAutomation => _selectionSnapshot?.ShapeCount ?? -1;

        /// <summary>Called by ThisAddIn to update size panel while editing text inside a shape.</summary>
        public void UpdateSizePanel(PowerPoint.Shape shape)
        {
            _sizePanel?.UpdateSize(shape);
        }

        // Current button background brush, set by RebuildUI and used by CreateFeatureButton
        private Brush _currentBtnBgBrush = Brushes.Transparent;

        private static List<SidebarFeature> GetSearchMatches(string searchTerm)
        {
            return SidebarSearchService.Search(UI.FeatureLibrary.Catalog, searchTerm)
                .Select(feature => new SidebarFeature
                {
                    Id = feature.Id,
                    Name = feature.Name
                })
                .ToList();
        }

        private void ExecuteFirstSearchMatch(string searchTerm)
        {
            var feature = GetSearchMatches(searchTerm).FirstOrDefault();
            if (feature == null) return;

            var target = FindElementByName(DynamicContentPanel, feature.Id);
            ExecuteFeature(feature.Id, target, CommandInvocationSource.Search);
        }

        private static FrameworkElement FindElementByName(DependencyObject parent, string name)
        {
            if (parent == null) return null;
            if (parent is FrameworkElement element && string.Equals(element.Name, name, StringComparison.Ordinal))
                return element;

            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var match = FindElementByName(VisualTreeHelper.GetChild(parent, i), name);
                if (match != null) return match;
            }

            return null;
        }

        public void RebuildUI(string searchTerm = "")
        {
            if (DynamicContentPanel == null) return;

            searchTerm = searchTerm ?? string.Empty;
            _activeSearchTerm = searchTerm;

            DynamicContentPanel.Children.Clear();

            // Update the current layout mode
            _isCurrentlyHorizontal = this.ActualWidth > (this.ActualHeight * 1.5);

            bool isSearchMode = !string.IsNullOrWhiteSpace(searchTerm);
            SidebarConfig config = JsonService.Load<SidebarConfig>("sidebar_config.json");
            bool hasSavedLayout = config != null;
            config = SidebarPresetFactory.ResolveInitialLayout(
                config,
                UI.FeatureLibrary.Catalog);
            if (hasSavedLayout)
                FilterObsoleteFeatures(config);

            bool appearanceModeChanged = SidebarBackgroundModes.Normalize(config);
            appearanceModeChanged |= SidebarSectionFontColorModes.Normalize(config);
            if (appearanceModeChanged && hasSavedLayout)
                JsonService.Save(config, "sidebar_config.json");

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

                var searchSection = new SidebarSection { Name = "Search Results", Color = "#10B981" };

                var matches = GetSearchMatches(searchTerm);

                foreach (var m in matches) searchSection.Features.Add(m);
                config.Sections = new List<SidebarSection> { searchSection };
            }

            // Apply either the live theme token or the user's explicit custom colour.
            try
            {
                if (!SidebarBackgroundModes.UsesTheme(config))
                {
                    this.Background = (SolidColorBrush)new BrushConverter().ConvertFrom(config.BackgroundColor);
                }
                else
                {
                    this.SetResourceReference(Control.BackgroundProperty, "SidebarBackgroundBrush");
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MainSidebar.ApplyBackground");
                this.SetResourceReference(Control.BackgroundProperty, "SidebarBackgroundBrush");
            }

            // Resolve section header font size
            double sectionFontSize = config.SectionFontSize > 0 ? config.SectionFontSize : 9;

            // Resolve icon size (button size)
            double iconSize = config.IconSize > 0 ? config.IconSize : 16;
            _currentIconColorMode = IconColorModes.Normalize(config.IconColorMode);

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
                        BorderThickness = new Thickness(0, 0, 0.5, 0), // Right border only
                        Margin = (Thickness)FindResource("Space4"),
                        Padding = (Thickness)FindResource("Space8"),
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalAlignment = VerticalAlignment.Top
                    };
                    sectionBorder.SetResourceReference(Border.BorderBrushProperty, "SidebarDividerBrush");

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
                        Style = (Style)FindResource("SidebarSectionHeader"),
                        FontSize = sectionFontSize,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        Margin = new Thickness(0, 0, 0, 4)
                    };
                    ApplySectionHeaderForeground(header, config);
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
                            var toggleBtn = CreateToggleButton(feature.Id, section.Color, iconSize);
                            btn = toggleBtn;
                        }
                        else
                        {
                            btn = CreateFeatureButton(feature.Id, section.Color, iconSize);
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

                // Keep Settings and Feature Explorer reachable for legacy customized
                // layouts that predate their persistent catalog section.
                bool needsSettings = !ContainsFeature(config, "BtnSettings");
                bool needsFeatureExplorer = !ContainsFeature(config, "BtnExploreFeatures");
                var settingsBorder = new Border
                {
                    BorderThickness = new Thickness(0, 0, 0.5, 0), // Right border only
                    Margin = (Thickness)FindResource("Space4"),
                    Padding = (Thickness)FindResource("Space8"),
                    HorizontalAlignment = HorizontalAlignment.Left,
                    VerticalAlignment = VerticalAlignment.Top
                };
                settingsBorder.SetResourceReference(Border.BorderBrushProperty, "SidebarDividerBrush");

                var settingsStack = new StackPanel
                {
                    Orientation = Orientation.Vertical,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Top
                };

                var settingsHeader = new TextBlock
                {
                    Text = "SETTINGS & HELP",
                    Style = (Style)FindResource("SidebarSectionHeader"),
                    HorizontalAlignment = HorizontalAlignment.Center
                };
                ApplySectionHeaderForeground(settingsHeader, config);
                settingsStack.Children.Add(settingsHeader);

                var settingsBtn = needsSettings
                    ? CreateFeatureButton("BtnSettings", null, iconSize)
                    : null;
                if (settingsBtn != null)
                {
                    settingsBtn.Width = iconSize + 10;
                    settingsBtn.Height = iconSize + 10;
                    settingsBtn.Margin = (Thickness)FindResource("Space2");
                    settingsBtn.HorizontalAlignment = HorizontalAlignment.Center;
                    settingsStack.Children.Add(settingsBtn);
                }

                var exploreBtn = needsFeatureExplorer
                    ? CreateFeatureButton("BtnExploreFeatures", null, iconSize)
                    : null;
                if (exploreBtn != null)
                {
                    exploreBtn.Width = iconSize + 10;
                    exploreBtn.Height = iconSize + 10;
                    exploreBtn.Margin = (Thickness)FindResource("Space2");
                    exploreBtn.HorizontalAlignment = HorizontalAlignment.Center;
                    settingsStack.Children.Add(exploreBtn);
                }

                // Set explicit height on settings border to fill available height
                double settingsHeight = Math.Max(0, this.ActualHeight - 16);
                if (settingsHeight > 0)
                {
                    settingsBorder.Height = settingsHeight;
                }

                settingsBorder.Child = settingsStack;
                if (needsSettings || needsFeatureExplorer)
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
                    Style      = (Style)FindResource("SidebarSectionHeader"),
                    FontSize   = sectionFontSize
                };
                ApplySectionHeaderForeground(header, config);
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
                            var toggleBtn = CreateToggleButton(feature.Id, section.Color, iconSize);
                            btn = toggleBtn;
                        }
                        else
                        {
                            btn = CreateFeatureButton(feature.Id, section.Color, iconSize);
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
                var separator = new Rectangle
                {
                    Height = 0.5,
                    Margin = (Thickness)FindResource("Space2")
                };
                separator.SetResourceReference(Shape.FillProperty, "SidebarDividerBrush");
                DynamicContentPanel.Children.Add(separator);
            }
            AddSettingsButton(
                includeSettings: !ContainsFeature(config, "BtnSettings"),
                includeFeatureExplorer: !ContainsFeature(config, "BtnExploreFeatures"));
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
                    {
                        _sizePanel = new UI.ShapeSizePanel(_powerPointManager);
                        AttachWidgetContextMenu(_sizePanel, "BtnSizePanel");
                    }
                    return _sizePanel;
                case "BtnFontPanel":
                    if (_fontPanel == null && _powerPointManager != null)
                    {
                        _fontPanel = new UI.FontPanel(_powerPointManager);
                        AttachWidgetContextMenu(_fontPanel, "BtnFontPanel");
                    }
                    return _fontPanel;
                case "BtnSearchBar":
                    if (_searchBar == null)
                    {
                        _searchBar = new UI.SearchBarControl();
                        _searchBar.SearchChanged += (s, term) => RebuildUI(term);
                        _searchBar.SearchSubmitted += (s, term) => ExecuteFirstSearchMatch(term);
                        AttachWidgetContextMenu(_searchBar, "BtnSearchBar");
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
            // Capture all shared COM-derived selection facts exactly once per refresh.
            _selectionSnapshot = CreateSelectionSnapshot(
                _powerPointManager?.GetCurrentSelectionMetadata());

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
            bool isEnabled = IsFeatureEnabled(info);
            if (isEnabled && IsTextTargetFeature(btn.Name))
                isEnabled = HasUsableTextTarget(btn.Name);
            if (isEnabled && AlignmentHelpers.IsStretchFeature(btn.Name))
                isEnabled = AlignmentHelpers.CanStretch(GetManager(), btn.Name);
            btn.IsEnabled = isEnabled;
        }

        private static bool IsTextTargetFeature(string id)
        {
            return id == "BtnSplitByParagraphs" ||
                id == "BtnApplyTextTool" ||
                id == "BtnParagraphDialog" ||
                id == "BtnSpecialChars";
        }

        private bool HasUsableTextTarget(string id)
        {
            return SidebarStateEvaluator.HasUsableTextTarget(id, _selectionSnapshot);
        }

        private void UpdateButtonState(ToggleButton btn)
        {
            var info = UI.FeatureLibrary.GetFeatureInfo(btn.Name);
            bool isEnabled = IsFeatureEnabled(info);
            btn.IsEnabled = isEnabled;

            if (btn.Name == "BtnSnapToObjects")
            {
                btn.IsChecked = SnapToObjectsFeature.IsEnabled;
                return;
            }
            if (btn.Name == "BtnSnapToGrid")
            {
                btn.IsChecked = SnapToGridFeature.IsEnabled;
                return;
            }

            if (btn.Name == "BtnShapeLock" || btn.Name == "BtnLockDimensions")
            {
                PowerPoint.Selection currentSelection = null;
                try
                {
                    currentSelection = _powerPointManager?.GetApplication()?.ActiveWindow?.Selection;
                    if (!isEnabled || currentSelection == null || currentSelection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        btn.IsChecked = false;
                        return;
                    }

                    int lockedCount = 0;
                    int totalCount = currentSelection.ShapeRange.Count;

                    for (int i = 1; i <= totalCount; i++)
                    {
                        bool isLocked = btn.Name == "BtnShapeLock"
                            ? ShapeLockingFeature.IsShapeLocked(currentSelection.ShapeRange[i])
                            : DimensionLockFeature.IsDimensionLocked(currentSelection.ShapeRange[i]);

                        if (isLocked) lockedCount++;
                    }

                    // For toggle button we can set IsThreeState to true to support null
                    btn.IsThreeState = true;

                    btn.IsChecked = SidebarStateEvaluator.ResolveAggregateToggleState(
                        isEnabled,
                        totalCount,
                        lockedCount);
                }
                catch
                {
                    btn.IsChecked = false;
                }
            }
            else if (btn.Name == "BtnAssetManager")
            {
                btn.IsChecked = _host?.IsAssetManagerVisible == true;
            }
            else if (btn.Name == "BtnFlightMode")
            {
                var manager = GetManager();
                btn.IsChecked = manager != null &&
                    FlightModeFeature.GetMode(manager) != FlightModeFeature.FlightModeType.None;
            }
        }

        /// <summary>
        /// Determines if a feature should be enabled based on the current selection.
        /// Checks selection type, shape count, table presence, and hidden shapes requirements.
        /// </summary>
        private bool IsFeatureEnabled(UI.FeatureDisplayInfo info)
        {
            var rule = new SidebarFeatureStateRule
            {
                MinSelection = info.MinSelection,
                MaxSelection = info.MaxSelection,
                RequiredType = ToOfficeSelectionType(info.RequiredType),
                RequiresTable = info.RequiresTable,
                RequiresHiddenShapes = info.RequiresHiddenShapes
            };
            return SidebarStateEvaluator.IsEnabled(rule, _selectionSnapshot);
        }

        private void MainSidebar_Unloaded(object sender, RoutedEventArgs e)
        {
            if (ReferenceEquals(_activeInstance, this)) _activeInstance = null;
            if (_themeSubscribed)
            {
                oPenEfficiency.Services.ThemeManager.ThemeChanged -= ThemeManager_ThemeChanged;
                _themeSubscribed = false;
            }
            foreach (var window in _floatingWindows.Values.ToList())
            {
                try { window.Close(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "MainSidebar.Unloaded.CloseWindow"); }
            }
            _floatingWindows.Clear();
            _navigation?.Dispose();
            _navigation = null;
        }

        internal void SetHost(UI.Panels.ISidebarHost host)
        {
            _host = host;
        }

        private static SidebarSelectionSnapshot CreateSelectionSnapshot(SelectionMetadata metadata)
        {
            if (metadata == null) return null;
            return new SidebarSelectionSnapshot
            {
                Type = ToOfficeSelectionType(metadata.Type),
                ShapeCount = metadata.ShapeCount,
                TextShapeCount = metadata.TextShapeCount,
                TextOrTableShapeCount = metadata.TextOrTableShapeCount,
                HasTextRange = metadata.HasTextRange,
                HasTable = metadata.HasTable,
                HasHiddenShapes = metadata.HasHiddenShapesOnSlide
            };
        }

        internal static bool IsContextCommandAvailable(FeatureCatalogRecord feature)
        {
            var sidebar = _activeInstance;
            if (sidebar == null || feature == null) return false;
            return SidebarStateEvaluator.IsEnabled(
                new SidebarFeatureStateRule
                {
                    MinSelection = feature.MinSelection,
                    MaxSelection = feature.MaxSelection,
                    RequiredType = feature.RequiredType,
                    RequiresTable = feature.RequiresTable,
                    RequiresHiddenShapes = feature.RequiresHiddenShapes
                },
                sidebar._selectionSnapshot);
        }

        /// <summary>
        /// Seam observability: lets the automation seam verify (and if needed create) the
        /// sidebar instance before dispatching a window command, without touching UI state.
        /// </summary>
        internal static bool HasActiveInstance => _activeInstance != null;

        internal static CommandDispatchResult ExecuteContextCommand(string id)
        {
            var sidebar = _activeInstance;
            if (sidebar == null) return CommandDispatchResult.Failed("The openEfficiency Sidebar is not available.");
            var ui = new SidebarCommandUiAdapter(sidebar, null);
            var manager = sidebar.GetManager();
            if (manager == null) return CommandDispatchResult.Failed("PowerPoint is not available.");
            var result = Registry.Value.ExecuteDefault(
                id,
                new CommandContext(manager, CommandInvocationSource.ContextMenu, ui));
            CompleteRegisteredCommand(result, ui);
            return result;
        }

        private static Services.Attributes.OfficeSelectionType ToOfficeSelectionType(
            PowerPoint.PpSelectionType type)
        {
            switch (type)
            {
                case PowerPoint.PpSelectionType.ppSelectionShapes:
                    return Services.Attributes.OfficeSelectionType.Shapes;
                case PowerPoint.PpSelectionType.ppSelectionText:
                    return Services.Attributes.OfficeSelectionType.Text;
                case PowerPoint.PpSelectionType.ppSelectionSlides:
                    return Services.Attributes.OfficeSelectionType.Slides;
                default:
                    return Services.Attributes.OfficeSelectionType.None;
            }
        }

        private static bool ContainsFeature(SidebarConfig config, string id)
        {
            return config?.Sections != null &&
                config.Sections.Any(section =>
                    section?.Features != null &&
                    section.Features.Any(feature =>
                        feature != null &&
                        string.Equals(feature.Id, id, StringComparison.Ordinal)));
        }

        private static void ApplySectionHeaderForeground(TextBlock header, SidebarConfig config)
        {
            if (header == null) return;

            if (SidebarSectionFontColorModes.UsesTheme(config))
            {
                header.SetResourceReference(TextBlock.ForegroundProperty, "SidebarSectionHeaderBrush");
                return;
            }

            try
            {
                header.Foreground = (SolidColorBrush)new BrushConverter().ConvertFrom(config.SectionFontColor);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MainSidebar.ApplySectionHeaderForeground");
                header.SetResourceReference(TextBlock.ForegroundProperty, "SidebarSectionHeaderBrush");
            }
        }

        private void AddSettingsButton(bool includeSettings, bool includeFeatureExplorer)
        {
            if (!includeSettings && !includeFeatureExplorer) return;

            // Vertical layout only - horizontal layout adds Settings inline in RebuildUI
            var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
            SidebarSectionFontColorModes.Normalize(config);
            var header = new TextBlock { Text = "SETTINGS & HELP", Style = (Style)FindResource("SidebarSectionHeader") };
            ApplySectionHeaderForeground(header, config);
            DynamicContentPanel.Children.Add(header);

            // Load icon size from config
            double iconSize = config?.IconSize > 0 ? config.IconSize : 16;

            var stack = new StackPanel { Orientation = Orientation.Horizontal, Margin = (Thickness)FindResource("Space8") };

            if (includeSettings)
            {
                var btn = CreateFeatureButton("BtnSettings", null, iconSize);
                btn.Width = iconSize + 10;
                btn.Height = iconSize + 10;
                btn.Margin = (Thickness)FindResource("Space4");
                stack.Children.Add(btn);
            }

            if (includeFeatureExplorer)
            {
                var exploreBtn = CreateFeatureButton("BtnExploreFeatures", null, iconSize);
                exploreBtn.Width = iconSize + 10;
                exploreBtn.Height = iconSize + 10;
                stack.Children.Add(exploreBtn);
            }

            DynamicContentPanel.Children.Add(stack);
        }

        private void FilterObsoleteFeatures(SidebarConfig config)
        {
            if (config?.Sections == null) return;

            var migration = SidebarLayoutMigration.Migrate(config, UI.FeatureLibrary.AllFeatures);
            if (migration.Changed)
            {
                JsonService.Save(config, "sidebar_config.json");
            }
        }

        private Button CreateFeatureButton(string id, string overrideColor = null, double iconSize = 22)
        {
            var info = UI.FeatureLibrary.GetFeatureInfo(id);
            if (!string.IsNullOrEmpty(overrideColor)) info.Color = overrideColor;
            return UI.Panels.SidebarControlFactory.CreateButton(
                id,
                info,
                _currentBtnBgBrush,
                iconSize,
                _currentIconColorMode,
                FindResource,
                button => ExecuteFeature(id, button, CommandInvocationSource.Sidebar),
                AttachStandardContextMenu);
        }



        private ToggleButton CreateToggleButton(string id, string overrideColor = null, double iconSize = 22)
        {
            var info = UI.FeatureLibrary.GetFeatureInfo(id);
            if (!string.IsNullOrEmpty(overrideColor)) info.Color = overrideColor;
            return UI.Panels.SidebarControlFactory.CreateToggleButton(
                id,
                info,
                _currentBtnBgBrush,
                iconSize,
                _currentIconColorMode,
                FindResource,
                button =>
                {
                    ExecuteFeature(id, button, CommandInvocationSource.Sidebar);
                    UpdateButtonStates();
                },
                AttachStandardContextMenu);
        }

        private void AttachStandardContextMenu(ButtonBase button)
        {
            ContextMenus.Attach(button);
        }

        private void ExecuteFeature(
            string id,
            FrameworkElement targetBtn = null,
            CommandInvocationSource source = CommandInvocationSource.Internal)
        {
            var ui = new SidebarCommandUiAdapter(this, targetBtn);
            var manager = GetManager();
            if (manager == null)
            {
                ExceptionLogger.Log(
                    new InvalidOperationException(
                        $"PowerPoint manager is unavailable for '{id}' from {source}."),
                    "MainSidebar.ExecuteFeature");
                var failure = CommandDispatchResult.Failed(
                    "openEfficiency could not access PowerPoint for this command. Please try again. Details were written to the local diagnostics log.");
                CompleteRegisteredCommand(failure, ui);
                return;
            }

            var context = new CommandContext(manager, source, ui);
            var registryResult = Registry.Value.ExecuteDefault(
                id,
                context,
                GetToggleIntent(source, targetBtn));
            if (registryResult == null ||
                registryResult.Status == CommandDispatchStatus.NotRegistered)
            {
                var exception = new InvalidOperationException(
                    $"Unknown command ID '{id}' from {source}.");
                ExceptionLogger.Log(exception, "MainSidebar.ExecuteFeature.UnknownCommand");
                ui.ShowError($"The command '{id}' is not registered.");
                return;
            }

            CompleteRegisteredCommand(registryResult, ui);
        }

        private void OpenApplyTextEditor()
        {
            Navigation.OpenApplyTextEditor();
        }

        private void ApplyTextFromMaster(bool useLastSelected)
        {
            Navigation.ApplyTextFromMaster(useLastSelected);
        }

        private void OpenSettings(int sectionIndex = 0)
        {
            Navigation.OpenSettings(sectionIndex);
        }

        private void OpenStatusStickerSettings()
        {
            Navigation.OpenSettings(2);
        }

        private static void AddStatusStickerMenuItem(ItemsControl parent, string header, Action action)
        {
            var item = new MenuItem { Header = header };
            item.Click += (s, e) => action?.Invoke();
            parent.Items.Add(item);
        }

        private static string PromptForStatusStickerText()
        {
            using (var form = new System.Windows.Forms.Form())
            using (var label = new System.Windows.Forms.Label())
            using (var textBox = new System.Windows.Forms.TextBox())
            using (var ok = new System.Windows.Forms.Button())
            using (var cancel = new System.Windows.Forms.Button())
            {
                form.Text = "Custom status sticker";
                form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                form.MinimizeBox = false;
                form.MaximizeBox = false;
                form.ClientSize = new System.Drawing.Size(380, 126);

                label.SetBounds(12, 12, 356, 20);
                label.Text = "Sticker label";
                textBox.SetBounds(12, 40, 356, 24);

                ok.SetBounds(200, 82, 80, 28);
                ok.Text = "OK";
                ok.DialogResult = System.Windows.Forms.DialogResult.OK;
                cancel.SetBounds(288, 82, 80, 28);
                cancel.Text = "Cancel";
                cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;

                form.Controls.AddRange(new System.Windows.Forms.Control[] { label, textBox, ok, cancel });
                form.AcceptButton = ok;
                form.CancelButton = cancel;
                return form.ShowDialog() == System.Windows.Forms.DialogResult.OK ? textBox.Text.Trim() : null;
            }
        }

        private void AttachWidgetContextMenu(FrameworkElement widget, string featureId)
        {
            var menu = new ContextMenu();
            if (featureId == "BtnSearchBar")
            {
                var focus = new MenuItem { Header = GetCatalogVariantLabel(featureId, 0, "Focus Command Search") };
                focus.Click += (s, e) => _searchBar?.FocusSearch();
                menu.Items.Add(focus);
                var clear = new MenuItem { Header = GetCatalogVariantLabel(featureId, 1, "Clear Command Search") };
                clear.Click += (s, e) => _searchBar?.SetSearchText(string.Empty);
                menu.Items.Add(clear);
            }
            else if (featureId == "BtnFontPanel")
            {
                var family = new MenuItem { Header = GetCatalogVariantLabel(featureId, 0, "Focus Font Family") };
                family.Click += (s, e) => _fontPanel?.FocusFontFamily();
                menu.Items.Add(family);
                var size = new MenuItem { Header = GetCatalogVariantLabel(featureId, 1, "Focus Font Size") };
                size.Click += (s, e) => _fontPanel?.FocusFontSize();
                menu.Items.Add(size);
            }
            else if (featureId == "BtnSizePanel")
            {
                var width = new MenuItem { Header = GetCatalogVariantLabel(featureId, 0, "Focus Width") };
                width.Click += (s, e) => _sizePanel?.FocusWidth();
                menu.Items.Add(width);
                var height = new MenuItem { Header = GetCatalogVariantLabel(featureId, 1, "Focus Height") };
                height.Click += (s, e) => _sizePanel?.FocusHeight();
                menu.Items.Add(height);
            }
            widget.ContextMenu = menu;
        }

        private void OpenTemplateManager()
        {
            Navigation.OpenTemplateManager();
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
                var app = _host?.Application;
                if (app == null)
                {
                    System.Diagnostics.Debug.WriteLine("GetManager: Sidebar host application is null");
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
            var shortcutEligibleIds = new HashSet<string>(
                UI.FeatureLibrary.Catalog
                    .Where(record => record.SupportsUserShortcut)
                    .Select(record => record.Id),
                StringComparer.Ordinal);
            foreach (var feature in UI.FeatureLibrary.AllFeatures
                .Where(feature => shortcutEligibleIds.Contains(feature.Id)))
            {
                string id = feature.Id;
                ShortcutManager.RegisterHandler(
                    id,
                    () => ExecuteFeature(id, null, CommandInvocationSource.Shortcut));
            }
            ShortcutManager.LoadShortcuts();
        }

        private static string GetCatalogVariantLabel(
            string featureId,
            int variantIndex,
            string fallback)
        {
            var variants = UI.FeatureLibrary.Catalog
                .FirstOrDefault(record =>
                    string.Equals(record.Id, featureId, StringComparison.Ordinal))
                ?.Variants;
            return variants != null &&
                   variantIndex >= 0 &&
                   variantIndex < variants.Count &&
                   !string.IsNullOrWhiteSpace(variants[variantIndex])
                ? variants[variantIndex]
                : fallback;
        }

        /// <summary>
        /// Attach context menu to a feature button. Works for both Button and ToggleButton.
        /// Only BtnSnapToObjects and BtnSnapToGrid use ToggleButton; all others use Button.
        /// </summary>
        private void OpenSpecialCharactersWindow(string selection)
        {
            Navigation.OpenSpecialCharacters(selection);
        }
    }
}
