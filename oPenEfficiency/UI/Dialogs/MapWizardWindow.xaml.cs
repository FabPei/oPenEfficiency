using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Features.Visuals;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class MapWizardWindow : Window
    {
        public MapWizardConfig Config { get; private set; }
        public PowerPoint.Shape MapGroup => _mapGroup;
        private readonly PowerPointManager _manager;
        private readonly IReadOnlyList<MapCatalogEntry> _countryCatalog;
        private readonly ICollectionView _countryView;
        private readonly List<RecentMapEntry> _recentMaps = new List<RecentMapEntry>();
        private PowerPoint.Shape _mapGroup;
        private string _countrySearchText = string.Empty;
        private bool _isUpdatingCountryFilter;
        private bool _isApplyingRecentMap;
        private bool _isInserting;
        private string _administrativeParentCountryId;
        private string _administrativeParentDetailLevel;

        public MapWizardWindow(
            PowerPointManager manager,
            PowerPoint.Shape mapGroup,
            MapWizardConfig existingConfig,
            string initialScope = null,
            bool selectMostRecent = false)
        {
            InitializeComponent();
            _manager = manager;
            var application = _manager?.GetApplication();
            if (application != null)
                new System.Windows.Interop.WindowInteropHelper(this).Owner = new IntPtr(application.HWND);
            SourceInitialized += (sender, args) => FitToOwnerWorkArea();
            Loaded += (sender, args) => UpdateHeaderCaption();
            HeaderDragRegion.SizeChanged += (sender, args) => UpdateHeaderCaption();
            _mapGroup = mapGroup;
            _countryCatalog = MapAssetService.GetCountryCatalog();
            _countryView = CollectionViewSource.GetDefaultView(_countryCatalog);
            _countryView.Filter = FilterCountry;
            CmbCountry.ItemsSource = _countryView;
            CmbCountry.DisplayMemberPath = nameof(MapCatalogEntry.DisplayName);
            CmbCountry.AddHandler(TextBox.TextChangedEvent, new TextChangedEventHandler(CmbCountry_TextChanged));
            if (_countryCatalog.Count > 0)
            {
                CmbCountry.SelectedIndex = 0;
            }

            if (existingConfig == null)
            {
                Config = new MapWizardConfig();
            }
            else
            {
                Config = existingConfig;
                string existingProjection = ShapeMetadataService.GetTag(_mapGroup, MapAssetService.ProjectionTag);
                if (!string.IsNullOrWhiteSpace(existingProjection))
                {
                    Config.Projection = existingProjection.IndexOf("Mercator", StringComparison.OrdinalIgnoreCase) >= 0
                        ? MapProjection.Mercator
                        : MapProjection.Natural;
                }
                LoadConfigIntoUI();

                string existingCountryId = ShapeMetadataService.GetTag(_mapGroup, MapAssetService.CountryIdTag);
                string existingDetailLevel = ShapeMetadataService.GetTag(_mapGroup, MapAssetService.DetailLevelTag);
                if (!string.IsNullOrWhiteSpace(existingCountryId))
                {
                    foreach (ComboBoxItem item in CmbMapScope.Items)
                    {
                        if (string.Equals(item.Tag?.ToString(), "Country", StringComparison.OrdinalIgnoreCase))
                        {
                            CmbMapScope.SelectedItem = item;
                            break;
                        }
                    }
                    CmbCountry.SelectedItem = _countryCatalog.FirstOrDefault(item =>
                        string.Equals(item.Id, existingCountryId, StringComparison.OrdinalIgnoreCase));
                    UpdateDetailControls();
                    SelectDetailLevel(existingDetailLevel);
                    string existingRemoteMode = ShapeMetadataService.GetTag(_mapGroup, MapAssetService.RemoteRegionModeTag);
                    if (!Enum.TryParse(existingRemoteMode, true, out MapRemoteRegionMode remoteRegionMode) &&
                        string.Equals(ShapeMetadataService.GetTag(_mapGroup, MapAssetService.CompactInsetsTag), "True", StringComparison.OrdinalIgnoreCase))
                    {
                        remoteRegionMode = MapRemoteRegionMode.CompactInsets;
                    }
                    Config.RemoteRegionMode = remoteRegionMode;
                    SelectRemoteRegionMode(remoteRegionMode);
                }
            }

            UpdateSwatches();
            LoadRecentMaps();
            if (!string.IsNullOrWhiteSpace(initialScope)) SelectScope(initialScope);
            if (selectMostRecent && _recentMaps.Count > 0) CmbRecentMaps.SelectedIndex = 0;
            UpdateScopeControls();
            UpdateMapStatus();
        }

        private void SelectScope(string scope)
        {
            foreach (ComboBoxItem item in CmbMapScope.Items)
            {
                if (string.Equals(item.Tag?.ToString(), scope, StringComparison.OrdinalIgnoreCase))
                {
                    CmbMapScope.SelectedItem = item;
                    return;
                }
            }
        }

        private void FitToOwnerWorkArea()
        {
            var helper = new System.Windows.Interop.WindowInteropHelper(this);
            var screen = System.Windows.Forms.Screen.FromHandle(helper.Owner);
            var dpi = VisualTreeHelper.GetDpi(this);
            double availableWidth = Math.Max(1, (screen.WorkingArea.Width / dpi.DpiScaleX) - 32);
            double availableHeight = Math.Max(1, (screen.WorkingArea.Height / dpi.DpiScaleY) - 32);
            MinWidth = Math.Min(MinWidth, availableWidth);
            MinHeight = Math.Min(MinHeight, availableHeight);
            MaxWidth = availableWidth;
            MaxHeight = availableHeight;
            Width = Math.Min(Width, availableWidth);
            Height = Math.Min(Height, availableHeight);
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            // Let ComboBox/DataGrid consume Escape first; insertion has no cancel path.
            if (e.Key == Key.Escape && !_isInserting)
            {
                BtnClose_Click(sender, e);
                e.Handled = true;
            }
        }

        private void UpdateHeaderCaption()
        {
            var chrome = System.Windows.Shell.WindowChrome.GetWindowChrome(this);
            if (chrome == null || HeaderDragRegion.ActualHeight <= 0) return;
            // Native caption hit testing includes the shell's top padding and follows
            // header wrapping/DPI layout. Close remains an interactive client control.
            chrome.CaptionHeight = HeaderDragRegion.TranslatePoint(
                new Point(0, HeaderDragRegion.ActualHeight), this).Y;
        }

        private void LoadConfigIntoUI()
        {
            if (Config.Mode == MapColorMode.Continuous)
                CmbMode.SelectedIndex = 0;
            else
                CmbMode.SelectedIndex = 1;

            TxtLowColor.Text = Config.LowColorHex;
            TxtHighColor.Text = Config.HighColorHex;
            CmbProjection.SelectedIndex = Config.Projection == MapProjection.Mercator ? 0 : 1;

            if (Config.Data.Count > 0)
            {
                var lines = Config.Data.Select(d => $"{d.Region}\t{d.RawValue}");
                TxtPasteArea.Text = string.Join(Environment.NewLine, lines);
                RefreshDataGrid();
            }
        }

        private void CmbMode_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SpHeatmap == null) return;

            var selected = (ComboBoxItem)CmbMode.SelectedItem;
            if (selected.Tag.ToString() == "Continuous")
            {
                SpHeatmap.Visibility = Visibility.Visible;
                TxtCategoricalNotice.Visibility = Visibility.Collapsed;
                Config.Mode = MapColorMode.Continuous;
            }
            else
            {
                SpHeatmap.Visibility = Visibility.Collapsed;
                TxtCategoricalNotice.Visibility = Visibility.Visible;
                Config.Mode = MapColorMode.Categorical;
            }
        }

        private void BtnParseData_Click(object sender, RoutedEventArgs e)
        {
            RefreshDataGrid();
        }
        
        private void TxtPasteArea_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Optional: auto-refresh on paste
        }

        private void RefreshDataGrid()
        {
            Config.Data.Clear();
            var lines = TxtPasteArea.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            var availableShapeNames = MapWizardFeature.GetAvailableRegionKeys(_mapGroup);

            foreach (var line in lines)
            {
                var parts = line.Split(new[] { '\t' }, StringSplitOptions.None);
                if (parts.Length >= 2)
                {
                    string region = parts[0].Trim();
                    string valStr = parts[1].Trim();
                    
                    double numVal = 0;
                    if (!double.TryParse(valStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out numVal))
                    {
                        double.TryParse(valStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out numVal);
                    }

                    Config.Data.Add(new MapDataRow
                    {
                        Region = region,
                        RawValue = valStr,
                        NumericValue = numVal,
                        IsMatched = availableShapeNames.Contains(region)
                    });
                }
            }

            DgDataPreview.ItemsSource = null;
            DgDataPreview.ItemsSource = Config.Data;
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            if (_mapGroup == null)
            {
                MessageBox.Show("Insert a map first, or close the builder and select an existing map shape.", "Map Builder", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            RefreshDataGrid();
            Config.LowColorHex = TxtLowColor.Text;
            Config.HighColorHex = TxtHighColor.Text;
            Config.Projection = GetSelectedProjection();
            
            this.DialogResult = true;
            this.Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void BtnInsertMap_Click(object sender, RoutedEventArgs e)
        {
            string mapScope = (CmbMapScope.SelectedItem as ComboBoxItem)?.Tag?.ToString();
            MapCatalogEntry selectedCountry = null;
            string selectedValue = string.Empty;
            bool largestLandmassOnly = ChkLargestLandmassOnly.IsChecked == true;
            MapProjection projection = GetSelectedProjection();
            string detailLevel = mapScope == "Country" ? GetSelectedDetailLevel() : "ADM0";
            string administrativeDistrictStateId = (detailLevel == "ADM2" || detailLevel == "ADM3")
                ? (CmbAdministrativeDistrictState.SelectedItem as ComboBoxItem)?.Tag?.ToString()
                : null;
            MapRemoteRegionMode remoteRegionMode = mapScope == "Country" && detailLevel == "ADM1"
                ? GetSelectedRemoteRegionMode()
                : MapRemoteRegionMode.Geographic;
            if (mapScope != "Country" || detailLevel == "ADM1")
            {
                largestLandmassOnly = false;
            }

            if (mapScope == "Country")
            {
                selectedCountry = ResolveSelectedCountry();
                if (selectedCountry == null)
                {
                    MessageBox.Show("Choose a country or territory from the filtered list.", "Map Builder", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                selectedValue = selectedCountry.Id;
                if (detailLevel == "ADM1" && MapAssetService.GetAdministrativeRegionCount(selectedValue) == 0)
                {
                    MessageBox.Show(
                        "No states or provinces are available for the selected country in the bundled Natural Earth data.",
                        "Map Builder",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }
                if (detailLevel == "ADM2" && MapAssetService.GetSecondLevelAdministrativeRegionCount(selectedValue) == 0)
                {
                    MessageBox.Show("No second-level regions are available for the selected country in the bundled geoBoundaries data.", "Map Builder", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                if (detailLevel == "ADM3" && MapAssetService.GetThirdLevelAdministrativeRegionCount(selectedValue) == 0)
                {
                    MessageBox.Show("No third-level regions are available for the selected country in the bundled geoBoundaries data.", "Map Builder", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                if (detailLevel == "ADM3" && MapAssetService.GetGermanyDistrictCount(administrativeDistrictStateId) == 0)
                {
                    MessageBox.Show("Choose a federal state with available districts. Berlin has no separate ADM3 units in the bundled source.", "Map Builder", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
            }
            else if (mapScope == "Continent")
            {
                selectedValue = (CmbContinent.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty;
                if (projection == MapProjection.Mercator && string.Equals(selectedValue, "Antarctica", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(
                        "Mercator cannot represent the poles. Antarctica will use Natural proportions.",
                        "Map Builder",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    projection = MapProjection.Natural;
                    CmbProjection.SelectedIndex = 1;
                }
            }

            PowerPoint.Shape insertedMap = null;
            SetInsertionBusy(true, "Inserting editable map shapes…");
            Dispatcher.Invoke(DispatcherPriority.Render, new Action(() => { }));
            try
            {
                switch (mapScope)
                {
                    case "Continent":
                        insertedMap = MapAssetService.InsertContinentCountries(_manager, selectedValue, projection);
                        break;
                    case "Country":
                        insertedMap = detailLevel == "ADM1"
                            ? MapAssetService.InsertCountryAdministrativeRegions(_manager, selectedCountry.Id, projection, remoteRegionMode)
                            : detailLevel == "ADM2"
                                ? string.Equals(selectedCountry.Id, "DEU", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(administrativeDistrictStateId)
                                    ? MapAssetService.InsertGermanyAdministrativeDistricts(_manager, administrativeDistrictStateId, projection)
                                    : MapAssetService.InsertCountrySecondLevelAdministrativeRegions(_manager, selectedCountry.Id, administrativeDistrictStateId, projection)
                                : detailLevel == "ADM3"
                                    ? string.Equals(selectedCountry.Id, "DEU", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(administrativeDistrictStateId)
                                        ? MapAssetService.InsertGermanyDistricts(_manager, administrativeDistrictStateId, projection)
                                        : MapAssetService.InsertCountryThirdLevelAdministrativeRegions(_manager, selectedCountry.Id, administrativeDistrictStateId, projection)
                                : MapAssetService.InsertCountry(_manager, selectedCountry.Id, largestLandmassOnly, projection);
                        break;
                    default:
                        insertedMap = MapAssetService.InsertWorldCountries(_manager, projection);
                        break;
                }
            }
            finally
            {
                SetInsertionBusy(false, null);
            }

            if (insertedMap == null)
            {
                MessageBox.Show("The selected map could not be inserted. Make sure a normal slide is active.", "Map Builder", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _mapGroup = insertedMap;
            Config = new MapWizardConfig { Projection = projection, RemoteRegionMode = remoteRegionMode };
            TxtPasteArea.Clear();
            DgDataPreview.ItemsSource = null;
            TxtLowColor.Text = Config.LowColorHex;
            TxtHighColor.Text = Config.HighColorHex;
            CmbProjection.SelectedIndex = Config.Projection == MapProjection.Mercator ? 0 : 1;
            UpdateSwatches();
            RememberRecentMap(mapScope ?? "World", selectedValue, largestLandmassOnly, projection, detailLevel, remoteRegionMode, administrativeDistrictStateId);
            UpdateMapStatus();
        }

        private bool FilterCountry(object item)
        {
            if (!(item is MapCatalogEntry entry) || string.IsNullOrWhiteSpace(_countrySearchText))
            {
                return true;
            }

            string search = _countrySearchText.Trim();
            return entry.DisplayName.IndexOf(search, StringComparison.CurrentCultureIgnoreCase) >= 0 ||
                   entry.LocalName.IndexOf(search, StringComparison.CurrentCultureIgnoreCase) >= 0 ||
                   entry.Id.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void CmbCountry_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingCountryFilter || !CmbCountry.IsKeyboardFocusWithin)
            {
                return;
            }

            _isUpdatingCountryFilter = true;
            try
            {
                _countrySearchText = CmbCountry.Text ?? string.Empty;
                _countryView.Refresh();
                CmbCountry.IsDropDownOpen = true;
            }
            finally
            {
                _isUpdatingCountryFilter = false;
            }
        }

        private void SetInsertionBusy(bool isBusy, string status)
        {
            _isInserting = isBusy;
            BtnInsertMap.IsEnabled = !isBusy;
            CmbMapScope.IsEnabled = !isBusy;
            CmbContinent.IsEnabled = !isBusy;
            CmbCountry.IsEnabled = !isBusy;
            CmbDetailLevel.IsEnabled = !isBusy;
            CmbProjection.IsEnabled = !isBusy;
            CmbRecentMaps.IsEnabled = !isBusy && _recentMaps.Count > 0;
            ChkLargestLandmassOnly.IsEnabled = !isBusy;
            CmbRemoteRegionMode.IsEnabled = !isBusy;
            MapInsertProgress.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;
            Mouse.OverrideCursor = isBusy ? Cursors.Wait : null;

            if (isBusy)
            {
                TxtMapStatus.Text = status;
            }
            else
            {
                UpdateScopeControls();
                UpdateMapStatus();
            }
        }

        private void CmbMapScope_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateScopeControls();
        }

        private void CmbCountry_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateDetailControls();
        }

        private void CmbDetailLevel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateDetailControls();
        }

        private void UpdateScopeControls()
        {
            if (CmbMapScope == null || TxtWorldSelection == null || CmbContinent == null || CmbCountry == null || ChkLargestLandmassOnly == null)
            {
                return;
            }

            string scope = (CmbMapScope.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "World";
            TxtWorldSelection.Visibility = scope == "World" ? Visibility.Visible : Visibility.Collapsed;
            CmbContinent.Visibility = scope == "Continent" ? Visibility.Visible : Visibility.Collapsed;
            bool isCountry = scope == "Country";
            CmbCountry.Visibility = isCountry ? Visibility.Visible : Visibility.Collapsed;
            CmbDetailLevel.IsEnabled = !_isInserting && isCountry;
            if (!_isInserting)
            {
                CmbContinent.IsEnabled = true;
                CmbCountry.IsEnabled = true;
            }
            UpdateDetailControls();
        }

        private void UpdateDetailControls()
        {
            if (CmbMapScope == null || CmbDetailLevel == null || CmbDetailAdmin1 == null || CmbDetailAdmin2 == null || CmbDetailAdmin3 == null || ChkLargestLandmassOnly == null ||
                TxtRemoteRegionMode == null || CmbRemoteRegionMode == null || TxtAdministrativeDistrictState == null || CmbAdministrativeDistrictState == null)
            {
                return;
            }

            string scope = (CmbMapScope.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "World";
            bool isCountry = scope == "Country";
            MapCatalogEntry country = isCountry ? ResolveSelectedCountry() : null;
            RefreshAdministrativeParentRegionCatalog(country, GetSelectedDetailLevel());
            int administrativeRegionCount = country == null ? 0 : MapAssetService.GetAdministrativeRegionCount(country.Id);
            CmbDetailAdmin1.IsEnabled = administrativeRegionCount > 0;
            CmbDetailAdmin1.ToolTip = administrativeRegionCount > 0
                ? $"Insert {administrativeRegionCount} editable first-level administrative regions."
                : "No first-level administrative regions are available for this country.";
            int secondLevelRegionCount = country == null ? 0 : MapAssetService.GetSecondLevelAdministrativeRegionCount(country.Id);
            CmbDetailAdmin2.IsEnabled = secondLevelRegionCount > 0;
            CmbDetailAdmin2.ToolTip = secondLevelRegionCount > 0
                ? $"Insert {secondLevelRegionCount} editable second-level administrative regions."
                : "No second-level regions are available for this country in the bundled geoBoundaries data.";
            int thirdLevelRegionCount = country == null ? 0 : MapAssetService.GetThirdLevelAdministrativeRegionCount(country.Id);
            CmbDetailAdmin3.IsEnabled = thirdLevelRegionCount > 0;
            CmbDetailAdmin3.ToolTip = thirdLevelRegionCount > 0
                ? $"Insert {thirdLevelRegionCount} editable third-level administrative regions."
                : "No third-level regions are available for this country in the bundled geoBoundaries data.";

            if ((!CmbDetailAdmin1.IsEnabled && GetSelectedDetailLevel() == "ADM1") ||
                (!CmbDetailAdmin2.IsEnabled && GetSelectedDetailLevel() == "ADM2") ||
                (!CmbDetailAdmin3.IsEnabled && GetSelectedDetailLevel() == "ADM3"))
            {
                CmbDetailLevel.SelectedIndex = 0;
            }

            bool isCountryOutline = isCountry && GetSelectedDetailLevel() == "ADM0";
            bool isAdministrativeMap = isCountry && GetSelectedDetailLevel() == "ADM1" && CmbDetailAdmin1.IsEnabled;
            bool isSubregionMap = isCountry && country != null && CmbAdministrativeDistrictState.Items.Count > 1 &&
                ((GetSelectedDetailLevel() == "ADM2" && CmbDetailAdmin2.IsEnabled) ||
                 (GetSelectedDetailLevel() == "ADM3" && CmbDetailAdmin3.IsEnabled));
            bool isGermanyCountyMap = isSubregionMap && string.Equals(country.Id, "DEU", StringComparison.OrdinalIgnoreCase) &&
                GetSelectedDetailLevel() == "ADM3";
            ChkLargestLandmassOnly.Visibility = isCountryOutline ? Visibility.Visible : Visibility.Collapsed;
            ChkLargestLandmassOnly.IsEnabled = !_isInserting && isCountryOutline;
            TxtRemoteRegionMode.Visibility = isAdministrativeMap ? Visibility.Visible : Visibility.Collapsed;
            CmbRemoteRegionMode.Visibility = isAdministrativeMap ? Visibility.Visible : Visibility.Collapsed;
            CmbRemoteRegionMode.IsEnabled = !_isInserting && isAdministrativeMap;
            TxtAdministrativeDistrictState.Visibility = isSubregionMap ? Visibility.Visible : Visibility.Collapsed;
            CmbAdministrativeDistrictState.Visibility = isSubregionMap ? Visibility.Visible : Visibility.Collapsed;
            CmbAdministrativeDistrictState.IsEnabled = !_isInserting && isSubregionMap;
            if (isGermanyCountyMap && string.IsNullOrWhiteSpace((CmbAdministrativeDistrictState.SelectedItem as ComboBoxItem)?.Tag?.ToString()))
            {
                CmbAdministrativeDistrictState.SelectedIndex = 1;
            }
        }

        private void RefreshAdministrativeParentRegionCatalog(MapCatalogEntry country, string detailLevel)
        {
            if (country == null || (string.Equals(_administrativeParentCountryId, country.Id, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(_administrativeParentDetailLevel, detailLevel, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            string previousId = (CmbAdministrativeDistrictState.SelectedItem as ComboBoxItem)?.Tag?.ToString();
            CmbAdministrativeDistrictState.Items.Clear();
            CmbAdministrativeDistrictState.Items.Add(new ComboBoxItem { Content = $"All {country.DisplayName}", Tag = string.Empty });
            foreach (MapCatalogEntry region in (detailLevel == "ADM2" || detailLevel == "ADM3"
                ? MapAssetService.GetLowerLevelParentRegionCatalog(country.Id, detailLevel)
                : new List<MapCatalogEntry>()))
            {
                CmbAdministrativeDistrictState.Items.Add(new ComboBoxItem { Content = region.DisplayName, Tag = region.Id });
            }
            _administrativeParentCountryId = country.Id;
            _administrativeParentDetailLevel = detailLevel;
            SelectAdministrativeDistrictState(previousId);
        }

        private void LoadRecentMaps()
        {
            _recentMaps.Clear();
            string stored = Properties.Settings.Default.RecentMapSelections ?? string.Empty;
            foreach (string token in stored.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string[] parts = token.Split('|');
                if (parts.Length < 3 || parts.Length > 7)
                {
                    continue;
                }

                try
                {
                    string scope = parts[0];
                    string value = Uri.UnescapeDataString(parts[1]);
                    bool largest = parts[2] == "1";
                    MapProjection projection = parts.Length >= 4 && string.Equals(parts[3], "Mercator", StringComparison.OrdinalIgnoreCase)
                        ? MapProjection.Mercator
                        : MapProjection.Natural;
                    string detailLevel = parts.Length >= 5 && string.Equals(parts[4], "ADM3", StringComparison.OrdinalIgnoreCase)
                        ? "ADM3"
                        : parts.Length >= 5 && string.Equals(parts[4], "ADM2", StringComparison.OrdinalIgnoreCase)
                            ? "ADM2"
                            : parts.Length >= 5 && string.Equals(parts[4], "ADM1", StringComparison.OrdinalIgnoreCase) ? "ADM1" : "ADM0";
                    MapRemoteRegionMode remoteRegionMode = ParseStoredRemoteRegionMode(parts.Length >= 6 ? parts[5] : null);
                    string administrativeDistrictStateId = parts.Length == 7 ? Uri.UnescapeDataString(parts[6]) : null;
                    string displayName = BuildRecentDisplayName(scope, value, largest, projection, detailLevel, remoteRegionMode, administrativeDistrictStateId);
                    if (!string.IsNullOrWhiteSpace(displayName))
                    {
                        _recentMaps.Add(new RecentMapEntry(scope, value, largest, projection, detailLevel, remoteRegionMode, administrativeDistrictStateId, displayName));
                    }
                }
                catch
                {
                    // Ignore obsolete or malformed user settings.
                }
            }

            CmbRecentMaps.ItemsSource = null;
            CmbRecentMaps.ItemsSource = _recentMaps;
            CmbRecentMaps.IsEnabled = _recentMaps.Count > 0;
            CmbRecentMaps.SelectedIndex = -1;
        }

        private void RememberRecentMap(string scope, string value, bool largestLandmassOnly, MapProjection projection, string detailLevel, MapRemoteRegionMode remoteRegionMode, string administrativeDistrictStateId = null)
        {
            string displayName = BuildRecentDisplayName(scope, value, largestLandmassOnly, projection, detailLevel, remoteRegionMode, administrativeDistrictStateId);
            if (string.IsNullOrWhiteSpace(displayName))
            {
                return;
            }

            _recentMaps.RemoveAll(item =>
                string.Equals(item.Scope, scope, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(item.Value, value, StringComparison.OrdinalIgnoreCase) &&
                item.LargestLandmassOnly == largestLandmassOnly &&
                item.Projection == projection &&
                string.Equals(item.DetailLevel, detailLevel, StringComparison.OrdinalIgnoreCase) &&
                item.RemoteRegionMode == remoteRegionMode &&
                string.Equals(item.AdministrativeDistrictStateId, administrativeDistrictStateId, StringComparison.OrdinalIgnoreCase));
            _recentMaps.Insert(0, new RecentMapEntry(scope, value, largestLandmassOnly, projection, detailLevel, remoteRegionMode, administrativeDistrictStateId, displayName));
            if (_recentMaps.Count > 5)
            {
                _recentMaps.RemoveRange(5, _recentMaps.Count - 5);
            }

            Properties.Settings.Default.RecentMapSelections = string.Join(";", _recentMaps.Select(item =>
                $"{item.Scope}|{Uri.EscapeDataString(item.Value ?? string.Empty)}|{(item.LargestLandmassOnly ? "1" : "0")}|{item.Projection}|{item.DetailLevel}|{item.RemoteRegionMode}|{Uri.EscapeDataString(item.AdministrativeDistrictStateId ?? string.Empty)}"));
            Properties.Settings.Default.Save();
            CmbRecentMaps.ItemsSource = null;
            CmbRecentMaps.ItemsSource = _recentMaps;
            CmbRecentMaps.IsEnabled = true;
            CmbRecentMaps.SelectedIndex = -1;
        }

        private string BuildRecentDisplayName(string scope, string value, bool largestLandmassOnly, MapProjection projection, string detailLevel, MapRemoteRegionMode remoteRegionMode, string administrativeDistrictStateId = null)
        {
            string projectionSuffix = projection == MapProjection.Mercator ? " · Mercator" : " · Natural";
            if (string.Equals(scope, "World", StringComparison.OrdinalIgnoreCase))
            {
                return "World" + projectionSuffix;
            }

            if (string.Equals(scope, "Continent", StringComparison.OrdinalIgnoreCase))
            {
                return string.IsNullOrWhiteSpace(value) ? null : $"Continent · {value}{projectionSuffix}";
            }

            if (string.Equals(scope, "Country", StringComparison.OrdinalIgnoreCase))
            {
                MapCatalogEntry country = _countryCatalog.FirstOrDefault(item => string.Equals(item.Id, value, StringComparison.OrdinalIgnoreCase));
                string detailSuffix = string.Equals(detailLevel, "ADM1", StringComparison.OrdinalIgnoreCase)
                    ? $" · states / provinces{GetRemoteRegionModeSuffix(remoteRegionMode)}"
                    : string.Equals(detailLevel, "ADM2", StringComparison.OrdinalIgnoreCase)
                        ? $" · government districts{GetAdministrativeDistrictStateSuffix(administrativeDistrictStateId)}"
                        : string.Equals(detailLevel, "ADM3", StringComparison.OrdinalIgnoreCase)
                            ? $" · districts / independent cities{GetAdministrativeDistrictStateSuffix(administrativeDistrictStateId)}"
                        : largestLandmassOnly ? " · largest landmass" : string.Empty;
                return country == null ? null : $"Country · {country.DisplayName}{detailSuffix}{projectionSuffix}";
            }

            return null;
        }

        private void CmbRecentMaps_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isApplyingRecentMap || !(CmbRecentMaps.SelectedItem is RecentMapEntry recent))
            {
                return;
            }

            _isApplyingRecentMap = true;
            try
            {
                foreach (ComboBoxItem item in CmbMapScope.Items)
                {
                    if (string.Equals(item.Tag?.ToString(), recent.Scope, StringComparison.OrdinalIgnoreCase))
                    {
                        CmbMapScope.SelectedItem = item;
                        break;
                    }
                }

                if (recent.Scope == "Continent")
                {
                    foreach (ComboBoxItem item in CmbContinent.Items)
                    {
                        if (string.Equals(item.Content?.ToString(), recent.Value, StringComparison.OrdinalIgnoreCase))
                        {
                            CmbContinent.SelectedItem = item;
                            break;
                        }
                    }
                }
                else if (recent.Scope == "Country")
                {
                    _isUpdatingCountryFilter = true;
                    _countrySearchText = string.Empty;
                    _countryView.Refresh();
                    CmbCountry.SelectedItem = _countryCatalog.FirstOrDefault(item => string.Equals(item.Id, recent.Value, StringComparison.OrdinalIgnoreCase));
                    _isUpdatingCountryFilter = false;
                    ChkLargestLandmassOnly.IsChecked = recent.LargestLandmassOnly;
                    UpdateDetailControls();
                    SelectDetailLevel(recent.DetailLevel);
                    SelectAdministrativeDistrictState(recent.AdministrativeDistrictStateId);
                    SelectRemoteRegionMode(recent.RemoteRegionMode);
                }

                CmbProjection.SelectedIndex = recent.Projection == MapProjection.Mercator ? 0 : 1;

                UpdateScopeControls();
            }
            finally
            {
                _isUpdatingCountryFilter = false;
                _isApplyingRecentMap = false;
            }
        }

        private MapCatalogEntry ResolveSelectedCountry()
        {
            if (CmbCountry.SelectedItem is MapCatalogEntry selected)
            {
                return selected;
            }

            string typed = CmbCountry.Text?.Trim();
            if (string.IsNullOrWhiteSpace(typed))
            {
                return null;
            }

            return _countryCatalog.FirstOrDefault(entry =>
                string.Equals(entry.Id, typed, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(entry.LocalName, typed, StringComparison.CurrentCultureIgnoreCase) ||
                string.Equals(entry.DisplayName, typed, StringComparison.CurrentCultureIgnoreCase) ||
                entry.DisplayName.StartsWith(typed, StringComparison.CurrentCultureIgnoreCase));
        }

        private MapProjection GetSelectedProjection()
        {
            string projection = (CmbProjection.SelectedItem as ComboBoxItem)?.Tag?.ToString();
            return string.Equals(projection, "Natural", StringComparison.OrdinalIgnoreCase)
                ? MapProjection.Natural
                : MapProjection.Mercator;
        }

        private MapRemoteRegionMode GetSelectedRemoteRegionMode()
        {
            string mode = (CmbRemoteRegionMode.SelectedItem as ComboBoxItem)?.Tag?.ToString();
            return Enum.TryParse(mode, true, out MapRemoteRegionMode parsed)
                ? parsed
                : MapRemoteRegionMode.Geographic;
        }

        private void SelectRemoteRegionMode(MapRemoteRegionMode mode)
        {
            foreach (ComboBoxItem item in CmbRemoteRegionMode.Items)
            {
                if (string.Equals(item.Tag?.ToString(), mode.ToString(), StringComparison.OrdinalIgnoreCase))
                {
                    CmbRemoteRegionMode.SelectedItem = item;
                    return;
                }
            }
            CmbRemoteRegionMode.SelectedIndex = 0;
        }

        private static MapRemoteRegionMode ParseStoredRemoteRegionMode(string stored)
        {
            if (stored == "1")
            {
                return MapRemoteRegionMode.CompactInsets;
            }
            if (stored == "0" || string.IsNullOrWhiteSpace(stored))
            {
                return MapRemoteRegionMode.Geographic;
            }
            return Enum.TryParse(stored, true, out MapRemoteRegionMode parsed)
                ? parsed
                : MapRemoteRegionMode.Geographic;
        }

        private static string GetRemoteRegionModeSuffix(MapRemoteRegionMode mode)
        {
            switch (mode)
            {
                case MapRemoteRegionMode.CompactInsets:
                    return " · compact insets";
                case MapRemoteRegionMode.Exclude:
                    return " · excludes remote regions";
                default:
                    return string.Empty;
            }
        }

        private string GetSelectedDetailLevel()
        {
            string selected = (CmbDetailLevel?.SelectedItem as ComboBoxItem)?.Tag?.ToString();
            return string.Equals(selected, "ADM3", StringComparison.OrdinalIgnoreCase)
                ? "ADM3"
                : string.Equals(selected, "ADM2", StringComparison.OrdinalIgnoreCase)
                    ? "ADM2"
                : string.Equals(selected, "ADM1", StringComparison.OrdinalIgnoreCase)
                    ? "ADM1"
                    : "ADM0";
        }

        private string GetAdministrativeDistrictStateSuffix(string stateId)
        {
            if (string.IsNullOrWhiteSpace(stateId))
            {
                return string.Empty;
            }

            MapCatalogEntry state = MapAssetService.GetGermanyStateCatalog()
                .FirstOrDefault(item => string.Equals(item.Id, stateId, StringComparison.OrdinalIgnoreCase));
            return state == null ? string.Empty : $" · {state.DisplayName}";
        }

        private void SelectDetailLevel(string detailLevel)
        {
            foreach (ComboBoxItem item in CmbDetailLevel.Items)
            {
                if (string.Equals(item.Tag?.ToString(), detailLevel, StringComparison.OrdinalIgnoreCase) && item.IsEnabled)
                {
                    CmbDetailLevel.SelectedItem = item;
                    return;
                }
            }
            CmbDetailLevel.SelectedIndex = 0;
        }

        private void SelectAdministrativeDistrictState(string stateId)
        {
            foreach (ComboBoxItem item in CmbAdministrativeDistrictState.Items)
            {
                if (string.Equals(item.Tag?.ToString(), stateId ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                {
                    CmbAdministrativeDistrictState.SelectedItem = item;
                    return;
                }
            }
            CmbAdministrativeDistrictState.SelectedIndex = 0;
        }

        private void UpdateMapStatus()
        {
            if (TxtMapStatus == null || BtnApply == null)
            {
                return;
            }

            if (_mapGroup == null)
            {
                TxtMapStatus.Text = "No map selected. Choose a scope and insert a bundled map, or select an existing map shape before opening the wizard.";
                BtnApply.IsEnabled = false;
            }
            else
            {
                string mapName = string.IsNullOrWhiteSpace(_mapGroup.AlternativeText) ? _mapGroup.Name : _mapGroup.AlternativeText;
                TxtMapStatus.Text = $"Active map: {mapName}";
                BtnApply.IsEnabled = true;
            }
        }
        
        private void BtnPickLowColor_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new WpfColorPickerDialog(TxtLowColor.Text) { Owner = this };
            if (dialog.ShowDialog() == true)
            {
                TxtLowColor.Text = $"#{dialog.SelectedColor.R:X2}{dialog.SelectedColor.G:X2}{dialog.SelectedColor.B:X2}";
                UpdateSwatches();
            }
        }
        
        private void BtnPickHighColor_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new WpfColorPickerDialog(TxtHighColor.Text) { Owner = this };
            if (dialog.ShowDialog() == true)
            {
                TxtHighColor.Text = $"#{dialog.SelectedColor.R:X2}{dialog.SelectedColor.G:X2}{dialog.SelectedColor.B:X2}";
                UpdateSwatches();
            }
        }
        
        private void UpdateSwatches()
        {
            try {
                if (TxtLowColor != null && !string.IsNullOrEmpty(TxtLowColor.Text))
                    BtnPickLow.Background = (SolidColorBrush)new BrushConverter().ConvertFromString(TxtLowColor.Text);
                
                if (TxtHighColor != null && !string.IsNullOrEmpty(TxtHighColor.Text))
                    BtnPickHigh.Background = (SolidColorBrush)new BrushConverter().ConvertFromString(TxtHighColor.Text);
            } catch { /* Invalid color text keeps the last valid preview brush. */ }
        }

        private sealed class RecentMapEntry
        {
            public RecentMapEntry(string scope, string value, bool largestLandmassOnly, MapProjection projection, string detailLevel, MapRemoteRegionMode remoteRegionMode, string administrativeDistrictStateId, string displayName)
            {
                Scope = scope;
                Value = value;
                LargestLandmassOnly = largestLandmassOnly;
                Projection = projection;
                DetailLevel = detailLevel;
                RemoteRegionMode = remoteRegionMode;
                AdministrativeDistrictStateId = administrativeDistrictStateId;
                DisplayName = displayName;
            }

            public string Scope { get; }
            public string Value { get; }
            public bool LargestLandmassOnly { get; }
            public MapProjection Projection { get; }
            public string DetailLevel { get; }
            public MapRemoteRegionMode RemoteRegionMode { get; }
            public string AdministrativeDistrictStateId { get; }
            public string DisplayName { get; }
        }
    }
}
