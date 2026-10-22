using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using oPenEfficiency.UI.Tabs;
using oPenEfficiency.Services;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI
{
    public partial class AssetManagerView : UserControl
    {
        private SlideLibraryTab _slideTab;
        private ImageLibraryTab _imageTab;
        private ShapeLibraryTab _shapeTab;
        private FavoritesTab _favoritesTab;
        private UserControl _activeTab;
        private LocalAssetService _assetService;
        private readonly Microsoft.Office.Interop.PowerPoint.Application _application;

        public AssetManagerView(PowerPointManager ppManager)
        {
            if (ppManager == null) throw new ArgumentNullException(nameof(ppManager));
            InitializeComponent();
            _application = ppManager.GetApplication();
            _assetService = new LocalAssetService(ppManager);

            // Proactively initialize tabs
            _slideTab = new SlideLibraryTab();
            _imageTab = new ImageLibraryTab();
            _shapeTab = new ShapeLibraryTab();
            _favoritesTab = new FavoritesTab();

            // Hide legacy internal tab components
            HideLegacyChrome(_slideTab);
            HideLegacyChrome(_imageTab);
            HideLegacyChrome(_shapeTab);

            // Default to Favorites
            SidebarMenu.SelectedIndex = 0;
            UpdateActiveTab();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            if (MainContent.Content == null)
            {
                UpdateActiveTab();
            }

            // Initial selection check
            try
            {
                if (_application?.ActiveWindow != null)
                {
                    OnSelectionChange(_application.ActiveWindow.Selection);
                }
            }
            catch { /* Asset loading failure leaves the panel empty and available for retry. */ }
        }

        public void OnSelectionChange(Microsoft.Office.Interop.PowerPoint.Selection selection)
        {
            if (selection == null || BtnSaveSelection == null) return;

            bool canSave = false;
            try
            {
                if (selection.Type == PpSelectionType.ppSelectionSlides ||
                    selection.Type == PpSelectionType.ppSelectionShapes ||
                    selection.Type == PpSelectionType.ppSelectionText)
                {
                    canSave = true;
                }
            }
            catch { /* Folder refresh failure retains the last successfully loaded asset list. */ }

            BtnSaveSelection.IsEnabled = canSave;
            BtnSaveSelection.ToolTip = canSave ? "Save current selection to Favorites" : "Select something in PowerPoint to save";
        }

        public void SelectSection(string section)
        {
            if (SidebarMenu == null) return;
            foreach (var itemValue in SidebarMenu.Items)
            {
                var item = itemValue as ListBoxItem;
                if (item != null && string.Equals(item.Tag?.ToString(), section, StringComparison.OrdinalIgnoreCase))
                {
                    SidebarMenu.SelectedItem = item;
                    return;
                }
            }
        }

        public void SaveCurrentSelection()
        {
            BtnSaveSelection_Click(this, new RoutedEventArgs());
        }

        private void HideLegacyChrome(UserControl tab)
        {
            try {
                if (tab.FindName("BrowserToolbar") is FrameworkElement tb) 
                    tb.Visibility = Visibility.Collapsed;
                
                if (tab.FindName("TxtStatus") is FrameworkElement status)
                    status.Visibility = Visibility.Collapsed;
            } catch { /* Status animation cleanup is cosmetic and may race with panel disposal. */ }
        }

        private void SidebarMenu_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateActiveTab();
        }

        private void UpdateActiveTab()
        {
            if (SidebarMenu == null || MainContent == null) return;

            if (SidebarMenu.SelectedItem is ListBoxItem item && item.Tag != null)
            {
                string tag = item.Tag.ToString();
                switch (tag)
                {
                    case "Slides":
                        _activeTab = _slideTab;
                        break;
                    case "Images":
                        _activeTab = _imageTab;
                        break;
                    case "Shapes":
                        _activeTab = _shapeTab;
                        break;
                    case "MyAssets":
                        _activeTab = _favoritesTab;
                        break;
                }
                
                if (_activeTab != null)
                {
                    MainContent.Content = _activeTab;
                    SearchBox.Text = "";

                    // The Favorites tab has no per-library sources to configure, so keep the
                    // Settings affordance visibly disabled there instead of ignoring clicks.
                    if (BtnOpenSettings != null)
                    {
                        bool isFavorites = _activeTab is FavoritesTab;
                        BtnOpenSettings.IsEnabled = !isFavorites;
                        BtnOpenSettings.ToolTip = isFavorites ? "No library settings for Favorites" : "Settings";
                    }
                }
            }
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_activeTab == null) return;
            
            // Delegate search to the active tab's internal search box
            if (_activeTab.FindName("TxtSearch") is TextBox internalSearch)
            {
                internalSearch.Text = SearchBox.Text;
            }
        }

        private void BtnOpenSettings_Click(object sender, RoutedEventArgs e)
        {
            if (_activeTab == null) return;

            // Use the robust toggle method in each tab
            if (_activeTab is SlideLibraryTab s) s.ToggleSettings();
            else if (_activeTab is ShapeLibraryTab sh) sh.ToggleSettings();
            else if (_activeTab is ImageLibraryTab img) img.ToggleSettings();
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            if (_activeTab == null) return;

            if (_activeTab is SlideLibraryTab s) s.RefreshLibraryAsync();
            else if (_activeTab is ShapeLibraryTab sh) sh.RefreshLibraryAsync();
            else if (_activeTab is ImageLibraryTab img) img.RefreshLibraryAsync();
            else if (_activeTab is FavoritesTab f) f.RefreshFavoritesAsync();
        }

        private async void BtnSaveSelection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var app = _application;
                if (app == null || app.ActivePresentation == null) return;

                var selection = app.ActiveWindow.Selection;
                if (selection == null || selection.Type == PpSelectionType.ppSelectionNone)
                {
                    MessageBox.Show("Please select something in PowerPoint to save (Slides or Shapes).", 
                        "No Selection", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                AssetType type;
                string defaultName;

                if (selection.Type == PpSelectionType.ppSelectionSlides)
                {
                    type = AssetType.Slides;
                    defaultName = "NewSlideAsset";
                }
                else if (selection.Type == PpSelectionType.ppSelectionShapes || selection.Type == PpSelectionType.ppSelectionText)
                {
                    type = AssetType.Shapes;
                    defaultName = "NewShapeAsset";
                }
                else
                {
                    MessageBox.Show("Only Slides and Shapes can be saved to Favorites.", 
                        "Invalid Selection", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var dialog = new SaveToMyAssetsDialog(_assetService, type, defaultName);
                new System.Windows.Interop.WindowInteropHelper(dialog).Owner = (IntPtr)app.HWND;
                if (dialog.ShowDialog() == true)
                {
                    SaveResult result;
                    if (type == AssetType.Slides)
                    {
                        result = await _assetService.SaveSelectedSlides(dialog.TargetFolderPath, dialog.FileName, dialog.Tags, dialog.AutoTag);
                    }
                    else
                    {
                        result = await _assetService.SaveSelectedShapes(dialog.TargetFolderPath, dialog.FileName, dialog.Tags, dialog.AutoTag);
                    }

                    if (result.Success)
                    {
                        MessageBox.Show($"Successfully saved to Favorites.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                        
                        // Refresh favorites tab if active
                        if (_activeTab is FavoritesTab f) f.RefreshFavoritesAsync();
                    }
                    else
                    {
                        MessageBox.Show($"Failed to save: {result.ErrorMessage}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
