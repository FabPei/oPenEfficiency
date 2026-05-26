using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using oPenEfficiency.Features;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.UI.Dialogs;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI.Tabs
{
    public partial class ShapeLibraryTab : UserControl
    {
        private bool _isRefreshing = false;
        private readonly ShapeLibraryManager _libraryManager;
        private readonly PowerPointManager _ppManager;
        private readonly AssetCacheService _cacheService;
        private readonly LocalAssetService _assetService;

        public ObservableCollection<ShapeItemDisplay> AllShapes { get; set; } = new ObservableCollection<ShapeItemDisplay>();
        public ObservableCollection<ShapeItemDisplay> FilteredShapes { get; set; } = new ObservableCollection<ShapeItemDisplay>();

        public ObservableCollection<string> LibraryFolders { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<string> SharedFolders { get; set; } = new ObservableCollection<string>();

        private List<ShapeFile> _scannedFiles = new List<ShapeFile>();
        private string _selectedFolderPath = null;
        private bool _isGenerating = false;
        private bool _isInitialized = false;

        public ShapeLibraryTab()
        {
            InitializeComponent();

            _ppManager = new PowerPointManager(Globals.ThisAddIn.Application);
            _libraryManager = new ShapeLibraryManager(_ppManager);
            _cacheService = new AssetCacheService();
            _assetService = new LocalAssetService(_ppManager);

            DataContext = this;

            ShapeGallery.ItemsSource = FilteredShapes;
            LstFolders.ItemsSource = LibraryFolders;
            LstSharedFolders.ItemsSource = SharedFolders;

            this.Loaded += (s, e) => {
                if (!_isInitialized) RefreshLibraryAsync();
            };
        }

        private List<string> _detectedEmpowerShapeFolders = new List<string>();
        private List<string> _detectedThinkCellShapeFolders = new List<string>();
        private List<string> _detectedEfficientShapeFolders = new List<string>();

        private void DetectThirdPartyTools()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            // empower check - check both shapes folder and general Library
            _detectedEmpowerShapeFolders.Clear();
            string empowerShapes = System.IO.Path.Combine(localAppData, "empower", "shapes");
            string empowerLibrary = System.IO.Path.Combine(localAppData, "empower", "Library");
            string empowerLibraryAppData = System.IO.Path.Combine(appData, "empower", "Library");
            if (System.IO.Directory.Exists(empowerShapes))
                _detectedEmpowerShapeFolders.Add(empowerShapes);
            if (System.IO.Directory.Exists(empowerLibrary))
                _detectedEmpowerShapeFolders.Add(empowerLibrary);
            if (System.IO.Directory.Exists(empowerLibraryAppData))
                _detectedEmpowerShapeFolders.Add(empowerLibraryAppData);

            if (_detectedEmpowerShapeFolders.Count > 0)
            {
                TxtEmpowerStatus.Text = "empower┬« Library found.";
                BtnAddEmpower.IsEnabled = true;
                BtnAddEmpower.Background = new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4));
                BtnAddEmpower.Foreground = new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF));
            }
            else
            {
                TxtEmpowerStatus.Text = "empower┬« not found.";
                BtnAddEmpower.IsEnabled = false;
                BtnAddEmpower.Background = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33));
                BtnAddEmpower.Foreground = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66));
            }

            // think-cell check - check both shapes folder and general Library
            _detectedThinkCellShapeFolders.Clear();
            string thinkCellShapes = System.IO.Path.Combine(appData, "think-cell", "shapes");
            string thinkCellLibrary = System.IO.Path.Combine(localAppData, "thinkcell");
            string thinkCellLibraryAppData = System.IO.Path.Combine(appData, "think-cell");
            if (System.IO.Directory.Exists(thinkCellShapes))
                _detectedThinkCellShapeFolders.Add(thinkCellShapes);
            if (System.IO.Directory.Exists(thinkCellLibrary))
                _detectedThinkCellShapeFolders.Add(thinkCellLibrary);
            if (System.IO.Directory.Exists(thinkCellLibraryAppData))
                _detectedThinkCellShapeFolders.Add(thinkCellLibraryAppData);

            if (_detectedThinkCellShapeFolders.Count > 0)
            {
                TxtThinkCellStatus.Text = "think-cell Library found.";
                BtnAddThinkCell.IsEnabled = true;
                BtnAddThinkCell.Background = new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4));
                BtnAddThinkCell.Foreground = new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF));
            }
            else
            {
                TxtThinkCellStatus.Text = "think-cell not found.";
                BtnAddThinkCell.IsEnabled = false;
                BtnAddThinkCell.Background = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33));
                BtnAddThinkCell.Foreground = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66));
            }

            // Efficient Elements check - check both Shapes subfolder and general Library
            _detectedEfficientShapeFolders.Clear();
            
            string[] eePaths = {
                System.IO.Path.Combine(localAppData, "Efficient Elements", "Efficient Elements for presentations", "Library", "Shapes"),
                System.IO.Path.Combine(localAppData, "Efficient Elements", "Efficient Elements for presentations", "Library"),
                System.IO.Path.Combine(appData, "Efficient Elements", "Efficient Elements for presentations", "Library", "Shapes"),
                System.IO.Path.Combine(appData, "Efficient Elements", "Efficient Elements for presentations", "Library")
            };

            foreach (var path in eePaths)
            {
                if (System.IO.Directory.Exists(path) && !_detectedEfficientShapeFolders.Contains(path))
                    _detectedEfficientShapeFolders.Add(path);
            }

            if (_detectedEfficientShapeFolders.Count > 0)
            {
                TxtEfficientStatus.Text = "Efficient Elements found.";
                BtnAddEfficient.IsEnabled = true;
                BtnAddEfficient.Background = new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4));
                BtnAddEfficient.Foreground = new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF));
            }
            else
            {
                TxtEfficientStatus.Text = "Efficient Elements not found.";
                BtnAddEfficient.IsEnabled = false;
                BtnAddEfficient.Background = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33));
                BtnAddEfficient.Foreground = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66));
            }
        }

        public async void RefreshLibraryAsync()
        {
            if (_isRefreshing) return;
            _isRefreshing = true;
            _isInitialized = true;
            if (_isGenerating) return;
            try
            {
                TxtStatus.Text = "Scanning shape folders...";
                LoadingOverlay.Visibility = Visibility.Visible;
                TxtLoadingDetails.Text = "";
                AllShapes.Clear();
                FilteredShapes.Clear();
                LibraryFolders.Clear();
                FolderTree.Items.Clear();
                _scannedFiles.Clear();
                _selectedFolderPath = null;

                var folders = _libraryManager.GetLibraryFolders();
                foreach (var f in folders) LibraryFolders.Add(f);

                if (folders.Count == 0)
                {
                    TxtStatus.Text = "No library folders configured. Please add a folder in Settings.";
                    LoadingOverlay.Visibility = Visibility.Collapsed;
                    ToggleSettings();
                    return;
                }

                ToggleSettings();

                var files = await _libraryManager.ScanLibrariesAsync();
                _scannedFiles = files;
                TxtStatus.Text = $"Found {files.Count} shape libraries.";

                bool foldersOnly = ChkFoldersOnly.IsChecked == true;
                var treeRoots = _libraryManager.BuildFolderTree(files, foldersOnly);
                FolderTree.Items.Clear();
                foreach (var root in treeRoots)
                    FolderTree.Items.Add(root);

                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string favoritesRoot = Path.Combine(appDataPath, "oPenEfficiency", "favorites");

                // For shape library, we do not eagerly generate COM thumbnails for ALL folders
                // as that takes way too long. We just load cached ones if they exist,
                // and defer generation of uncached ones until a specific file is selected.
                int loadedCache = 0;
                foreach (var file in files)
                {
                    bool isFavorite = file.FilePath.StartsWith(favoritesRoot, StringComparison.OrdinalIgnoreCase);
                    string category = file.FolderName;
                    if (!string.IsNullOrEmpty(file.SubFolder))
                        category = file.SubFolder.Replace(System.IO.Path.DirectorySeparatorChar.ToString(), " / ");

                    if (_libraryManager.IsCacheValid(file.FilePath))
                    {
                        var items = _libraryManager.GetCachedShapeItems(file.FilePath);
                        foreach(var item in items)
                        {
                            var tpl = new ShapeItemDisplay 
                            { 
                                Title = item.Title,
                                Category = category,
                                FilePath = file.FilePath,
                                SlideIndex = item.SlideIndex,
                                OriginalShapeId = item.OriginalShapeId,
                                UniqueId = item.UniqueId,
                                IsFavorite = isFavorite
                            };
                            AllShapes.Add(tpl);
                            FilteredShapes.Add(tpl);
                            loadedCache++;

                            // Lazy Load Thumbnail
                            LoadThumbnailForTemplate(tpl);
                        }
                    }
                }
                
                TxtStatus.Text = $"Ready. {loadedCache} shapes loaded from cache. Uncached files must be clicked to generate.";
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "Error: " + ex.Message;
            }
            finally
            {
                LoadingOverlay.Visibility = Visibility.Collapsed;
                _isRefreshing = false;
            }
        }

        private async void BtnFavoriteSelectionContext_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selection = _ppManager.GetApplication().ActiveWindow.Selection;
                if (selection.Type != PpSelectionType.ppSelectionShapes && selection.Type != PpSelectionType.ppSelectionText)
                {
                    MessageBox.Show("Please select one or more objects in PowerPoint to add them to your favorites.", "Selection Error", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                string suggestedName = "Favorite_" + DateTime.Now.ToString("yyyyMMdd_HHmm");
                var dialog = new SaveToMyAssetsDialog(_assetService, AssetType.Shapes, suggestedName);
                if (dialog.ShowDialog() == true)
                {
                    var result = await _assetService.SaveSelectedShapes(dialog.TargetFolderPath, dialog.FileName, dialog.Tags, dialog.AutoTag);
                    if (result.Success)
                    {
                        MessageBox.Show($"Saved {result.ItemCount} object(s) to favorites.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                        RefreshLibraryAsync();
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
        }

        private async void BtnAddThisToFavorites_Click(object sender, RoutedEventArgs e)
        {
            if (ShapeGallery.SelectedItem is ShapeItemDisplay t)
            {
                try
                {
                    string suggestedName = t.Title;
                    var dialog = new SaveToMyAssetsDialog(_assetService, AssetType.Shapes, suggestedName);
                    if (dialog.ShowDialog() == true)
                    {
                        // For shapes, we currently only support favoriting the entire container presentation
                        // or we'd need a more complex 'Extract and Save' logic.
                        // For now, let's copy the container PPTX.
                        string destPath = Path.Combine(dialog.TargetFolderPath, dialog.FileName + ".pptx");
                        File.Copy(t.FilePath, destPath, true);

                        if (!string.IsNullOrEmpty(dialog.Tags))
                        {
                            _assetService.UpdateAssetDetails(destPath, dialog.FileName, dialog.Tags);
                        }

                        MessageBox.Show("Asset container added to favorites.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                        RefreshLibraryAsync();
                    }
                }
                catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
            }
        }

        private async void LoadThumbnailForTemplate(ShapeItemDisplay template)
        {
            var cached = await Task.Run(() => _libraryManager.LoadCachedThumbnail(template.FilePath, template.UniqueId));
            if (cached != null)
            {
                Dispatcher.Invoke(() => template.ThumbnailSource = cached);
            }
        }

        private void BtnToggleSettings_Checked(object sender, RoutedEventArgs e) { }
        private void BtnToggleSettings_Unchecked(object sender, RoutedEventArgs e) { }

        private void BtnShowAll_Click(object sender, MouseButtonEventArgs e)
        {
            _selectedFolderPath = null;
            ApplyFilters();
        }

        private void UserControl_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine($"[ShapeLibraryTab] PreviewMouseWheel fired - Delta: {e.Delta}, SettingsView.Visibility: {SettingsView.Visibility}");

            // Only handle mouse wheel when Settings view is visible
            if (SettingsView.Visibility != Visibility.Visible)
            {
                System.Diagnostics.Debug.WriteLine("[ShapeLibraryTab] Skipping - SettingsView not visible");
                return;
            }

            if (SettingsScrollViewer != null)
            {
                System.Diagnostics.Debug.WriteLine($"[ShapeLibraryTab] Scrolling - Current Offset: {SettingsScrollViewer.VerticalOffset}, New Offset: {SettingsScrollViewer.VerticalOffset - e.Delta}");
                System.Diagnostics.Debug.WriteLine($"[ShapeLibraryTab] ScrollViewer ExtentHeight: {SettingsScrollViewer.ExtentHeight}, ViewportHeight: {SettingsScrollViewer.ViewportHeight}");
                SettingsScrollViewer.ScrollToVerticalOffset(SettingsScrollViewer.VerticalOffset - e.Delta);
                e.Handled = true;
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[ShapeLibraryTab] SettingsScrollViewer is null!");
            }
        }

        public void ShowSettings()
        {
            SettingsView.Visibility = Visibility.Visible;
            SettingsView.Height = 500; // Force fixed height to enable scrolling
            BrowserView.Visibility = Visibility.Collapsed;
            BrowserToolbar.Visibility = Visibility.Collapsed;

            // Load shared folders
            SharedFolders.Clear();
            foreach (var folder in _cacheService.GetSharedFolders())
            {
                SharedFolders.Add(folder);
            }

            DetectThirdPartyTools();
        }

        private bool _isSettingsOpen = false;

        public void ToggleSettings()
        {
            if (_isSettingsOpen)
            {
                // Slide out
                var anim = new System.Windows.Media.Animation.DoubleAnimation(400, new Duration(TimeSpan.FromMilliseconds(250)))
                {
                    EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseIn }
                };
                anim.Completed += (s, e) => SettingsView.Visibility = Visibility.Collapsed;
                
                if (SettingsView.RenderTransform is System.Windows.Media.TranslateTransform transform)
                {
                    transform.BeginAnimation(System.Windows.Media.TranslateTransform.XProperty, anim);
                }
                
                _isSettingsOpen = false;
                
                // Keep BrowserView visible and interactive
                if (BrowserView != null) { BrowserView.IsHitTestVisible = true; BrowserView.Opacity = 1.0; }
            }
            else
            {
                // Slide in
                DetectThirdPartyTools();
                SettingsView.Visibility = Visibility.Visible;
                
                if (!(SettingsView.RenderTransform is System.Windows.Media.TranslateTransform))
                {
                    SettingsView.RenderTransform = new System.Windows.Media.TranslateTransform(400, 0);
                }
                
                var transform = (System.Windows.Media.TranslateTransform)SettingsView.RenderTransform;
                var anim = new System.Windows.Media.Animation.DoubleAnimation(400, 0, new Duration(TimeSpan.FromMilliseconds(250)))
                {
                    EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut }
                };
                transform.BeginAnimation(System.Windows.Media.TranslateTransform.XProperty, anim);
                
                _isSettingsOpen = true;
                
                // Dim and disable BrowserView
                if (BrowserView != null) { BrowserView.IsHitTestVisible = false; BrowserView.Opacity = 0.4; }
            }
        }

        private void FolderTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is FolderTreeNode node)
            {
                if (node.IsFile)
                {
                    _selectedFolderPath = null;
                    // Filter or Generate
                    if (_libraryManager.IsCacheValid(node.FilePath))
                    {
                        FilteredShapes.Clear();
                        var matches = AllShapes.Where(t => t.FilePath.Equals(node.FilePath, StringComparison.OrdinalIgnoreCase)).ToList();
                        foreach (var m in matches) FilteredShapes.Add(m);
                        TxtStatus.Text = $"Showing {matches.Count} shapes from {node.Name} (Caching...)";
                    }
                    else
                    {
                        // Generate COM cache
                        GenerateShapesForFile(node.FilePath);
                    }
                }
                else
                {
                    _selectedFolderPath = node.FullPath;
                    ApplyFilters();
                    TxtStatus.Text = $"Showing templates from: {node.Name} ({node.TotalFileCount} files) (Caching...)";
                }
            }
        }

        private void GenerateShapesForFile(string filePath)
        {
            if (_isGenerating) return;
            
            _isGenerating = true;
            FilteredShapes.Clear();
            LoadingOverlay.Visibility = Visibility.Visible;
            string fileName = System.IO.Path.GetFileName(filePath);
            TxtStatus.Text = $"Generating shapes for {fileName}...";
            TxtLoadingDetails.Text = "This may take a moment via PowerPoint COM.";

            // Run COM on UI Thread using Dispatcher.BeginInvoke to not freeze everything synchronously,
            // but COM calls themselves will still block the UI thread unfortunately unless we start a new PPT instance. 
            // For now, this is inline with our other library managers.
            Dispatcher.BeginInvoke(new Action(async () =>
            {
                try
                {
                    int count = 0;
                    await _libraryManager.GenerateThumbnailCache(filePath, (item, img) => {
                        
                        var sf = _scannedFiles.FirstOrDefault(f => f.FilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase));
                        string category = sf != null ? sf.FolderName : "Generated";

                        var tpl = new ShapeItemDisplay 
                        { 
                            Title = item.Title,
                            Category = category,
                            FilePath = filePath,
                            SlideIndex = item.SlideIndex,
                            OriginalShapeId = item.OriginalShapeId,
                            UniqueId = item.UniqueId,
                            ThumbnailSource = img
                        };
                        
                        // Remove old versions of this shape if any
                        var old = AllShapes.FirstOrDefault(t => t.FilePath == filePath && t.UniqueId == item.UniqueId);
                        if (old != null) AllShapes.Remove(old);

                        AllShapes.Add(tpl);
                        FilteredShapes.Add(tpl);
                        count++;
                    });
                    
                    TxtStatus.Text = $"Added {count} newly generated shapes from {fileName}.";
                }
                catch(Exception ex) 
                {
                    TxtStatus.Text = $"Error generating: {ex.Message}";
                }
                finally
                {
                    _isGenerating = false;
                    LoadingOverlay.Visibility = Visibility.Collapsed;
                }
            }));
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e) => ApplyFilters();

        private void ApplyFilters()
        {
            if (_isGenerating) return;

            string query = TxtSearch?.Text?.ToLower() ?? "";
            FilteredShapes.Clear();

            foreach (var t in AllShapes)
            {
                if (_selectedFolderPath != null && !t.FilePath.StartsWith(_selectedFolderPath, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!string.IsNullOrWhiteSpace(query))
                {
                    bool matchesText = t.Title.ToLower().Contains(query) ||
                                       (t.Category != null && t.Category.ToLower().Contains(query));
                    if (!matchesText)
                        continue;
                }

                FilteredShapes.Add(t);
            }
        }

                private void ShapeGallery_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int count = ShapeGallery.SelectedItems.Count;
            if (BtnInsertSelection != null) BtnInsertSelection.Visibility = count > 0 ? Visibility.Visible : Visibility.Collapsed;
            if (TxtSelectionCount != null)
            {
                TxtSelectionCount.Visibility = count > 0 ? Visibility.Visible : Visibility.Collapsed;
                TxtSelectionCount.Text = $"{count} shapes selected";
            }
        }

        private void ShapeGallery_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ShapeGallery.SelectedItem is ShapeItemDisplay item)
            {
                InsertShape(item);
            }
        }

        private void BtnInsertShapeContext_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem mi && mi.DataContext is ShapeItemDisplay item)
            {
                InsertShape(item);
            }
            else if (ShapeGallery.SelectedItem is ShapeItemDisplay selected)
            {
                InsertShape(selected);
            }
        }

        private void BtnInsertShape_Click(object sender, RoutedEventArgs e)
        {
            ShapeItemDisplay item = null;
            if (sender is Button btn && btn.Tag is ShapeItemDisplay i1) item = i1;
            else if (sender is MenuItem mi && mi.DataContext is ShapeItemDisplay i2) item = i2;
            else if (sender is FrameworkElement fe && fe.DataContext is ShapeItemDisplay i3) item = i3;

            if (item != null)
            {
                // Check if the clicked template is part of a multi-selection
                var selectedShapes = ShapeGallery.SelectedItems.Cast<ShapeItemDisplay>().ToList();
                if (selectedShapes.Count > 1 && selectedShapes.Contains(item))
                {
                    // User clicked Insert on an item that is part of a multi-selection. Insert all selected.
                    BtnInsertSelected_Click(null, null);
                }
                else
                {
                    InsertShape(item);
                }
            }
        }

        private void BtnInsertSelected_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = ShapeGallery.SelectedItems.Cast<ShapeItemDisplay>().ToList();
            if (selectedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one shape to insert.", "No Selection", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int successCount = 0;
            foreach (var item in selectedItems)
            {
                TxtStatus.Text = $"Inserting '{item.Title}'...";
                bool success = _libraryManager.InsertShape(item.FilePath, item.SlideIndex, item.OriginalShapeId);
                if (success) successCount++;
            }
            
            TxtStatus.Text = $"Inserted {successCount} out of {selectedItems.Count} shape(s).";
        }

        private void InsertShape(ShapeItemDisplay item)
        {
            TxtStatus.Text = $"Inserting '{item.Title}'...";
            bool success = _libraryManager.InsertShape(item.FilePath, item.SlideIndex, item.OriginalShapeId);
            
            if (success)
                TxtStatus.Text = $"Inserted '{item.Title}'.";
            else
            {
                TxtStatus.Text = "Insertion failed.";
                MessageBox.Show("Could not insert the shape. Check if the file is valid and accessible.", "Insertion Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnAddFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Select a folder containing PowerPoint library files (.pptx)";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    _libraryManager.AddLibraryFolder(dialog.SelectedPath);
                    LibraryFolders.Clear();
                    foreach (var f in _libraryManager.GetLibraryFolders()) LibraryFolders.Add(f);
                    RefreshLibraryAsync();
                }
            }
        }

        private async void BtnNewFolder_Click(object sender, RoutedEventArgs e)
        {
            FolderTreeNode selectedNode = FolderTree.SelectedItem as FolderTreeNode;
            if (selectedNode == null)
            {
                MessageBox.Show("Please select a folder first.", "No Folder Selected",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dialog = new CreateFolderDialog();
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    string newPath = _assetService.CreateSubFolder(selectedNode.FullPath, dialog.FolderName);
                    RefreshLibraryAsync();
                    await Task.Delay(500); // Brief delay to allow refresh to complete
                    SelectFolderByPath(newPath);
                    MessageBox.Show($"Folder '{dialog.FolderName}' created successfully.",
                        "Folder Created", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Could not create folder: {ex.Message}",
                        "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async void BtnRefreshCache_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem mi && mi.DataContext is FolderTreeNode node && node.IsFile)
            {
                TxtLoadingDetails.Text = $"Forcing cache refresh for {node.Name}...";
                LoadingOverlay.Visibility = Visibility.Visible;
                
                await _libraryManager.GenerateThumbnailCache(node.FilePath);
                
                RefreshLibraryAsync();
                LoadingOverlay.Visibility = Visibility.Collapsed;
                MessageBox.Show("Cache refreshed successfully.", "Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem mi && mi.DataContext is FolderTreeNode node)
            {
                string path = node.IsFile ? Path.GetDirectoryName(node.FullPath) : node.FullPath;
                if (Directory.Exists(path))
                {
                    System.Diagnostics.Process.Start("explorer.exe", path);
                }
            }
        }

        private void SelectFolderByPath(string path)
        {
            if (FolderTree.Items.Count == 0) return;
            var node = FindFolderNode(FolderTree.Items[0] as FolderTreeNode, path);
            if (node != null)
            {
                SetTreeViewItemSelection(node);
            }
        }

        private void SetTreeViewItemSelection(object dataNode)
        {
            TreeViewItem item = FindTreeViewItem(FolderTree, dataNode);
            if (item != null)
            {
                item.IsSelected = true;
            }
        }

        private TreeViewItem FindTreeViewItem(TreeViewItem root, object dataNode)
        {
            if (root.DataContext == dataNode) return root;

            root.IsExpanded = true;
            foreach (TreeViewItem child in root.Items)
            {
                var found = FindTreeViewItem(child, dataNode);
                if (found != null) return found;
            }

            return null;
        }

        private TreeViewItem FindTreeViewItem(TreeView tree, object dataNode)
        {
            foreach (TreeViewItem item in tree.Items)
            {
                var found = FindTreeViewItem(item, dataNode);
                if (found != null) return found;
            }
            return null;
        }

        private FolderTreeNode FindFolderNode(FolderTreeNode node, string path)
        {
            if (node == null) return null;
            if (node.FullPath == path) return node;
            foreach (var child in node.Children)
            {
                var found = FindFolderNode(child, path);
                if (found != null) return found;
            }
            return null;
        }

        private void BtnRemoveFolder_Click(object sender, RoutedEventArgs e)
        {
            if (LstFolders.SelectedItem is string path)
            {
                _libraryManager.RemoveLibraryFolder(path);
                LibraryFolders.Remove(path);
                RefreshLibraryAsync();
            }
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e) => RefreshLibraryAsync();

        private void ChkFoldersOnly_Changed(object sender, RoutedEventArgs e)
        {
            if (IsLoaded) RefreshLibraryAsync();
        }

        private void BtnAddEmpower_Click(object sender, RoutedEventArgs e)
        {
            int addedCount = 0;
            foreach (var path in _detectedEmpowerShapeFolders)
            {
                if (!LibraryFolders.Contains(path, StringComparer.OrdinalIgnoreCase))
                {
                    _libraryManager.AddLibraryFolder(path);
                    addedCount++;
                }
            }
            if (addedCount > 0)
            {
                LibraryFolders.Clear();
                foreach (var f in _libraryManager.GetLibraryFolders()) LibraryFolders.Add(f);
                RefreshLibraryAsync();
                MessageBox.Show($"Added {addedCount} empower folder(s).", "Added", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("All found empower folders are already in the library.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnAddThinkCell_Click(object sender, RoutedEventArgs e)
        {
            int addedCount = 0;
            foreach (var path in _detectedThinkCellShapeFolders)
            {
                if (!LibraryFolders.Contains(path, StringComparer.OrdinalIgnoreCase))
                {
                    _libraryManager.AddLibraryFolder(path);
                    addedCount++;
                }
            }
            if (addedCount > 0)
            {
                LibraryFolders.Clear();
                foreach (var f in _libraryManager.GetLibraryFolders()) LibraryFolders.Add(f);
                RefreshLibraryAsync();
                MessageBox.Show($"Added {addedCount} think-cell folder(s).", "Added", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("All found think-cell folders are already in the library.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnAddEfficient_Click(object sender, RoutedEventArgs e)
        {
            int addedCount = 0;
            foreach (var path in _detectedEfficientShapeFolders)
            {
                if (!LibraryFolders.Contains(path, StringComparer.OrdinalIgnoreCase))
                {
                    _libraryManager.AddLibraryFolder(path);
                    addedCount++;
                }
            }
            if (addedCount > 0)
            {
                LibraryFolders.Clear();
                foreach (var f in _libraryManager.GetLibraryFolders()) LibraryFolders.Add(f);
                RefreshLibraryAsync();
                MessageBox.Show($"Added {addedCount} Efficient Elements folder(s).", "Added", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private async void BtnGenerateCache_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "Generating thumbnails for all shape libraries can take several minutes depending on the number and size of files.\n\nDo you want to continue?",
                "Generate Thumbnails",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes) return;

            try
            {
                BtnGenerateCache.IsEnabled = false;
                BtnGenerateCache.Content = "Generating...";

                var files = await _libraryManager.ScanLibrariesAsync();
                int processed = 0;
                int failed = 0;

                foreach (var file in files)
                {
                    try
                    {
                        string cacheDir = _libraryManager.GetCacheFolder(file.FilePath);
                        bool cacheValid = _libraryManager.IsCacheValid(file.FilePath);

                        if (!cacheValid)
                        {
                            await _libraryManager.GenerateThumbnailCache(file.FilePath);
                        }
                        processed++;
                    }
                    catch
                    {
                        failed++;
                    }
                }

                MessageBox.Show(
                    $"Thumbnail generation completed.\nProcessed: {processed}\nFailed: {failed}",
                    "Complete",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating thumbnails: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnGenerateCache.IsEnabled = true;
                BtnGenerateCache.Content = "ÔÜí Generate All Thumbnails";
            }
        }

        private void BtnAddSharedFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Select a shared folder (OneDrive, network share) where local caching should be used";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    _cacheService.AddSharedFolder(dialog.SelectedPath);
                    SharedFolders.Add(dialog.SelectedPath);
                    TxtStatus.Text = "Shared folder added. Local cache will be used for files in this folder.";
                }
            }
        }

        private void BtnRemoveSharedFolder_Click(object sender, RoutedEventArgs e)
        {
            if (LstSharedFolders.SelectedItem is string path)
            {
                _cacheService.RemoveSharedFolder(path);
                SharedFolders.Remove(path);
                TxtStatus.Text = "Shared folder removed.";
            }
        }

        private async void BtnSaveToMyAssets_Click(object sender, RoutedEventArgs e)
        {
            if (ShapeGallery.SelectedItem == null)
            {
                MessageBox.Show("Please select a shape to save.", "No Selection",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var selectedShapes = ShapeGallery.SelectedItems.Cast<ShapeItemDisplay>().ToList();
            if (selectedShapes.Count == 0)
            {
                MessageBox.Show("Please select at least one shape to save.", "No Selection",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // For shapes, we need to get the actual selection from PowerPoint
            var dialog = new SaveToMyAssetsDialog(_assetService, AssetType.Shapes, "SelectedShapes");
            if (dialog.ShowDialog() == true)
            {
                TxtStatus.Text = "Saving shape(s) to My Assets...";
                var result = await _assetService.SaveSelectedShapes(dialog.TargetFolderPath, dialog.FileName, dialog.Tags, dialog.AutoTag);

                if (result.Success)
                {
                    TxtStatus.Text = $"Saved {result.ItemCount} shape(s) to My Assets.";
                    MessageBox.Show($"Successfully saved {result.ItemCount} shape(s) to your assets.",
                        "Save Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    TxtStatus.Text = "Failed to save shape(s).";
                    MessageBox.Show($"Failed to save: {result.ErrorMessage}",
                        "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnMyAssets_Click(object sender, MouseButtonEventArgs e)
        {
            // Navigate to My Assets root folder
            string myAssetsPath = _assetService.GetAssetFolderPath(AssetType.Shapes);
            if (Directory.Exists(myAssetsPath))
            {
                _selectedFolderPath = myAssetsPath;
                ApplyFilters();
                TxtStatus.Text = "Viewing My Assets (Shapes)";

                // Load shape files from My Assets
                var myAssetFiles = _assetService.GetShapeFiles(myAssetsPath);
                FilteredShapes.Clear();
                foreach (var file in myAssetFiles)
                {
                    var template = new ShapeItemDisplay
                    {
                        Title = file.Name,
                        FilePath = file.FilePath,
                        Category = "My Assets",
                        ThumbnailSource = null // Will be loaded lazily if needed
                    };
                    FilteredShapes.Add(template);
                }
            }
        }

                private void BtnEditDetails_Click(object sender, RoutedEventArgs e)
        {
            ShapeItemDisplay item = null;
            if (sender is MenuItem mi && mi.DataContext is ShapeItemDisplay i1) item = i1;
            else if (ShapeGallery.SelectedItem is ShapeItemDisplay i2) item = i2;

            if (item == null) return;

            string currentKeywords = _assetService.GetAssetKeywords(item.FilePath);
            var dialog = new oPenEfficiency.UI.Dialogs.EditAssetDetailsDialog(item.Title, currentKeywords, item.FilePath);
            if (dialog.ShowDialog() == true)
            {
                if (_assetService.UpdateAssetDetails(item.FilePath, dialog.AssetName, dialog.Tags))
                {
                    TxtStatus.Text = "Asset updated successfully.";
                    RefreshLibraryAsync();
                }
                else
                {
                    MessageBox.Show("Could not update asset. Ensure it's not open in another application.", "Update Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnDeleteAsset_Click(object sender, RoutedEventArgs e)
        {
            ShapeItemDisplay item = null;
            if (sender is MenuItem mi && mi.DataContext is ShapeItemDisplay i1) item = i1;
            else if (ShapeGallery.SelectedItem is ShapeItemDisplay i2) item = i2;

            if (item == null) return;

            if (MessageBox.Show($"Are you sure you want to permanently delete '{item.Title}'?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                if (_assetService.DeleteAsset(item.FilePath))
                {
                    TxtStatus.Text = "Asset deleted.";
                    RefreshLibraryAsync();
                }
                else
                {
                    MessageBox.Show("Could not delete asset.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }


    // Helper model for data binding shape display properties
    public class ShapeItemDisplay : System.ComponentModel.INotifyPropertyChanged
    {
        public string Title { get; set; }
        public string Category { get; set; }
        public string FilePath { get; set; }
        public int SlideIndex { get; set; }
        public int OriginalShapeId { get; set; }
        public int UniqueId { get; set; }

        private bool _isFavorite;
        public bool IsFavorite
        {
            get => _isFavorite;
            set
            {
                _isFavorite = value;
                PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(IsFavorite)));
            }
        }

        private BitmapImage _thumbnailSource;
        public BitmapImage ThumbnailSource
        {
            get => _thumbnailSource;
            set
            {
                _thumbnailSource = value;
                PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(ThumbnailSource)));
            }
        }

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
    }
}
