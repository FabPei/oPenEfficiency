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
using oPenEfficiency.Features;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.UI.Dialogs;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI.Tabs
{
    public partial class ImageLibraryTab : UserControl
    {
        private bool _isRefreshing = false;
        private readonly ImageLibraryManager _libraryManager;
        private readonly PowerPointManager _ppManager;
        private readonly AssetCacheService _cacheService;
        private readonly LocalAssetService _assetService;

        public ObservableCollection<SlideTemplate> AllImages { get; set; } = new ObservableCollection<SlideTemplate>();
        public ObservableCollection<SlideTemplate> FilteredImages { get; set; } = new ObservableCollection<SlideTemplate>();

        public ObservableCollection<string> LibraryFolders { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<string> SharedFolders { get; set; } = new ObservableCollection<string>();

        private string _selectedFolderPath = null;
        private bool _isInitialized = false;

        public ImageLibraryTab()
        {
            InitializeComponent();

            _ppManager = new PowerPointManager(Globals.ThisAddIn.Application);
            _libraryManager = new ImageLibraryManager(_ppManager);
            _cacheService = new AssetCacheService();
            _assetService = new LocalAssetService(_ppManager);

            DataContext = this;

            SlideGallery.ItemsSource = FilteredImages;
            LstFolders.ItemsSource = LibraryFolders;
            LstSharedFolders.ItemsSource = SharedFolders;

            // Initialize Settings
            ChkIgnoreSmall.IsChecked = oPenEfficiency.Properties.Settings.Default.IgnoreSmallImages;
            SldThreshold.Value = oPenEfficiency.Properties.Settings.Default.SmallImageThresholdKB;

            ChkIgnoreSmall.Checked += (s, e) => { oPenEfficiency.Properties.Settings.Default.IgnoreSmallImages = true; oPenEfficiency.Properties.Settings.Default.Save(); RefreshLibraryAsync(); };
            ChkIgnoreSmall.Unchecked += (s, e) => { oPenEfficiency.Properties.Settings.Default.IgnoreSmallImages = false; oPenEfficiency.Properties.Settings.Default.Save(); RefreshLibraryAsync(); };
            SldThreshold.ValueChanged += (s, e) => { 
                if (_isInitialized) {
                    oPenEfficiency.Properties.Settings.Default.SmallImageThresholdKB = (int)SldThreshold.Value; 
                    oPenEfficiency.Properties.Settings.Default.Save(); 
                    // We don't refresh immediately on slider move to avoid lag, 
                    // the user can click Refresh or it will happen on next scan.
                }
            };

            this.Loaded += (s, e) => {
                if (!_isInitialized) RefreshLibraryAsync();
            };
        }

        private List<string> _detectedEmpowerImageFolders = new List<string>();
        private List<string> _detectedThinkCellImageFolders = new List<string>();
        private List<string> _detectedEfficientImageFolders = new List<string>();

        private void DetectThirdPartyTools()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            _detectedEmpowerImageFolders.Clear();
            _detectedThinkCellImageFolders.Clear();
            _detectedEfficientImageFolders.Clear();

            // empower check
            string empowerShapes = Path.Combine(localAppData, "empower", "shapes");
            if (Directory.Exists(empowerShapes)) _detectedEmpowerImageFolders.Add(empowerShapes);
            
            if (_detectedEmpowerImageFolders.Count > 0)
            {
                TxtEmpowerStatus.Text = "empower┬« folders found.";
                BtnAddEmpower.IsEnabled = true;
            }
            else
            {
                TxtEmpowerStatus.Text = "empower┬« not found.";
                BtnAddEmpower.IsEnabled = false;
            }

            // think-cell check
            string thinkCellShapes = Path.Combine(appData, "think-cell", "shapes");
            if (Directory.Exists(thinkCellShapes)) _detectedThinkCellImageFolders.Add(thinkCellShapes);

            if (_detectedThinkCellImageFolders.Count > 0)
            {
                TxtThinkCellStatus.Text = "think-cell folders found.";
                BtnAddThinkCell.IsEnabled = true;
            }
            else
            {
                TxtThinkCellStatus.Text = "think-cell not found.";
                BtnAddThinkCell.IsEnabled = false;
            }

            // Efficient Elements check
            string[] eePaths = {
                Path.Combine(localAppData, "Efficient Elements", "Efficient Elements for presentations", "Library"),
                Path.Combine(appData, "Efficient Elements", "Efficient Elements for presentations", "Library")
            };

            foreach (var path in eePaths)
            {
                if (Directory.Exists(path) && !_detectedEfficientImageFolders.Contains(path))
                    _detectedEfficientImageFolders.Add(path);
            }

            if (_detectedEfficientImageFolders.Count > 0)
            {
                TxtEfficientStatus.Text = "Efficient Elements found.";
                BtnAddEfficient.IsEnabled = true;
            }
            else
            {
                TxtEfficientStatus.Text = "Efficient Elements not found.";
                BtnAddEfficient.IsEnabled = false;
            }
        }

        public async void RefreshLibraryAsync()
        {
            if (_isRefreshing) return;
            _isRefreshing = true;
            try
            {
                TxtStatus.Text = "Scanning image folders...";
                LoadingOverlay.Visibility = Visibility.Visible;
                AllImages.Clear();
                FilteredImages.Clear();
                LibraryFolders.Clear();
                FolderTree.Items.Clear();
                _selectedFolderPath = null;

                var folders = _libraryManager.GetLibraryFolders();
                foreach (var f in folders) LibraryFolders.Add(f);

                if (folders.Count == 0)
                {
                    TxtStatus.Text = "No library folders configured.";
                    LoadingOverlay.Visibility = Visibility.Collapsed;
                    return;
                }

                var files = await _libraryManager.ScanLibrariesAsync();
                
                bool foldersOnly = ChkFoldersOnly.IsChecked == true;
                bool ignoreSmall = oPenEfficiency.Properties.Settings.Default.IgnoreSmallImages;
                long thresholdBytes = oPenEfficiency.Properties.Settings.Default.SmallImageThresholdKB * 1024;

                var treeRoots = _libraryManager.BuildFolderTree(files, foldersOnly);
                FolderTree.Items.Clear();
                foreach (var root in treeRoots)
                    FolderTree.Items.Add(root);

                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string favoritesRoot = Path.Combine(appDataPath, "oPenEfficiency", "favorites");

                var newImages = new List<SlideTemplate>();
                foreach (var file in files)
                {
                    if (newImages.Any(at => at.FilePath.Equals(file.FilePath, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    // Small image filtering
                    if (ignoreSmall)
                    {
                        string ext = Path.GetExtension(file.FilePath).ToLower();
                        if (ext == ".png" || ext == ".jpg" || ext == ".jpeg")
                        {
                            try
                            {
                                long size = new FileInfo(file.FilePath).Length;
                                if (size < thresholdBytes) continue;
                            }
                            catch { }
                        }
                    }

                    bool isFavorite = file.FilePath.StartsWith(favoritesRoot, StringComparison.OrdinalIgnoreCase);
                    var t = new SlideTemplate
                    {
                        Title = file.FileName,
                        FilePath = file.FilePath,
                        Category = file.FolderName,
                        Keywords = file.FileName,
                        IsFavorite = isFavorite
                    };
                    newImages.Add(t);
                }

                // Batch update UI
                AllImages.Clear();
                FilteredImages.Clear();
                foreach (var i in newImages)
                {
                    AllImages.Add(i);
                    FilteredImages.Add(i);
                    LoadThumbnailForTemplate(i);
                }
                
                TxtStatus.Text = $"Ready. {AllImages.Count} images loaded.";
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
                    MessageBox.Show("Please select one or more objects/images in PowerPoint to add them to your favorites.", "Selection Error", MessageBoxButton.OK, MessageBoxImage.Information);
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
            if (SlideGallery.SelectedItem is SlideTemplate t)
            {
                try
                {
                    string suggestedName = t.Title;
                    var dialog = new SaveToMyAssetsDialog(_assetService, AssetType.Images, suggestedName);
                    if (dialog.ShowDialog() == true)
                    {
                        string ext = Path.GetExtension(t.FilePath);
                        string destPath = Path.Combine(dialog.TargetFolderPath, dialog.FileName + ext);
                        File.Copy(t.FilePath, destPath, true);

                        MessageBox.Show("Image added to favorites.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                        RefreshLibraryAsync();
                    }
                }
                catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
            }
        }

        private async void LoadThumbnailForTemplate(SlideTemplate template)
        {
            var thumb = await _libraryManager.GetThumbnailAsync(template.FilePath);
            if (thumb != null)
            {
                Dispatcher.Invoke(() => template.ThumbnailSource = thumb);
            }
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

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e) => ApplyFilters();

        private void ApplyFilters()
        {
            string query = TxtSearch?.Text?.ToLower() ?? "";
            FilteredImages.Clear();

            foreach (var t in AllImages)
            {
                if (_selectedFolderPath != null && !t.FilePath.StartsWith(_selectedFolderPath, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!string.IsNullOrWhiteSpace(query))
                {
                    bool matchesText = t.Title.ToLower().Contains(query) ||
                                       (t.Keywords != null && t.Keywords.ToLower().Contains(query));
                    if (!matchesText)
                        continue;
                }

                FilteredImages.Add(t);
            }
        }

        private void FolderTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is FolderTreeNode node)
            {
                if (node.IsFile)
                {
                    _selectedFolderPath = null;
                    FilteredImages.Clear();
                    var match = AllImages.FirstOrDefault(t => t.FilePath.Equals(node.FilePath, StringComparison.OrdinalIgnoreCase));
                    if (match != null) FilteredImages.Add(match);
                }
                else
                {
                    _selectedFolderPath = node.FullPath;
                    ApplyFilters();
                }
            }
        }

        private void BtnShowAll_Click(object sender, MouseButtonEventArgs e)
        {
            _selectedFolderPath = null;
            ApplyFilters();
        }

                private void SlideGallery_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int count = SlideGallery.SelectedItems.Count;
            if (BtnInsertSelection != null) BtnInsertSelection.Visibility = count > 0 ? Visibility.Visible : Visibility.Collapsed;
            if (TxtSelectionCount != null)
            {
                TxtSelectionCount.Visibility = count > 0 ? Visibility.Visible : Visibility.Collapsed;
                TxtSelectionCount.Text = $"{count} images selected";
            }
        }

        private void SlideGallery_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (SlideGallery.SelectedItem is SlideTemplate template)
            {
                InsertImage(template);
            }
        }

        private void BtnInsertImageContext_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem mi && mi.DataContext is SlideTemplate template)
            {
                InsertImage(template);
            }
            else if (SlideGallery.SelectedItem is SlideTemplate selected)
            {
                InsertImage(selected);
            }
        }

        private void BtnInsertImage_Click(object sender, RoutedEventArgs e)
        {
            SlideTemplate template = null;
            if (sender is Button btn && btn.Tag is SlideTemplate t1) template = t1;
            else if (sender is MenuItem mi && mi.DataContext is SlideTemplate t2) template = t2;
            else if (sender is FrameworkElement fe && fe.DataContext is SlideTemplate t3) template = t3;

            if (template != null)
            {
                var selectedImages = SlideGallery.SelectedItems.Cast<SlideTemplate>().ToList();
                if (selectedImages.Count > 1 && selectedImages.Contains(template))
                {
                    BtnInsertSelected_Click(null, null);
                }
                else
                {
                    InsertImage(template);
                }
            }
        }

        private void BtnInsertSelected_Click(object sender, RoutedEventArgs e)
        {
            var selectedImages = SlideGallery.SelectedItems.Cast<SlideTemplate>().ToList();
            if (selectedImages.Count == 0)
            {
                MessageBox.Show("Please select at least one image to insert.", "No Selection", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int successCount = 0;
            bool maintainRatio = ChkMaintainRatio.IsChecked == true;

            foreach (var template in selectedImages)
            {
                TxtStatus.Text = $"Inserting '{template.Title}'...";
                bool success = _libraryManager.InsertLibraryImage(template.FilePath, maintainRatio);
                if (success) successCount++;
            }

            TxtStatus.Text = $"Inserted {successCount} out of {selectedImages.Count} image(s).";
        }

        private void InsertImage(SlideTemplate template)
        {
            bool maintainRatio = ChkMaintainRatio.IsChecked == true;
            TxtStatus.Text = $"Inserting '{template.Title}'...";
            
            bool success = _libraryManager.InsertLibraryImage(template.FilePath, maintainRatio);
            
            if (success)
                TxtStatus.Text = $"Successfully inserted '{template.Title}'.";
            else
            {
                TxtStatus.Text = "Insertion failed.";
                MessageBox.Show("Could not insert the image. Check if the file is valid and accessible.", "Insertion Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnAddFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Select a folder containing images (.png, .jpg, .svg)";
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
            if (SlideGallery.SelectedItem == null)
            {
                MessageBox.Show("Please select an image to save.", "No Selection",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var selectedImages = SlideGallery.SelectedItems.Cast<SlideTemplate>().ToList();
            if (selectedImages.Count == 0)
            {
                MessageBox.Show("Please select at least one image to save.", "No Selection",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Get source image paths
            var sourcePaths = selectedImages.Select(i => i.FilePath).ToList();
            string suggestedName = selectedImages.FirstOrDefault()?.Title ?? "Image";

            var dialog = new SaveToMyAssetsDialog(_assetService, AssetType.Images, suggestedName);
            if (dialog.ShowDialog() == true)
            {
                TxtStatus.Text = "Saving image(s) to My Assets...";
                var result = await _assetService.SaveSelectedImages(sourcePaths, dialog.TargetFolderPath, dialog.FileName);

                if (result.Success)
                {
                    TxtStatus.Text = $"Saved {result.ItemCount} image(s) to My Assets.";
                    MessageBox.Show($"Successfully saved {result.ItemCount} image(s) to your assets.",
                        "Save Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    TxtStatus.Text = "Failed to save image(s).";
                    MessageBox.Show($"Failed to save: {result.ErrorMessage}",
                        "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnMyAssets_Click(object sender, MouseButtonEventArgs e)
        {
            // Navigate to My Assets root folder
            string myAssetsPath = _assetService.GetAssetFolderPath(AssetType.Images);
            if (Directory.Exists(myAssetsPath))
            {
                _selectedFolderPath = myAssetsPath;
                ApplyFilters();
                TxtStatus.Text = "Viewing My Assets (Images)";

                // Load image files from My Assets
                var myAssetFiles = _assetService.GetImageFiles(myAssetsPath);
                FilteredImages.Clear();
                foreach (var file in myAssetFiles)
                {
                    var template = new SlideTemplate
                    {
                        Title = file.Name,
                        FilePath = file.FilePath,
                        Category = "My Assets",
                        Keywords = file.Name
                    };
                    FilteredImages.Add(template);
                }
            }
        }

                private void BtnEditDetails_Click(object sender, RoutedEventArgs e)
        {
            SlideTemplate template = null;
            if (sender is MenuItem mi && mi.DataContext is SlideTemplate t1) template = t1;
            else if (SlideGallery.SelectedItem is SlideTemplate t2) template = t2;

            if (template == null) return;

            string currentKeywords = _assetService.GetAssetKeywords(template.FilePath);
            var dialog = new EditAssetDetailsDialog(template.Title, currentKeywords, template.FilePath);
            if (dialog.ShowDialog() == true)
            {
                if (_assetService.UpdateAssetDetails(template.FilePath, dialog.AssetName, dialog.Tags))
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
            SlideTemplate template = null;
            if (sender is MenuItem mi && mi.DataContext is SlideTemplate t1) template = t1;
            else if (SlideGallery.SelectedItem is SlideTemplate t2) template = t2;

            if (template == null) return;

            if (MessageBox.Show($"Are you sure you want to permanently delete '{template.Title}'?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                if (_assetService.DeleteAsset(template.FilePath))
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

        private void UserControl_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Implementation similar to SlideLibraryTab
            if (sender is DependencyObject obj)
            {
                var scrollViewer = FindVisualChild<ScrollViewer>(obj);
                if (scrollViewer != null)
                {
                    scrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset - e.Delta);
                    e.Handled = true;
                }
            }
        }

        private T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T t) return t;
                var found = FindVisualChild<T>(child);
                if (found != null) return found;
            }
            return null;
        }

        public void BtnAddEmpower_Click(object sender, RoutedEventArgs e)
        {
            foreach (var folder in _detectedEmpowerImageFolders)
            {
                _libraryManager.AddLibraryFolder(folder);
            }
            RefreshLibraryAsync();
        }

        public void BtnAddThinkCell_Click(object sender, RoutedEventArgs e)
        {
            foreach (var folder in _detectedThinkCellImageFolders)
            {
                _libraryManager.AddLibraryFolder(folder);
            }
            RefreshLibraryAsync();
        }

        public void BtnAddEfficient_Click(object sender, RoutedEventArgs e)
        {
            foreach (var folder in _detectedEfficientImageFolders)
            {
                _libraryManager.AddLibraryFolder(folder);
            }
            RefreshLibraryAsync();
        }

        public async void BtnGenerateCache_Click(object sender, RoutedEventArgs e)
        {
            TxtStatus.Text = "Generating thumbnails...";
            // Images don't really have a complex cache generation like PPTX, 
            // but we can just refresh to ensure everything is loaded.
            RefreshLibraryAsync();
            await Task.Yield();
        }
            private List<string> GetRegistryInstallPaths(string softwareName)
        {
            var paths = new List<string>();
            try
            {
                var registryViews = new[] { Microsoft.Win32.RegistryView.Registry64, Microsoft.Win32.RegistryView.Registry32 };
                foreach (var view in registryViews)
                {
                    using (var baseKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, view))
                    {
                        var uninstallKey = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
                        if (uninstallKey != null)
                        {
                            foreach (var subKeyName in uninstallKey.GetSubKeyNames())
                            {
                                using (var subKey = uninstallKey.OpenSubKey(subKeyName))
                                {
                                    if (subKey != null)
                                    {
                                        var displayName = subKey.GetValue("DisplayName") as string;
                                        var installLocation = subKey.GetValue("InstallLocation") as string;
                                        if (!string.IsNullOrEmpty(displayName) && displayName.IndexOf(softwareName, StringComparison.OrdinalIgnoreCase) >= 0 && !string.IsNullOrEmpty(installLocation))
                                        {
                                            paths.Add(installLocation);
                                            var libraryPath = System.IO.Path.Combine(installLocation, "Library");
                                            if (System.IO.Directory.Exists(libraryPath)) paths.Add(libraryPath);
                                            var templatePath = System.IO.Path.Combine(installLocation, "Templates");
                                            if (System.IO.Directory.Exists(templatePath)) paths.Add(templatePath);
                                        }
                                    }
                                }
                            }
                        }
                        var softwareKey = baseKey.OpenSubKey($@"SOFTWARE\{softwareName}");
                        if (softwareKey != null)
                        {
                            var installPath = softwareKey.GetValue("InstallPath") as string ?? softwareKey.GetValue("Path") as string ?? softwareKey.GetValue(null) as string;
                            if (!string.IsNullOrEmpty(installPath) && System.IO.Directory.Exists(installPath)) paths.Add(installPath);
                        }
                    }
                }
            } catch { }
            return paths.Where(p => !string.IsNullOrEmpty(p) && System.IO.Directory.Exists(p)).Distinct().ToList();
        }
}
}