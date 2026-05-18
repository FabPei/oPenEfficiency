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
using oPenEfficiency.UI;
using oPenEfficiency.UI.Dialogs;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI.Tabs
{
    public partial class SlideLibraryTab : UserControl
    {
        private readonly SlideLibraryManager _libraryManager;
        private readonly PowerPointManager _ppManager;
        private readonly AssetCacheService _cacheService;
        private readonly LocalAssetService _assetService;

        public ObservableCollection<SlideTemplate> AllTemplates { get; set; } = new ObservableCollection<SlideTemplate>();
        public ObservableCollection<SlideTemplate> FilteredTemplates { get; set; } = new ObservableCollection<SlideTemplate>();

        public ObservableCollection<string> LibraryFolders { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<string> SharedFolders { get; set; } = new ObservableCollection<string>();

        private string _selectedFolderPath = null;
        private bool _isRefreshing = false;
        private bool _isInitialized = false;

        private List<string> _detectedEmpowerFolders = new List<string>();
        private List<string> _detectedThinkCellFolders = new List<string>();
        private List<string> _detectedEfficientFolders = new List<string>();

        public SlideLibraryTab()
        {
            InitializeComponent();

            _ppManager = new PowerPointManager(Globals.ThisAddIn.Application);
            _libraryManager = new SlideLibraryManager(_ppManager);
            _cacheService = new AssetCacheService();
            _assetService = new LocalAssetService(_ppManager);

            DataContext = this;

            SlideGallery.ItemsSource = FilteredTemplates;
            LstFolders.ItemsSource = LibraryFolders;
            LstSharedFolders.ItemsSource = SharedFolders;

            this.Loaded += (s, e) => {
                if (!_isInitialized) RefreshLibraryAsync();
            };
        }

        public async void RefreshLibraryAsync()
        {
            if (_isRefreshing) return;
            _isRefreshing = true;
            _isInitialized = true;

            try
            {
                TxtStatus.Text = "Scanning library folders...";
                LoadingOverlay.Visibility = Visibility.Visible;
                
                var folders = _libraryManager.GetLibraryFolders();
                LibraryFolders.Clear();
                foreach (var f in folders) LibraryFolders.Add(f);

                if (folders.Count == 0)
                {
                    TxtStatus.Text = "No library folders configured.";
                    LoadingOverlay.Visibility = Visibility.Collapsed;
                    ToggleSettings();
                    _isRefreshing = false;
                    return;
                }

                var files = await _libraryManager.ScanLibrariesAsync();
                
                bool foldersOnly = ChkFoldersOnly.IsChecked == true;
                var treeRoots = _libraryManager.BuildFolderTree(files, foldersOnly);
                FolderTree.Items.Clear();
                foreach (var root in treeRoots)
                    FolderTree.Items.Add(root);

                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string favoritesRoot = Path.Combine(appDataPath, "oPenEfficiency", "favorites");

                var newTemplates = new List<SlideTemplate>();
                foreach (var file in files)
                {
                    bool isFavorite = file.FilePath.StartsWith(favoritesRoot, StringComparison.OrdinalIgnoreCase);
                    var slides = _libraryManager.GetSlidesFromFile(file.FilePath);
                    foreach (var s in slides)
                    {
                        if (newTemplates.Any(t => t.FilePath.Equals(s.FilePath, StringComparison.OrdinalIgnoreCase) && t.SlideIndex == s.SlideIndex))
                            continue;

                        s.IsFavorite = isFavorite;
                        newTemplates.Add(s);
                    }
                }

                // Batch update UI
                AllTemplates.Clear();
                FilteredTemplates.Clear();
                foreach (var t in newTemplates)
                {
                    AllTemplates.Add(t);
                    FilteredTemplates.Add(t);
                }

                TxtStatus.Text = $"Ready. {AllTemplates.Count} slides loaded.";
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
                if (selection.Type != PpSelectionType.ppSelectionSlides)
                {
                    MessageBox.Show("Please select one or more slides in PowerPoint to add them to your favorites.", "Selection Error", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                string suggestedName = "Favorite_" + DateTime.Now.ToString("yyyyMMdd_HHmm");
                var dialog = new SaveToMyAssetsDialog(_assetService, AssetType.Slides, suggestedName);
                if (dialog.ShowDialog() == true)
                {
                    var result = await _assetService.SaveSelectedSlides(dialog.TargetFolderPath, dialog.FileName, dialog.Tags, dialog.AutoTag);
                    if (result.Success)
                    {
                        MessageBox.Show($"Saved {result.ItemCount} slide(s) to favorites.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
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
                    var dialog = new SaveToMyAssetsDialog(_assetService, AssetType.Slides, suggestedName);
                    if (dialog.ShowDialog() == true)
                    {
                        // Extract only this slide instead of copying the whole file
                        var result = await _assetService.ExtractAndSaveSlide(t.FilePath, t.SlideIndex, dialog.TargetFolderPath, dialog.FileName, dialog.Tags);
                        
                        if (result.Success)
                        {
                            MessageBox.Show("Slide added to favorites.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                            RefreshLibraryAsync();
                        }
                        else
                        {
                            MessageBox.Show("Could not add to favorites: " + result.ErrorMessage, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
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

        private void DetectThirdPartyTools()
        {
            _detectedEmpowerFolders.Clear();
            _detectedThinkCellFolders.Clear();
            _detectedEfficientFolders.Clear();

            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

            // empower detection
            string[] empPaths = {
                System.IO.Path.Combine(localAppData, "empower", "Library"),
                System.IO.Path.Combine(appData, "empower", "Library"),
                System.IO.Path.Combine(programFiles, "empower", "Library"),
                System.IO.Path.Combine(programFilesX86, "empower", "Library")
            };
            foreach (var p in empPaths) if (System.IO.Directory.Exists(p)) _detectedEmpowerFolders.Add(p);

            // think-cell detection
            string[] tcPaths = {
                System.IO.Path.Combine(localAppData, "thinkcell"),
                System.IO.Path.Combine(appData, "think-cell"),
                System.IO.Path.Combine(programFiles, "think-cell"),
                System.IO.Path.Combine(programFilesX86, "think-cell")
            };
            foreach (var p in tcPaths) if (System.IO.Directory.Exists(p)) _detectedThinkCellFolders.Add(p);

            // Efficient Elements
            string[] eePaths = {
                Path.Combine(localAppData, "Efficient Elements", "Efficient Elements for presentations", "Library"),
                Path.Combine(appData, "Efficient Elements", "Efficient Elements for presentations", "Library")
            };

            foreach (var path in eePaths)
            {
                if (Directory.Exists(path) && !_detectedEfficientFolders.Contains(path))
                    _detectedEfficientFolders.Add(path);
            }

            // Update UI status
            if (_detectedEmpowerFolders.Count > 0)
            {
                TxtEmpowerStatus.Text = "empower┬« Library found.";
                BtnAddEmpower.IsEnabled = true;
            }
            else
            {
                TxtEmpowerStatus.Text = "empower┬« not found.";
                BtnAddEmpower.IsEnabled = false;
            }

            if (_detectedThinkCellFolders.Count > 0)
            {
                TxtThinkCellStatus.Text = "think-cell Library found.";
                BtnAddThinkCell.IsEnabled = true;
            }
            else
            {
                TxtThinkCellStatus.Text = "think-cell not found.";
                BtnAddThinkCell.IsEnabled = false;
            }

            if (_detectedEfficientFolders.Count > 0)
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
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            } catch { }
            return paths.Where(p => !string.IsNullOrEmpty(p) && System.IO.Directory.Exists(p)).Distinct().ToList();
        }

        private void BtnOpenSettings_Click(object sender, RoutedEventArgs e) => ToggleSettings();
        private void BtnRefresh_Click(object sender, RoutedEventArgs e) => RefreshLibraryAsync();
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

        private void BtnRemoveFolder_Click(object sender, RoutedEventArgs e)
        {
            if (LstFolders.SelectedItem is string path)
            {
                _libraryManager.RemoveLibraryFolder(path);
                LibraryFolders.Remove(path);
                RefreshLibraryAsync();
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

        private void BtnAddEmpower_Click(object sender, RoutedEventArgs e)
        {
            foreach (var folder in _detectedEmpowerFolders) _libraryManager.AddLibraryFolder(folder);
            RefreshLibraryAsync();
        }

        private void BtnAddThinkCell_Click(object sender, RoutedEventArgs e)
        {
            foreach (var folder in _detectedThinkCellFolders) _libraryManager.AddLibraryFolder(folder);
            RefreshLibraryAsync();
        }

        private void BtnAddEfficient_Click(object sender, RoutedEventArgs e)
        {
            foreach (var folder in _detectedEfficientFolders) _libraryManager.AddLibraryFolder(folder);
            RefreshLibraryAsync();
        }

        private async void BtnGenerateCache_Click(object sender, RoutedEventArgs e)
        {
            TxtStatus.Text = "Generating thumbnails (this may take a while)...";
            LoadingOverlay.Visibility = Visibility.Visible;
            await _libraryManager.GenerateAllThumbnails();
            LoadingOverlay.Visibility = Visibility.Collapsed;
            RefreshLibraryAsync();
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e) => ApplyFilters();
        private void ChkFoldersOnly_Changed(object sender, RoutedEventArgs e) => RefreshLibraryAsync();

        private void FolderTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is FolderTreeNode node)
            {
                _selectedFolderPath = node.FullPath;
                ApplyFilters();
            }
        }

        private void ApplyFilters()
        {
            string query = TxtSearch?.Text?.ToLower() ?? "";
            FilteredTemplates.Clear();

            foreach (var t in AllTemplates)
            {
                if (_selectedFolderPath != null && !t.FilePath.StartsWith(_selectedFolderPath, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!string.IsNullOrWhiteSpace(query))
                {
                    bool matchesText = t.Title.ToLower().Contains(query) ||
                                       (t.Keywords != null && t.Keywords.ToLower().Contains(query)) ||
                                       (t.Category != null && t.Category.ToLower().Contains(query));
                    if (!matchesText)
                        continue;
                }

                FilteredTemplates.Add(t);
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
                TxtSelectionCount.Text = $"{count} slides selected";
            }
        }

        private void BtnNewFolder_Click(object sender, RoutedEventArgs e)
        {
            if (FolderTree.SelectedItem is FolderTreeNode node)
            {
                var dialog = new CreateFolderDialog();
                if (dialog.ShowDialog() == true)
                {
                    _assetService.CreateSubFolder(node.FullPath, dialog.FolderName);
                    RefreshLibraryAsync();
                }
            }
        }

        private async void BtnRefreshCache_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem mi && mi.DataContext is FolderTreeNode node && node.IsFile)
            {
                TxtStatus.Text = $"Forcing cache refresh for {node.Name}...";
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

        private void SlideGallery_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (SlideGallery.SelectedItem is SlideTemplate t)
                _libraryManager.InsertSlide(t.FilePath, t.SlideIndex, ChkMatchTheme.IsChecked == true);
        }

        private void BtnInsertSlideContext_Click(object sender, RoutedEventArgs e)
        {
            if (SlideGallery.SelectedItem is SlideTemplate t)
                _libraryManager.InsertSlide(t.FilePath, t.SlideIndex, ChkMatchTheme.IsChecked == true);
        }

        private void BtnInsertSlide_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button b && b.Tag is SlideTemplate t)
                _libraryManager.InsertSlide(t.FilePath, t.SlideIndex, ChkMatchTheme.IsChecked == true);
        }

        private void BtnInsertSelected_Click(object sender, RoutedEventArgs e)
        {
            foreach (SlideTemplate t in SlideGallery.SelectedItems)
                _libraryManager.InsertSlide(t.FilePath, t.SlideIndex, ChkMatchTheme.IsChecked == true);
        }

        private async void BtnSaveToMyAssets_Click(object sender, RoutedEventArgs e)
        {
            if (SlideGallery.SelectedItem == null) return;

            var selectedSlides = SlideGallery.SelectedItems.Cast<SlideTemplate>().ToList();
            if (selectedSlides.Count == 0) return;

            string suggestedName = selectedSlides.FirstOrDefault()?.Title ?? "Slide";
            var dialog = new SaveToMyAssetsDialog(_assetService, AssetType.Slides, suggestedName);
            if (dialog.ShowDialog() == true)
            {
                var result = await _assetService.SaveSelectedSlides(dialog.TargetFolderPath, dialog.FileName, dialog.Tags, dialog.AutoTag);
                if (result.Success) RefreshLibraryAsync();
            }
        }

        private void BtnMyAssets_Click(object sender, MouseButtonEventArgs e)
        {
            string myAssetsPath = _assetService.GetAssetFolderPath(AssetType.Slides);
            if (Directory.Exists(myAssetsPath))
            {
                _selectedFolderPath = myAssetsPath;
                ApplyFilters();
            }
        }

        private void BtnEditDetails_Click(object sender, RoutedEventArgs e)
        {
            if (SlideGallery.SelectedItem is SlideTemplate t)
            {
                string currentKeywords = _assetService.GetAssetKeywords(t.FilePath);
                var dialog = new EditAssetDetailsDialog(t.Title, currentKeywords, t.FilePath);
                if (dialog.ShowDialog() == true)
                {
                    _assetService.UpdateAssetDetails(t.FilePath, dialog.AssetName, dialog.Tags);
                    RefreshLibraryAsync();
                }
            }
        }

        private void BtnDeleteAsset_Click(object sender, RoutedEventArgs e)
        {
            if (SlideGallery.SelectedItem is SlideTemplate t)
            {
                if (MessageBox.Show("Delete slide?", "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    _assetService.DeleteAsset(t.FilePath);
                    RefreshLibraryAsync();
                }
            }
        }
        private void UserControl_PreviewMouseWheel(object sender, MouseWheelEventArgs e) { if (SettingsView.Visibility == Visibility.Visible) { SettingsScrollViewer.ScrollToVerticalOffset(SettingsScrollViewer.VerticalOffset - e.Delta); e.Handled = true; } }
    }
}
