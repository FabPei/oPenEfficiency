using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Diagnostics;
using oPenEfficiency.Features;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.UI.Dialogs;

namespace oPenEfficiency.UI.Tabs
{
    public partial class FavoritesTab : UserControl
    {
        private readonly SlideLibraryManager _slideManager;
        private readonly ShapeLibraryManager _shapeManager;
        private readonly ImageLibraryManager _imageManager;
        private readonly LocalAssetService _assetService;
        private readonly PowerPointManager _ppManager;

        public ObservableCollection<SlideTemplate> FavoriteSlides { get; set; } = new ObservableCollection<SlideTemplate>();
        public ObservableCollection<ShapeItemDisplay> FavoriteShapes { get; set; } = new ObservableCollection<ShapeItemDisplay>();
        public ObservableCollection<SlideTemplate> FavoriteImages { get; set; } = new ObservableCollection<SlideTemplate>();

        private List<SlideTemplate> _allSlides = new List<SlideTemplate>();
        private List<ShapeItemDisplay> _allShapes = new List<ShapeItemDisplay>();
        private List<SlideTemplate> _allImages = new List<SlideTemplate>();

        private string _selectedFolderPath = null;
        private bool _isRefreshing = false;
        private bool _isInitialized = false;

        // Drag & Drop State
        private Point _startPoint;
        private object _draggedItem;
        private FolderTreeNode _draggedNode;

        public FavoritesTab()
        {
            InitializeComponent();

            _ppManager = new PowerPointManager(Globals.ThisAddIn.Application);
            _slideManager = new SlideLibraryManager(_ppManager);
            _shapeManager = new ShapeLibraryManager(_ppManager);
            _imageManager = new ImageLibraryManager(_ppManager);
            _assetService = new LocalAssetService(_ppManager);

            DataContext = this;

            SlideGallery.ItemsSource = FavoriteSlides;
            ShapeGallery.ItemsSource = FavoriteShapes;
            ImageGallery.ItemsSource = FavoriteImages;

            this.Loaded += (s, e) => {
                if (!_isInitialized) RefreshFavoritesAsync();
            };
        }

        public async void RefreshFavoritesAsync()
        {
            if (_isRefreshing) return;
            _isRefreshing = true;
            _isInitialized = true;

            try
            {
                LoadingOverlay.Visibility = Visibility.Visible;
                TxtStatus.Text = "Loading favorites...";

                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string favoritesRoot = Path.Combine(appDataPath, "oPenEfficiency", "favorites");

                var newSlides = new List<SlideTemplate>();
                var newShapes = new List<ShapeItemDisplay>();
                var newImages = new List<SlideTemplate>();

                FolderTree.Items.Clear();

                // Build tree from physical directories to include empty folders
                string slidesPath = Path.Combine(favoritesRoot, "slides");
                if (!Directory.Exists(slidesPath)) Directory.CreateDirectory(slidesPath);
                FolderTree.Items.Add(BuildPhysicalTree(slidesPath, "Slides"));

                string shapesPath = Path.Combine(favoritesRoot, "shapes");
                if (!Directory.Exists(shapesPath)) Directory.CreateDirectory(shapesPath);
                FolderTree.Items.Add(BuildPhysicalTree(shapesPath, "Objects"));

                string imagesPath = Path.Combine(favoritesRoot, "images");
                if (!Directory.Exists(imagesPath)) Directory.CreateDirectory(imagesPath);
                FolderTree.Items.Add(BuildPhysicalTree(imagesPath, "Pictures"));

                // Load all items for filtering
                var slideFiles = oPenEfficiency.Utils.CommonUtils.SafeGetFiles(slidesPath, "*.pptx");
                foreach (var file in slideFiles)
                {
                    string fullPath = Path.GetFullPath(file);
                    var slides = _slideManager.GetSlidesFromFile(fullPath);
                    foreach (var s in slides) { s.IsFavorite = true; newSlides.Add(s); }
                }

                var shapeFilesRaw = oPenEfficiency.Utils.CommonUtils.SafeGetFiles(shapesPath, "*.pptx");
                foreach (var file in shapeFilesRaw)
                {
                    string fullPath = Path.GetFullPath(file);
                    if (_shapeManager.IsCacheValid(fullPath))
                    {
                        var items = _shapeManager.GetCachedShapeItems(fullPath);
                        foreach (var item in items)
                        {
                            newShapes.Add(new ShapeItemDisplay { 
                                Title = item.Title, FilePath = fullPath, SlideIndex = item.SlideIndex, 
                                OriginalShapeId = item.OriginalShapeId, UniqueId = item.UniqueId, IsFavorite = true 
                            });
                        }
                    }
                }

                var imageFilesRaw = oPenEfficiency.Utils.CommonUtils.SafeGetFiles(imagesPath, "*.*")
                    .Where(f => new[] { ".png", ".jpg", ".jpeg", ".svg" }.Contains(Path.GetExtension(f).ToLower())).ToList();
                foreach (var file in imageFilesRaw)
                {
                    string fullPath = Path.GetFullPath(file);
                    newImages.Add(new SlideTemplate { Title = Path.GetFileNameWithoutExtension(fullPath), FilePath = fullPath, IsFavorite = true });
                }

                _allSlides = newSlides;
                _allShapes = newShapes;
                _allImages = newImages;

                ApplyFilters();
                TxtStatus.Text = $"Ready. {_allSlides.Count} slides, {_allShapes.Count} objects, {_allImages.Count} images.";
            }
            catch (Exception ex) { TxtStatus.Text = "Error: " + ex.Message; }
            finally { LoadingOverlay.Visibility = Visibility.Collapsed; _isRefreshing = false; }
        }

        private FolderTreeNode BuildPhysicalTree(string path, string displayName)
        {
            var node = new FolderTreeNode { Name = displayName, FullPath = Path.GetFullPath(path) };
            try
            {
                foreach (var dir in Directory.GetDirectories(path))
                {
                    node.Children.Add(BuildPhysicalTree(dir, Path.GetFileName(dir)));
                }
                foreach (var file in Directory.GetFiles(path, "*.pptx"))
                {
                    node.Children.Add(new FolderTreeNode { Name = Path.GetFileNameWithoutExtension(file), FullPath = Path.GetFullPath(file), FilePath = Path.GetFullPath(file), IsFile = true, IsCached = true });
                }
                // Also add other image types for the Pictures folder
                if (displayName == "Pictures" || path.Contains("images"))
                {
                    foreach (var file in Directory.GetFiles(path).Where(f => new[] { ".png", ".jpg", ".jpeg", ".svg" }.Contains(Path.GetExtension(f).ToLower())))
                    {
                        if (file.EndsWith(".pptx")) continue;
                        node.Children.Add(new FolderTreeNode { Name = Path.GetFileNameWithoutExtension(file), FullPath = Path.GetFullPath(file), FilePath = Path.GetFullPath(file), IsFile = true, IsCached = true });
                    }
                }
            } catch { }
            return node;
        }

        private void ApplyFilters()
        {
            FavoriteSlides.Clear();
            FavoriteShapes.Clear();
            FavoriteImages.Clear();

            foreach (var s in _allSlides) if (string.IsNullOrEmpty(_selectedFolderPath) || s.FilePath.StartsWith(_selectedFolderPath, StringComparison.OrdinalIgnoreCase)) { FavoriteSlides.Add(s); if (s.ThumbnailSource == null) LoadSlideThumbnail(s); }
            foreach (var s in _allShapes) if (string.IsNullOrEmpty(_selectedFolderPath) || s.FilePath.StartsWith(_selectedFolderPath, StringComparison.OrdinalIgnoreCase)) { FavoriteShapes.Add(s); if (s.ThumbnailSource == null) LoadShapeThumbnail(s); }
            foreach (var i in _allImages) if (string.IsNullOrEmpty(_selectedFolderPath) || i.FilePath.StartsWith(_selectedFolderPath, StringComparison.OrdinalIgnoreCase)) { FavoriteImages.Add(i); if (i.ThumbnailSource == null) LoadImageThumbnail(i); }

            TitleSlides.Visibility = FavoriteSlides.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            TitleShapes.Visibility = FavoriteShapes.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            TitleImages.Visibility = FavoriteImages.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            SlideGallery.Visibility = FavoriteSlides.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            ShapeGallery.Visibility = FavoriteShapes.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            ImageGallery.Visibility = FavoriteImages.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void FolderTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is FolderTreeNode node) { _selectedFolderPath = node.FullPath; ApplyFilters(); }
        }

        private void BtnShowAll_Click(object sender, MouseButtonEventArgs e) { _selectedFolderPath = null; ApplyFilters(); }

        private void BtnNewFolder_Click(object sender, RoutedEventArgs e)
        {
            var selectedNode = FolderTree.SelectedItem as FolderTreeNode;
            if (selectedNode == null || selectedNode.IsFile) { MessageBox.Show("Please select a folder first."); return; }

            var dialog = new CreateFolderDialog();
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    string newPath = Path.Combine(selectedNode.FullPath, dialog.FolderName);
                    if (!Directory.Exists(newPath)) { Directory.CreateDirectory(newPath); RefreshFavoritesAsync(); }
                }
                catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
            }
        }

        // --- Drag & Drop ---
        private void FolderTree_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e) { _startPoint = e.GetPosition(null); _draggedNode = FindNodeAtPoint(e.GetPosition(FolderTree)); }
        private void FolderTree_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && _draggedNode != null)
            {
                if (Math.Abs(e.GetPosition(null).X - _startPoint.X) > SystemParameters.MinimumHorizontalDragDistance || Math.Abs(e.GetPosition(null).Y - _startPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
                    DragDrop.DoDragDrop(FolderTree, _draggedNode, DragDropEffects.Move);
            }
        }

        private void Gallery_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Do not initiate drag if clicking an interactive element like a Button
            var originalSource = e.OriginalSource as DependencyObject;
            while (originalSource != null)
            {
                if (originalSource is System.Windows.Controls.Button || originalSource is System.Windows.Controls.TextBox)
                {
                    _draggedItem = null;
                    return;
                }
                originalSource = VisualTreeHelper.GetParent(originalSource);
            }

            _startPoint = e.GetPosition(null);
            var lb = sender as ListBox;
            var item = lb?.InputHitTest(e.GetPosition(lb)) as DependencyObject;
            while (item != null && !(item is ListBoxItem)) item = VisualTreeHelper.GetParent(item);
            _draggedItem = (item as ListBoxItem)?.Content;
        }

        private void Gallery_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && _draggedItem != null)
            {
                if (Math.Abs(e.GetPosition(null).X - _startPoint.X) > SystemParameters.MinimumHorizontalDragDistance || Math.Abs(e.GetPosition(null).Y - _startPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
                    DragDrop.DoDragDrop((DependencyObject)sender, _draggedItem, DragDropEffects.Move);
            }
        }

        private void FolderTree_DragOver(object sender, DragEventArgs e) { var targetNode = FindNodeAtPoint(e.GetPosition(FolderTree)); e.Effects = (targetNode != null && !targetNode.IsFile) ? DragDropEffects.Move : DragDropEffects.None; e.Handled = true; }

        private void FolderTree_Drop(object sender, DragEventArgs e)
        {
            var targetNode = FindNodeAtPoint(e.GetPosition(FolderTree));
            if (targetNode == null || targetNode.IsFile) return;

            string targetFolder = targetNode.FullPath;
            if (e.Data.GetData(typeof(FolderTreeNode)) is FolderTreeNode sourceNode)
            {
                string sourcePath = sourceNode.IsFile ? sourceNode.FilePath : sourceNode.FullPath;
                if (!string.IsNullOrEmpty(sourcePath)) MoveAsset(sourcePath, targetFolder);
            }
            else if (e.Data.GetData(typeof(SlideTemplate)) is SlideTemplate slide) MoveAsset(slide.FilePath, targetFolder);
            else if (e.Data.GetData(typeof(ShapeItemDisplay)) is ShapeItemDisplay shape) MoveAsset(shape.FilePath, targetFolder);
        }

        private void MoveAsset(string sourcePath, string targetFolder)
        {
            try
            {
                string destPath = Path.Combine(targetFolder, Path.GetFileName(sourcePath));
                if (sourcePath.Equals(destPath, StringComparison.OrdinalIgnoreCase)) return;
                if (File.Exists(destPath) || Directory.Exists(destPath)) { if (MessageBox.Show("Overwrite?", "Conflict", MessageBoxButton.YesNo) == MessageBoxResult.Yes) { if (File.Exists(destPath)) File.Delete(destPath); else Directory.Delete(destPath, true); } else return; }
                if (File.Exists(sourcePath)) File.Move(sourcePath, destPath);
                else if (Directory.Exists(sourcePath)) Directory.Move(sourcePath, destPath);
                RefreshFavoritesAsync();
            }
            catch (Exception ex) { MessageBox.Show("Move failed: " + ex.Message); }
        }

        private FolderTreeNode FindNodeAtPoint(Point p)
        {
            var hit = FolderTree.InputHitTest(p) as DependencyObject;
            while (hit != null && !(hit is TreeViewItem)) hit = VisualTreeHelper.GetParent(hit);
            return (hit as TreeViewItem)?.Header as FolderTreeNode;
        }

        private async void LoadSlideThumbnail(SlideTemplate item) { var thumb = await _slideManager.GetThumbnailAsync(item.FilePath); if (thumb != null) Dispatcher.Invoke(() => item.ThumbnailSource = thumb); }
        private async void LoadShapeThumbnail(ShapeItemDisplay item) { var thumb = await Task.Run(() => _shapeManager.LoadCachedThumbnail(item.FilePath, item.UniqueId)); if (thumb != null) Dispatcher.Invoke(() => item.ThumbnailSource = thumb); }
        private async void LoadImageThumbnail(SlideTemplate item) { var thumb = await _imageManager.GetThumbnailAsync(item.FilePath); if (thumb != null) Dispatcher.Invoke(() => item.ThumbnailSource = thumb); }
        private void BtnInsertSlide_Click(object sender, RoutedEventArgs e) { if (sender is Button b && b.Tag is SlideTemplate t) _slideManager.InsertSlide(t.FilePath, t.SlideIndex, true); }
        private void BtnInsertShape_Click(object sender, RoutedEventArgs e) { if (sender is Button b && b.Tag is ShapeItemDisplay t) _shapeManager.InsertShape(t.FilePath, t.SlideIndex, t.OriginalShapeId); }
        private void BtnInsertImage_Click(object sender, RoutedEventArgs e) { if (sender is Button b && b.Tag is SlideTemplate t) _imageManager.InsertLibraryImage(t.FilePath, true); }
        private void BtnInsertContext_Click(object sender, RoutedEventArgs e)
        {
            if (SlideGallery.SelectedItem is SlideTemplate st) _slideManager.InsertSlide(st.FilePath, st.SlideIndex, true);
            else if (ShapeGallery.SelectedItem is ShapeItemDisplay shd) _shapeManager.InsertShape(shd.FilePath, shd.SlideIndex, shd.OriginalShapeId);
            else if (ImageGallery.SelectedItem is SlideTemplate it) _imageManager.InsertLibraryImage(it.FilePath, true);
        }
        private void BtnRemoveFavorite_Click(object sender, RoutedEventArgs e)
        {
            object item = SlideGallery.SelectedItem ?? ShapeGallery.SelectedItem ?? ImageGallery.SelectedItem;
            if (item == null) return;
            string path = (item is SlideTemplate st) ? st.FilePath : (item is ShapeItemDisplay shd) ? shd.FilePath : "";
            if (!string.IsNullOrEmpty(path) && MessageBox.Show("Remove from favorites?", "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                try { _assetService.DeleteAsset(path); RefreshFavoritesAsync(); } catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
            }
        }
    }
}
