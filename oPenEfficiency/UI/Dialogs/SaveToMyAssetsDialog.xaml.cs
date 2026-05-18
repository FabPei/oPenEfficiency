using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI
{
    public partial class SaveToMyAssetsDialog : Window
    {
        private readonly LocalAssetService _assetService;
        private readonly AssetType _assetType;
        private AssetFolder _selectedFolder;

        public string FileName => TxtFileName.Text;
        public string Tags => TxtTags.Text;
        public bool AutoTag => ChkAutoTag.IsChecked == true;
        public string TargetFolderPath => _selectedFolder?.FullPath;

        public SaveToMyAssetsDialog(LocalAssetService assetService, AssetType assetType, string suggestedName = null)
        {
            InitializeComponent();
            _assetService = assetService;
            _assetType = assetType;

            // Set default filename
            if (!string.IsNullOrEmpty(suggestedName))
                TxtFileName.Text = suggestedName;

            // Load folder tree
            LoadFolderTree();

            // Auto-select root folder for this asset type
            string defaultPath = _assetService.GetAssetFolderPath(_assetType);
            if (!string.IsNullOrEmpty(defaultPath))
            {
                SelectFolderByPath(defaultPath);
            }
        }

        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                this.DragMove();
        }

        private void LoadFolderTree()
        {
            var folders = _assetService.GetAssetFolders(_assetType);
            var rootItems = new ObservableCollection<AssetFolderTreeNode>();

            // Build tree from flat list
            foreach (var folder in folders.Where(f => f.IsRoot))
            {
                var rootNode = new AssetFolderTreeNode
                {
                    Folder = folder,
                    SubFolders = new ObservableCollection<AssetFolderTreeNode>()
                };

                AddSubFolders(rootNode, folders);
                rootItems.Add(rootNode);
            }

            FolderTree.ItemsSource = rootItems;
        }

        private void AddSubFolders(AssetFolderTreeNode parentNode, System.Collections.Generic.List<AssetFolder> allFolders)
        {
            var subFolders = allFolders.Where(f => f.ParentPath == parentNode.Folder.FullPath && !f.IsRoot);
            foreach (var sub in subFolders.OrderBy(f => f.Name))
            {
                var childNode = new AssetFolderTreeNode
                {
                    Folder = sub,
                    SubFolders = new ObservableCollection<AssetFolderTreeNode>()
                };
                parentNode.SubFolders.Add(childNode);
                AddSubFolders(childNode, allFolders);
            }
        }

        private void FolderTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is AssetFolderTreeNode node)
            {
                _selectedFolder = node.Folder;
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private async void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtFileName.Text))
            {
                MessageBox.Show("Please enter a name for the asset.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_selectedFolder == null)
            {
                MessageBox.Show("Please select a folder.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
            Close();
        }

        private async void BtnNewFolder_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedFolder == null)
            {
                MessageBox.Show("Please select a parent folder.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var inputDialog = new CreateFolderDialog();
            if (inputDialog.ShowDialog() == true)
            {
                try
                {
                    string newPath = _assetService.CreateSubFolder(_selectedFolder.FullPath, inputDialog.FolderName);

                    // Refresh tree
                    LoadFolderTree();

                    // Select the new folder
                    SelectFolderByPath(newPath);

                    MessageBox.Show($"Folder '{inputDialog.FolderName}' created successfully.",
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
            if (FolderTree.ItemsSource == null) return;

            var rootItems = FolderTree.ItemsSource as ObservableCollection<AssetFolderTreeNode>;
            if (rootItems == null || rootItems.Count == 0) return;

            foreach (var root in rootItems)
            {
                var node = FindFolderNode(root, path);
                if (node != null)
                {
                    node.IsSelected = true;
                    // Ensure parents are expanded
                    ExpandPath(root, node);
                    return;
                }
            }
        }

        private void ExpandPath(AssetFolderTreeNode root, AssetFolderTreeNode target)
        {
            // Simple approach: if target is child of root or is root, expand root
            if (root == target)
            {
                root.IsExpanded = true;
                return;
            }

            foreach (var child in root.SubFolders)
            {
                if (IsDescendant(child, target))
                {
                    root.IsExpanded = true;
                    ExpandPath(child, target);
                    return;
                }
            }
        }

        private bool IsDescendant(AssetFolderTreeNode potentialParent, AssetFolderTreeNode target)
        {
            if (potentialParent == target) return true;
            foreach (var child in potentialParent.SubFolders)
            {
                if (IsDescendant(child, target)) return true;
            }
            return false;
        }

        private AssetFolderTreeNode FindFolderNode(AssetFolderTreeNode node, string path)
        {
            if (node == null) return null;
            if (node.Folder.FullPath.Equals(path, StringComparison.OrdinalIgnoreCase)) return node;

            foreach (var child in node.SubFolders)
            {
                var found = FindFolderNode(child, path);
                if (found != null) return found;
            }

            return null;
        }
    }

    public class AssetFolderTreeNode : INotifyPropertyChanged
    {
        public AssetFolder Folder { get; set; }
        public ObservableCollection<AssetFolderTreeNode> SubFolders { get; set; }

        public string Name => Folder?.Name ?? "";
        public string FullPath => Folder?.FullPath ?? "";

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        private bool _isExpanded;
        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                if (_isExpanded != value)
                {
                    _isExpanded = value;
                    OnPropertyChanged(nameof(IsExpanded));
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
