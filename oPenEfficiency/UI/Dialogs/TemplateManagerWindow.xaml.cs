using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI.Dialogs
{
    public class TemplateViewModel : TemplateItem
    {
        public bool IsCustom { get; set; }
    }

    public partial class TemplateManagerWindow : Window
    {
        private readonly PowerPointManager _powerPointManager;
        private readonly TemplateService _templateService;
        private ObservableCollection<TemplateViewModel> _allTemplates = new ObservableCollection<TemplateViewModel>();
        private ObservableCollection<string> _watchedDirectories = new ObservableCollection<string>();

        public TemplateManagerWindow(PowerPointManager powerPointManager)
        {
            _powerPointManager = powerPointManager;
            _templateService = new TemplateService();
            InitializeComponent();
            LoadData();
        }

        private void LoadData()
        {
            RefreshTemplateList();
            
            _watchedDirectories.Clear();
            var config = _templateService.GetConfig();
            if (config.WatchedDirectories != null)
            {
                foreach (var dir in config.WatchedDirectories)
                    _watchedDirectories.Add(dir);
            }
            
            DirectoriesList.ItemsSource = _watchedDirectories;
            CbHideDefault.IsChecked = config.HideDefaultLocations;
        }

        private void RefreshTemplateList()
        {
            var config = _templateService.GetConfig();
            var templates = _templateService.GetAllTemplates();
            
            _allTemplates.Clear();
            foreach (var t in templates)
            {
                bool isCustom = config.CustomTemplates != null && config.CustomTemplates.Any(ct => ct.Path.Equals(t.Path, StringComparison.OrdinalIgnoreCase));
                _allTemplates.Add(new TemplateViewModel 
                { 
                    Name = t.Name, 
                    Path = t.Path, 
                    IsCustom = isCustom 
                });
            }
            
            FilterTemplates();
        }

        private void FilterTemplates()
        {
            if (SearchBox == null) return;

            string search = SearchBox.Text.ToLower();
            if (string.IsNullOrWhiteSpace(search))
            {
                TemplatesList.ItemsSource = _allTemplates;
            }
            else
            {
                TemplatesList.ItemsSource = _allTemplates.Where(t => t.Name.ToLower().Contains(search) || t.Path.ToLower().Contains(search)).ToList();
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            FilterTemplates();
        }

        private void TemplatesList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TemplatesList.SelectedItem is TemplateViewModel template)
            {
                NoPreviewPanel.Visibility = Visibility.Collapsed;
                TemplatePreviewImage.Visibility = Visibility.Visible;
                DetailsPanel.Visibility = Visibility.Visible;
                
                TxtSelectedName.Text = template.Name;
                TxtSelectedPath.Text = template.Path;

                // Load Thumbnail
                byte[] data = _templateService.GetThumbnail(template.Path);
                if (data != null)
                {
                    var image = new BitmapImage();
                    using (var ms = new MemoryStream(data))
                    {
                        image.BeginInit();
                        image.CacheOption = BitmapCacheOption.OnLoad;
                        image.StreamSource = ms;
                        image.EndInit();
                    }
                    TemplatePreviewImage.Source = image;
                }
                else
                {
                    TemplatePreviewImage.Source = null; // Maybe a generic placeholder icon later
                }
            }
            else
            {
                NoPreviewPanel.Visibility = Visibility.Visible;
                TemplatePreviewImage.Visibility = Visibility.Collapsed;
                DetailsPanel.Visibility = Visibility.Collapsed;
            }
        }

        private void OpenTemplate_Click(object sender, RoutedEventArgs e)
        {
            TemplateViewModel template = null;
            if (sender is Button btn && btn.DataContext is TemplateViewModel t)
                template = t;
            else if (TemplatesList.SelectedItem is TemplateViewModel sel)
                template = sel;

            if (template != null)
            {
                _templateService.OpenTemplate(_powerPointManager, template.Path);
                this.Close();
            }
        }

        private void HideTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is TemplateViewModel template)
            {
                _templateService.AddToBlacklist(template.Path);
                RefreshTemplateList();
            }
        }

        private void DeleteTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is TemplateViewModel template)
            {
                var result = MessageBox.Show($"Are you sure you want to permanently delete '{template.Name}' from your hard drive?\n\nThis action cannot be undone.", 
                    "Confirm Permanent Deletion", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                
                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        if (File.Exists(template.Path))
                        {
                            File.Delete(template.Path);
                            // Also remove from custom links if it was there
                            _templateService.RemoveCustomTemplate(template.Path);
                            RefreshTemplateList();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Failed to delete file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void RestoreHidden_Click(object sender, RoutedEventArgs e)
        {
            _templateService.ClearBlacklist();
            RefreshTemplateList();
            MessageBox.Show("All hidden templates have been restored.", "Restored", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void CbHideDefault_Changed(object sender, RoutedEventArgs e)
        {
            if (_templateService != null && CbHideDefault != null)
            {
                _templateService.SetHideDefaultLocations(CbHideDefault.IsChecked == true);
                RefreshTemplateList();
            }
        }

        private void AddDirectory_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    _templateService.AddWatchedDirectory(dialog.SelectedPath);
                    LoadData();
                }
            }
        }

        private void RemoveDirectory_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is string dir)
            {
                _templateService.RemoveWatchedDirectory(dir);
                LoadData();
            }
        }

        private void AddTemplate_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.OpenFileDialog())
            {
                dialog.Filter = "PowerPoint Templates (*.potx;*.pot)|*.potx;*.pot|PowerPoint Presentations (*.pptx;*.pptm)|*.pptx;*.pptm|All Files (*.*)|*.*";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    _templateService.AddCustomTemplate(dialog.FileName);
                    RefreshTemplateList();
                }
            }
        }
    }
}
