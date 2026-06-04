using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class FeatureExplorerWindow : Window, System.ComponentModel.INotifyPropertyChanged
    {
        private List<FeatureViewModel> _allFeatures;
        private string _currentCategory = "All";

        private IEnumerable<FeatureViewModel> _displayedFeatures;
        public IEnumerable<FeatureViewModel> DisplayedFeatures
        {
            get => _displayedFeatures;
            set { _displayedFeatures = value; PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(DisplayedFeatures))); }
        }

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        public FeatureExplorerWindow()
        {
            InitializeComponent();
            DataContext = this;
            LoadFeatures();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void LoadFeatures()
        {
            _allFeatures = new List<FeatureViewModel>();
            
            foreach (var wrapper in FeatureDiscovery.AllFeatures)
            {
                var info = wrapper.DisplayInfo;
                _allFeatures.Add(new FeatureViewModel
                {
                    Id = wrapper.Id,
                    Name = wrapper.Name,
                    Description = info.Description,
                    DetailedHelpText = string.IsNullOrEmpty(info.DetailedHelpText) ? info.Description : info.DetailedHelpText,
                    IconData = info.IconData,
                    ColorHex = info.Color,
                    HelpImagePath = info.HelpImagePath,
                    Category = DetermineCategory(wrapper.Id, wrapper.Name),
                    Keywords = info.Keywords
                });
            }

            foreach (var sf in FeatureLibrary.AllFeatures)
            {
                if (!_allFeatures.Any(f => f.Id == sf.Id))
                {
                    var info = FeatureLibrary.GetFeatureInfo(sf.Id);
                    _allFeatures.Add(new FeatureViewModel
                    {
                        Id = sf.Id,
                        Name = sf.Name,
                        Description = info.Description,
                        DetailedHelpText = string.IsNullOrEmpty(info.DetailedHelpText) ? info.Description : info.DetailedHelpText,
                        IconData = info.IconData,
                        ColorHex = info.Color,
                        HelpImagePath = info.HelpImagePath,
                        Category = DetermineCategory(sf.Id, sf.Name),
                        Keywords = info.Keywords
                    });
                }
            }

            _allFeatures = _allFeatures.OrderBy(f => f.Name).ToList();
            FilterFeatures();
            
            if (_allFeatures.Count > 0)
            {
                FeatureList.SelectedIndex = 0;
            }
        }

        private string DetermineCategory(string id, string name)
        {
            string lowerId = id.ToLowerInvariant();
            string lowerName = name.ToLowerInvariant();
            
            if (lowerId.Contains("align") || lowerName.Contains("align") || lowerId.Contains("arrange") || lowerId.Contains("distribute") || lowerId.Contains("dock") || lowerId.Contains("stretch") || lowerId.Contains("swap"))
                return "Align";
                
            if (lowerId.Contains("format") || lowerName.Contains("format") || lowerId.Contains("style") || lowerId.Contains("theme") || lowerId.Contains("size") || lowerId.Contains("font"))
                return "Format";
                
            if (lowerId.Contains("color") || lowerName.Contains("chart") || lowerId.Contains("visual") || lowerId.Contains("shape") || lowerId.Contains("rating") || lowerId.Contains("map"))
                return "Visual";
                
            if (lowerId.Contains("data") || lowerId.Contains("excel") || lowerId.Contains("text") || lowerId.Contains("word") || lowerId.Contains("table") || lowerId.Contains("spell") || lowerId.Contains("agenda"))
                return "Data";

            return "Utility";
        }

        private void FilterFeatures()
        {
            if (_allFeatures == null) return;

            var query = SearchBox?.Text?.ToLowerInvariant() ?? "";

            DisplayedFeatures = _allFeatures.Where(f => 
                (_currentCategory == "All" || f.Category == _currentCategory) &&
                (string.IsNullOrEmpty(query) || 
                 f.Name.ToLowerInvariant().Contains(query) || 
                 (f.Description != null && f.Description.ToLowerInvariant().Contains(query)) ||
                 (f.Keywords != null && f.Keywords.ToLowerInvariant().Contains(query)))
            ).ToList();
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            FilterFeatures();
        }

        private void Category_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton rb)
            {
                _currentCategory = rb.Tag?.ToString() ?? "All";
                FilterFeatures();
            }
        }

        private void FeatureList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FeatureList.SelectedItem is FeatureViewModel feature)
            {
                DetailsPanel.Visibility = Visibility.Visible;
                DetailName.Text = feature.Name;
                DetailDescription.Text = feature.Description;
                DetailDetailedHelp.Text = feature.DetailedHelpText;

                try
                {
                    DetailIcon.Data = feature.PathData;
                    DetailIcon.Fill = feature.PathFill;
                    DetailIcon.Stroke = feature.PathStroke;
                    DetailIcon.StrokeThickness = feature.PathStrokeThickness;
                    DetailIcon.StrokeLineJoin = feature.PathStrokeLineJoin;
                    DetailIcon.StrokeStartLineCap = feature.PathStrokeStartLineCap;
                    DetailIcon.StrokeEndLineCap = feature.PathStrokeEndLineCap;
                }
                catch { DetailIcon.Data = null; }

                var shortcuts = ShortcutManager.GetShortcuts();
                if (shortcuts.TryGetValue(feature.Id, out string shortcutStr) && !string.IsNullOrEmpty(shortcutStr))
                {
                    ShortcutPanel.Visibility = Visibility.Visible;
                    DetailShortcut.Text = shortcutStr;
                }
                else
                {
                    ShortcutPanel.Visibility = Visibility.Collapsed;
                }

                if (!string.IsNullOrEmpty(feature.HelpImagePath))
                {
                    try
                    {
                        DetailImage.Source = new BitmapImage(new Uri(feature.HelpImagePath, UriKind.RelativeOrAbsolute));
                        DetailImageBorder.Visibility = Visibility.Visible;
                    }
                    catch
                    {
                        DetailImageBorder.Visibility = Visibility.Collapsed;
                    }
                }
                else
                {
                    DetailImageBorder.Visibility = Visibility.Collapsed;
                }
            }
            else
            {
                DetailsPanel.Visibility = Visibility.Collapsed;
            }
        }

        public class FeatureViewModel
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public string Description { get; set; }
            public string DetailedHelpText { get; set; }
            public string IconData { get; set; }
            public string ColorHex { get; set; }
            public string HelpImagePath { get; set; }
            public string Category { get; set; }
            public string Keywords { get; set; }

            private Geometry _pathData;
            private Brush _pathFill = Brushes.Transparent;
            private Brush _pathStroke = Brushes.Transparent;
            private double _pathStrokeThickness = 0;
            private bool _parsed = false;

            public Geometry PathData { get { ParseIcon(); return _pathData; } }
            public Brush PathFill { get { ParseIcon(); return _pathFill; } }
            public Brush PathStroke { get { ParseIcon(); return _pathStroke; } }
            public double PathStrokeThickness { get { ParseIcon(); return _pathStrokeThickness; } }
            public PenLineJoin PathStrokeLineJoin => PenLineJoin.Round;
            public PenLineCap PathStrokeStartLineCap => PenLineCap.Round;
            public PenLineCap PathStrokeEndLineCap => PenLineCap.Round;

            private void ParseIcon()
            {
                if (_parsed) return;
                _parsed = true;

                try
                {
                    var baseColor = (Brush)new BrushConverter().ConvertFrom(ColorHex ?? "#888888");

                    string codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
                    UriBuilder uri = new UriBuilder(codeBase);
                    string assemblyDir = System.IO.Path.GetDirectoryName(Uri.UnescapeDataString(uri.Path));
                    string svgFilePath = System.IO.Path.Combine(assemblyDir, "UI", "Assets", "Icons", Id + ".svg");

                    if (!System.IO.File.Exists(svgFilePath))
                    {
                        string devPath = System.IO.Path.GetFullPath(System.IO.Path.Combine(assemblyDir, "..", "..", "UI", "Assets", "Icons", Id + ".svg"));
                        if (System.IO.File.   Exists(devPath)) svgFilePath = devPath;
                    }

                    bool isFilled;
                    var svgGeom = oPenEfficiency.Utils.SvgParser.ParseLucideSvg(svgFilePath, out isFilled);
                    if (svgGeom != null && svgGeom.Children.Count > 0)
                    {
                        _pathData = svgGeom;
                        if (isFilled)
                        {
                            _pathFill = baseColor;
                        }
                        else
                        {
                            _pathStroke = baseColor;
                            _pathStrokeThickness = 2;
                        }
                    }
                    else
                    {
                        _pathData = Geometry.Parse(IconData);
                        _pathFill = baseColor;
                    }
                }
                catch
                {
                    try { _pathData = Geometry.Parse(IconData); } catch { }
                    try { _pathFill = (Brush)new BrushConverter().ConvertFrom(ColorHex ?? "#888888"); } catch { _pathFill = Brushes.Gray; }
                }
            }
        }
    }
}