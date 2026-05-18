using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using oPenEfficiency.Models;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class MapWizardWindow : Window
    {
        public MapWizardConfig Config { get; private set; }
        private PowerPoint.Shape _mapGroup;

        public MapWizardWindow(PowerPoint.Shape mapGroup, MapWizardConfig existingConfig)
        {
            InitializeComponent();
            _mapGroup = mapGroup;

            if (existingConfig == null)
            {
                Config = new MapWizardConfig();
            }
            else
            {
                Config = existingConfig;
                LoadConfigIntoUI();
            }
            
            UpdateSwatches();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void LoadConfigIntoUI()
        {
            if (Config.Mode == MapColorMode.Continuous)
                CmbMode.SelectedIndex = 0;
            else
                CmbMode.SelectedIndex = 1;

            TxtLowColor.Text = Config.LowColorHex;
            TxtHighColor.Text = Config.HighColorHex;

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

            var availableShapeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (_mapGroup.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
            {
                foreach (PowerPoint.Shape shp in _mapGroup.GroupItems)
                {
                    availableShapeNames.Add(shp.Name);
                    if (!string.IsNullOrEmpty(shp.AlternativeText))
                    {
                        availableShapeNames.Add(shp.AlternativeText);
                    }
                }
            }

            foreach (var line in lines)
            {
                var parts = line.Split(new[] { '\t' }, StringSplitOptions.None);
                if (parts.Length >= 2)
                {
                    string region = parts[0].Trim();
                    string valStr = parts[1].Trim();
                    
                    double numVal = 0;
                    double.TryParse(valStr, out numVal);

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
            RefreshDataGrid(); 
            Config.LowColorHex = TxtLowColor.Text;
            Config.HighColorHex = TxtHighColor.Text;
            
            this.DialogResult = true;
            this.Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
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
            } catch { }
        }
    }
}