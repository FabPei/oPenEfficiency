using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using oPenEfficiency.Features;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public partial class ProgressSeriesToolbar : Window
    {
        private PowerPointManager _manager;
        private PowerPoint.Shape _targetShape;
        private ProgressSeriesFeature.SeriesSettings _settings;
        private bool _isUpdatingUI = false;

        private bool _userChangedColor = false;
        private string _originalShapeColor = "#E53935";
        private DateTime _windowOpenedTime;

        public ProgressSeriesToolbar(PowerPointManager manager, PowerPoint.Shape targetShape)
        {
            try
            {
                InitializeComponent();
                _manager = manager;
                _targetShape = targetShape;
                _windowOpenedTime = DateTime.Now;
                
                LoadSettings();
                UpdateUI();

                if (_targetShape == null)
                {
                    ApplyChanges();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ProgressSeriesToolbar constructor error: {ex.Message}");
                throw;
            }
        }

        public void UpdateTarget(PowerPoint.Shape newTarget)
        {
            if (_isUpdatingUI) return; // Prevent recursive updates during apply
            _targetShape = newTarget;
            LoadSettings();
            UpdateUI();
        }

        private void LoadSettings()
        {
            try
            {
                if (_targetShape == null) 
                {
                    _settings = ProgressSeriesFeature.SeriesSettings.Default();
                    return;
                }
                string tag = _targetShape.AlternativeText ?? "";
                _settings = ProgressSeriesFeature.SeriesSettings.FromTag(tag);
                
                // Get the actual current color from the active shape in the series
                _originalShapeColor = GetActualShapeColor();
                _userChangedColor = false; // Reset flag when loading settings
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadSettings error: {ex.Message}");
                _settings = ProgressSeriesFeature.SeriesSettings.Default();
                _originalShapeColor = GetActualShapeColor();
                _userChangedColor = false;
            }
        }

        private string GetActualShapeColor()
        {
            try
            {
                if (_targetShape == null) 
                {
                    System.Diagnostics.Debug.WriteLine("GetActualShapeColor: _targetShape is null");
                    return "#E53935";
                }
                
                System.Diagnostics.Debug.WriteLine($"GetActualShapeColor: Analyzing shape {_targetShape.Name}, type {_targetShape.Type}");
                
                // For now, use the tag color to avoid COM exceptions from ungrouping
                // TODO: Implement safer color detection without ungrouping
                System.Diagnostics.Debug.WriteLine("GetActualShapeColor: Using tag color to avoid COM exceptions");
                return _settings.HighlightColor;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetActualShapeColor error: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }
            
            // Fallback to default color
            System.Diagnostics.Debug.WriteLine("GetActualShapeColor: Using fallback default color");
            return "#E53935";
        }

        private void UpdateUI()
        {
            _isUpdatingUI = true;
            
            if (TxtCount != null) TxtCount.Text = _settings.Count.ToString();
            if (TxtActive != null) TxtActive.Text = _settings.ActiveIndex.ToString();
            if (SliderActive != null)
            {
                SliderActive.Maximum = _settings.Count;
                SliderActive.Value = _settings.ActiveIndex;
            }
            if (TxtOrientation != null) TxtOrientation.Text = _settings.Layout == "H" ? "" : "";
            
            // Set mode index
            if (CmbMode != null)
            {
                switch (_settings.Mode)
                {
                    case "Letters":
                        CmbMode.SelectedIndex = 1;
                        break;
                    case "SmallLetters":
                        CmbMode.SelectedIndex = 2;
                        break;
                    case "Roman":
                        CmbMode.SelectedIndex = 3;
                        break;
                    case "SmallRoman":
                        CmbMode.SelectedIndex = 4;
                        break;
                    case "Bullets":
                        CmbMode.SelectedIndex = 5;
                        break;
                    default:
                        CmbMode.SelectedIndex = 0;
                        break;
                }
            }

            // Set shape index
            if (CmbShape != null)
            {
                switch (_settings.ShapeType)
                {
                    case "Circle":
                        CmbShape.SelectedIndex = 1;
                        break;
                    case "Diamond":
                        CmbShape.SelectedIndex = 2;
                        break;
                    case "Triangle":
                        CmbShape.SelectedIndex = 3;
                        break;
                    case "Star":
                        CmbShape.SelectedIndex = 4;
                        break;
                    case "Check":
                        CmbShape.SelectedIndex = 5;
                        break;
                    case "Cross":
                        CmbShape.SelectedIndex = 6;
                        break;
                    case "Dot":
                        CmbShape.SelectedIndex = 7;
                        break;
                    default:
                        CmbShape.SelectedIndex = 0;
                        break;
                }
            }

            _isUpdatingUI = false;
        }

        private void ApplyChanges()
        {
            // Preserve the original color unless user explicitly changed it
            if (!_userChangedColor)
            {
                _settings.HighlightColor = _originalShapeColor;
            }
            
            ProgressSeriesFeature.Execute(_manager, _settings, _targetShape);
            
            // After re-creating the group, the target shape reference changes
            if (_manager != null)
            {
                var selection = _manager.GetSelectedShapes();
                if (selection != null && selection.Count == 1)
                {
                    _targetShape = selection[1];
                    // Update the original color with the actual current color of the active shape
                    if (!_userChangedColor)
                    {
                        _originalShapeColor = GetActualShapeColor();
                    }
                }
            }
        }

        private void SliderActive_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_isUpdatingUI) return;
            int newVal = (int)SliderActive?.Value;
            if (_settings.ActiveIndex != newVal)
            {
                _settings.ActiveIndex = newVal;
                if (TxtActive != null) TxtActive.Text = _settings.ActiveIndex.ToString();
                ApplyChanges();
            }
        }

        private void Setting_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingUI) return;
            
            bool changed = false;
            if (CmbMode?.SelectedItem is ComboBoxItem modeItem && modeItem.Tag != null)
            {
                if (_settings.Mode != modeItem.Tag.ToString())
                {
                    _settings.Mode = modeItem.Tag.ToString();
                    changed = true;
                }
            }
            if (CmbShape?.SelectedItem is ComboBoxItem shapeItem && shapeItem.Tag != null)
            {
                if (_settings.ShapeType != shapeItem.Tag.ToString())
                {
                    _settings.ShapeType = shapeItem.Tag.ToString();
                    changed = true;
                }
            }
            if (changed) ApplyChanges();
        }

        private void BtnOrientation_Click(object sender, RoutedEventArgs e)
        {
            _settings.Layout = _settings.Layout == "H" ? "V" : "H";
            UpdateUI();
            ApplyChanges();
        }

        private void Color_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string color)
            {
                _settings.HighlightColor = color;
                _userChangedColor = true; // Mark that user explicitly changed the color
                _originalShapeColor = color; // Update the original color to match user's choice
                ApplyChanges();
            }
        }

        private void BtnColorWell_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.ColorDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var color = dialog.Color;
                _settings.HighlightColor = string.Format("#{0:X2}{1:X2}{2:X2}", color.R, color.G, color.B);
                _userChangedColor = true; // Mark that user explicitly changed the color
                _originalShapeColor = _settings.HighlightColor; // Update the original color
                ApplyChanges();
            }
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            _settings = ProgressSeriesFeature.SeriesSettings.Default();
            UpdateUI();
            ApplyChanges();
        }

        private void Input_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                ValidateAndApplyCount();
            }
        }

        private void Input_LostFocus(object sender, RoutedEventArgs e)
        {
            ValidateAndApplyCount();
        }

        private void ValidateAndApplyCount()
        {
            if (TxtCount?.Text != null && int.TryParse(TxtCount.Text, out int result))
            {
                if (result < 1) result = 1;
                if (result > 50) result = 50; 
                if (_settings.Count != result)
                {
                    _settings.Count = result;
                    if (_settings.ActiveIndex > _settings.Count) _settings.ActiveIndex = _settings.Count;
                    UpdateUI();
                    ApplyChanges();
                }
            }
            else
            {
                if (TxtCount != null) TxtCount.Text = _settings.Count.ToString();
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) this.DragMove();
        }
    }
}
