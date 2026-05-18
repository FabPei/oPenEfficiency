using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using oPenEfficiency.Features;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI
{
    public partial class NumerationToolbar : Window
    {
        private PowerPointManager _manager;
        private PowerPoint.Shape _targetShape;
        private NumerationFeature.NumerationSettings _settings;
        private bool _isUpdatingUI = false;
        private bool _isDialogOpened = false;

        public NumerationToolbar(PowerPointManager manager, PowerPoint.Shape targetShape)
        {
            InitializeComponent();
            _manager = manager;
            _targetShape = targetShape;
            
            LoadSettings();
            UpdateUI();
        }

        private void LoadSettings()
        {
            if (_targetShape == null) return;
            string tag = _targetShape.AlternativeText ?? "";
            _settings = NumerationFeature.NumerationSettings.FromTag(tag);
        }

        private void UpdateUI()
        {
            _isUpdatingUI = true;
            TxtNumber.Text = _settings.Number.ToString();
            TxtSize.Text = Math.Round(_settings.Size, 0).ToString();
            BtnColorWell.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(_settings.ColorHex));
            
            // Set shape index
            int index = 0;
            if (_settings.ShapeType == "Square") index = 1;
            else if (_settings.ShapeType == "Diamond") index = 2;
            else if (_settings.ShapeType == "Hexagon") index = 3;
            CmbShape.SelectedIndex = index;
            
            _isUpdatingUI = false;
        }

        private void ApplyChanges()
        {
            NumerationFeature.Execute(_manager, _targetShape, _settings);
            
            // Re-bind to the new shape
            var selection = _manager.GetSelectedShapes();
            if (selection != null && selection.Count == 1)
            {
                _targetShape = selection[1];
            }
        }

        private void Setting_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingUI) return;
            
            if (CmbShape.SelectedItem is ComboBoxItem shapeItem)
            {
                _settings.ShapeType = shapeItem.Tag.ToString();
            }
            ApplyChanges();
        }

        private void BtnColorWell_Click(object sender, RoutedEventArgs e)
        {
            _isDialogOpened = true;
            var dialog = new System.Windows.Forms.ColorDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var color = dialog.Color;
                _settings.ColorHex = string.Format("#{0:X2}{1:X2}{2:X2}", color.R, color.G, color.B);
                BtnColorWell.Background = new SolidColorBrush(Color.FromRgb(color.R, color.G, color.B));
                ApplyChanges();
            }
            _isDialogOpened = false;
        }

        private void Input_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                ValidateAndApplyInputs();
            }
        }

        private void Input_LostFocus(object sender, RoutedEventArgs e)
        {
            ValidateAndApplyInputs();
        }

        private void ValidateAndApplyInputs()
        {
            if (_isUpdatingUI) return;

            bool changed = false;
            
            if (int.TryParse(TxtNumber.Text, out int n))
            {
                if (n != _settings.Number) { _settings.Number = n; changed = true; }
            }
            
            if (float.TryParse(TxtSize.Text, out float s))
            {
                if (s < 5) s = 5;
                if (s > 200) s = 200;
                if (Math.Abs(s - _settings.Size) > 0.1) { _settings.Size = s; changed = true; }
            }

            if (changed) ApplyChanges();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) this.DragMove();
        }

        private void Window_Deactivated(object sender, EventArgs e)
        {
            if (!_isDialogOpened)
            {
                try { this.Close(); } catch { }
            }
        }
    }
}
