using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace oPenEfficiency.UI
{
    public partial class WpfColorPickerDialog : Window
    {
        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                this.DragMove();
        }
        
        private void BtnClose_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            this.Close();
        }

        public Color SelectedColor { get; private set; }
        private bool _isUpdating = false;

        public WpfColorPickerDialog(string initialHex)
        {
            InitializeComponent();
            try
            {
                if (!string.IsNullOrEmpty(initialHex))
                {
                    SelectedColor = (Color)ColorConverter.ConvertFromString(initialHex);
                    UpdateUIFromColor();
                }
                else
                {
                    SelectedColor = Colors.White;
                    UpdateUIFromColor();
                }
            }
            catch
            {
                SelectedColor = Colors.White;
                UpdateUIFromColor();
            }
        }

        private void UpdateUIFromColor()
        {
            _isUpdating = true;
            RSlider.Value = SelectedColor.R;
            GSlider.Value = SelectedColor.G;
            BSlider.Value = SelectedColor.B;
            HexTextBox.Text = $"#{SelectedColor.R:X2}{SelectedColor.G:X2}{SelectedColor.B:X2}";
            ColorPreview.Background = new SolidColorBrush(SelectedColor);
            _isUpdating = false;
        }

        private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_isUpdating) return;
            SelectedColor = Color.FromRgb((byte)RSlider.Value, (byte)GSlider.Value, (byte)BSlider.Value);
            _isUpdating = true;
            HexTextBox.Text = $"#{SelectedColor.R:X2}{SelectedColor.G:X2}{SelectedColor.B:X2}";
            ColorPreview.Background = new SolidColorBrush(SelectedColor);
            _isUpdating = false;
        }

        private void HexTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdating) return;
            try
            {
                if (HexTextBox.Text.StartsWith("#") && (HexTextBox.Text.Length == 7 || HexTextBox.Text.Length == 9))
                {
                    var color = (Color)ColorConverter.ConvertFromString(HexTextBox.Text);
                    SelectedColor = color;
                    _isUpdating = true;
                    RSlider.Value = SelectedColor.R;
                    GSlider.Value = SelectedColor.G;
                    BSlider.Value = SelectedColor.B;
                    ColorPreview.Background = new SolidColorBrush(SelectedColor);
                    _isUpdating = false;
                }
            }
            catch { }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}