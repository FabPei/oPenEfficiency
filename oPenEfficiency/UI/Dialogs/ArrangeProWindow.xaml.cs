using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using oPenEfficiency.Utils;
using oPenEfficiency.Features;

namespace oPenEfficiency.UI
{
    public partial class ArrangeProWindow : Window
    {
        private readonly PowerPointManager _powerPointManager;

        public ArrangeProWindow(PowerPointManager powerPointManager)
        {
            InitializeComponent();
            _powerPointManager = powerPointManager;

            // Position near selection
            try
            {
                _powerPointManager.GetSelectionScreenLocation(out int x, out int y);

                double dpiScaleX = 1.0;
                double dpiScaleY = 1.0;
                var source = PresentationSource.FromVisual(this);
                if (source?.CompositionTarget != null)
                {
                    dpiScaleX = source.CompositionTarget.TransformToDevice.M11;
                    dpiScaleY = source.CompositionTarget.TransformToDevice.M22;
                }

                this.Left = (x / dpiScaleX) - (this.Width / 2);
                this.Top = (y / dpiScaleY) - (this.Height / 2);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ArrangeProWindow positioning error: {ex.Message}");
                // Fallback: center on screen
                this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
        }

        private bool _isUpdating = false;
        private float? _anchorCenterX = null;
        private float? _anchorCenterY = null;

        private void OnSettingChanged(object sender, EventArgs e)
        {
            if (_isUpdating || _powerPointManager == null) return;
            ApplyArrangement();
        }

        private void ApplyArrangement()
        {
            try
            {
                _isUpdating = true;

                string shapeType = (ShapeTypeCombo.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "Circle";
                int instances = int.TryParse(InstancesBox.Text, out int inst) ? inst : 8;
                float radius = (float)RadiusSlider.Value;
                float startAngle = (float)AngleSlider.Value;
                bool centerOnSlide = CenterSlide.IsChecked == true;
                bool rotateToCenter = RotateToCenterCheck.IsChecked == true;
                bool fillShape = FillShapeCheck.IsChecked == true;

                // Handle anchor logic to prevent drift
                if (centerOnSlide)
                {
                    _anchorCenterX = null;
                    _anchorCenterY = null;
                }
                else
                {
                    // Basic selection detection: if count changes, reset anchor
                    _powerPointManager.GetSelectionScreenLocation(out int sx, out int sy); 
                    // This is a bit of a hack, but let's just use the current selection's center 
                    // if we don't have an anchor yet.
                    if (_anchorCenterX == null)
                    {
                        var selection = _powerPointManager.GetSelectedShapes();
                        if (selection != null && selection.Count > 0)
                        {
                            float minL = float.MaxValue, minT = float.MaxValue, maxR = float.MinValue, maxB = float.MinValue;
                            for (int i = 1; i <= selection.Count; i++)
                            {
                                var s = selection[i];
                                if (s.Left < minL) minL = s.Left;
                                if (s.Top < minT) minT = s.Top;
                                if (s.Left + s.Width > maxR) maxR = s.Left + s.Width;
                                if (s.Top + s.Height > maxB) maxB = s.Top + s.Height;
                            }
                            _anchorCenterX = (minL + maxR) / 2f;
                            _anchorCenterY = (minT + maxB) / 2f;
                        }
                    }
                }

                ArrangeInShapeFeature.ArrangeInShapePro(_powerPointManager, shapeType, instances, radius, startAngle, centerOnSlide, rotateToCenter, fillShape, _anchorCenterX, _anchorCenterY);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error applying arrangement: " + ex.Message);
            }
            finally
            {
                _isUpdating = false;
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e) => OnSettingChanged(sender, e);
        private void Combo_SelectionChanged(object sender, SelectionChangedEventArgs e) => OnSettingChanged(sender, e);
        private void Radio_Checked(object sender, RoutedEventArgs e) => OnSettingChanged(sender, e);
        private void Check_Changed(object sender, RoutedEventArgs e) => OnSettingChanged(sender, e);
    }

    public class IndexToBoolConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (int)value == 0; // 0 is Circle
        }
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) => throw new NotImplementedException();
    }

    public class IndexToOpacityConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (int)value == 0 ? 1.0 : 0.4;
        }
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) => throw new NotImplementedException();
    }
}
