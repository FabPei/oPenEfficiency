using System;
using System.Collections.Generic;
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

        public ArrangeProWindow(
            PowerPointManager powerPointManager,
            ArrangeProFeature.Preset preset = ArrangeProFeature.Preset.Custom)
        {
            InitializeComponent();
            _powerPointManager = powerPointManager;
            SetPreset(preset);
            Loaded += (sender, args) => PositionNearSelection();
        }

        public void SetPreset(ArrangeProFeature.Preset preset)
        {
            if (preset == ArrangeProFeature.Preset.Custom || ShapeTypeCombo == null) return;

            bool wasUpdating = _isUpdating;
            _isUpdating = true;
            try
            {
                ShapeTypeCombo.SelectedIndex = preset == ArrangeProFeature.Preset.ArcAroundSelection ? 1 : 0;
                CenterSelection.IsChecked = preset != ArrangeProFeature.Preset.CircleAroundSlide;
                CenterSlide.IsChecked = preset == ArrangeProFeature.Preset.CircleAroundSlide;
                CenterFirstShape.IsChecked = false;
                CenterLastShape.IsChecked = false;
                RotateToCenterCheck.IsChecked = preset == ArrangeProFeature.Preset.RotateToCenter;
            }
            finally
            {
                _isUpdating = wasUpdating;
            }

            ApplyArrangement();
        }

        private void PositionNearSelection()
        {
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

                Left = Math.Max(SystemParameters.WorkArea.Left, Math.Min((x / dpiScaleX) - (ActualWidth / 2), SystemParameters.WorkArea.Right - ActualWidth));
                Top = Math.Max(SystemParameters.WorkArea.Top, Math.Min((y / dpiScaleY) - (ActualHeight / 2), SystemParameters.WorkArea.Bottom - ActualHeight));
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ArrangeProWindow.PositionNearSelection");
            }
        }

        private bool _isUpdating = false;
        private float? _anchorCenterX = null;
        private float? _anchorCenterY = null;
        private string _selectionKey;
        private Dictionary<int, Tuple<float, float>> _originalOrderPositions =
            new Dictionary<int, Tuple<float, float>>();

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
                ArrangeProAnchorMode anchorMode = CenterFirstShape.IsChecked == true
                    ? ArrangeProAnchorMode.FirstShape
                    : CenterLastShape.IsChecked == true
                        ? ArrangeProAnchorMode.LastShape
                        : ArrangeProAnchorMode.SelectionCenter;
                ArrangeProOrder order = OrderCombo.SelectedIndex == 1
                    ? ArrangeProOrder.LeftToRight
                    : OrderCombo.SelectedIndex == 2
                        ? ArrangeProOrder.TopToBottom
                        : ArrangeProOrder.Selection;
                bool clockwise = DirectionCombo.SelectedIndex != 1;
                float arcSpan = (float)ArcSpanSlider.Value;
                bool rotateToCenter = RotateToCenterCheck.IsChecked == true;
                bool fillShape = FillShapeCheck.IsChecked == true;
                bool avoidOverlap = AvoidOverlapCheck.IsChecked == true;

                var currentSelection = _powerPointManager.GetSelectedShapes();
                if (currentSelection == null || currentSelection.Count == 0) return;
                var selectedIds = new string[currentSelection.Count];
                for (int i = 1; i <= currentSelection.Count; i++) selectedIds[i - 1] = currentSelection[i].Id.ToString();
                string selectionKey = string.Join(",", selectedIds);
                if (!string.Equals(selectionKey, _selectionKey, StringComparison.Ordinal))
                {
                    _selectionKey = selectionKey;
                    _anchorCenterX = null;
                    _anchorCenterY = null;
                    _originalOrderPositions = new Dictionary<int, Tuple<float, float>>();
                    for (int i = 1; i <= currentSelection.Count; i++)
                    {
                        var shape = currentSelection[i];
                        _originalOrderPositions[shape.Id] = Tuple.Create(
                            shape.Left + shape.Width / 2f,
                            shape.Top + shape.Height / 2f);
                    }
                }

                // Handle anchor logic to prevent drift
                if (centerOnSlide)
                {
                    _anchorCenterX = null;
                    _anchorCenterY = null;
                }
                else
                {
                    if (_anchorCenterX == null)
                    {
                        if (currentSelection.Count > 0)
                        {
                            float minL = float.MaxValue, minT = float.MaxValue, maxR = float.MinValue, maxB = float.MinValue;
                            for (int i = 1; i <= currentSelection.Count; i++)
                            {
                                var s = currentSelection[i];
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

                ArrangeInShapeFeature.ArrangeInShapePro(
                    _powerPointManager,
                    shapeType,
                    instances,
                    radius,
                    startAngle,
                    centerOnSlide,
                    rotateToCenter,
                    fillShape,
                    _anchorCenterX,
                    _anchorCenterY,
                    anchorMode,
                    order,
                    clockwise,
                    arcSpan,
                    avoidOverlap,
                    _originalOrderPositions);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ArrangeProWindow.ApplyArrangement");
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
        private void Combo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FillShapeCheck != null)
            {
                string shapeType = (ShapeTypeCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Circle";
                bool supportsFill = shapeType == "Circle" || shapeType == "Triangle" || shapeType == "Square";
                bool isArc = shapeType == "Arc";
                bool wasUpdating = _isUpdating;
                _isUpdating = true;
                FillShapeCheck.IsEnabled = supportsFill;
                if (!supportsFill) FillShapeCheck.IsChecked = false;
                if (ArcSpanPanel != null)
                {
                    ArcSpanPanel.IsEnabled = isArc;
                    ArcSpanPanel.Opacity = isArc ? 1.0 : 0.45;
                }
                _isUpdating = wasUpdating;
            }
            OnSettingChanged(sender, e);
        }
        private void OrderCombo_SelectionChanged(object sender, SelectionChangedEventArgs e) => OnSettingChanged(sender, e);
        private void Radio_Checked(object sender, RoutedEventArgs e) => OnSettingChanged(sender, e);
        private void Check_Changed(object sender, RoutedEventArgs e) => OnSettingChanged(sender, e);
    }

    public class IndexToBoolConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (int)value == 0; // 0 is Circle
        }
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) => System.Windows.Data.Binding.DoNothing;
    }

    public class IndexToOpacityConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (int)value == 0 ? 1.0 : 0.4;
        }
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) => System.Windows.Data.Binding.DoNothing;
    }
}
