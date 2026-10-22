using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using DiagonalPlacementPolicy = oPenEfficiency.Core.Services.DiagonalPlacementPolicy;

namespace oPenEfficiency.UI.Toolbars
{
    public partial class DiagonalAlignToolbar : Window
    {
        private class ShapeData
        {
            public Shape Shape { get; set; }
            public float OrigCenterX { get; set; }
            public float OrigCenterY { get; set; }
            public float OrigWidth { get; set; }
            public float OrigHeight { get; set; }
        }

        private readonly PowerPointManager _manager;
        private List<ShapeData> _shapes = new List<ShapeData>();
        private bool _isUpdating = false;

        public DiagonalAlignToolbar(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            LoadShapes();
        }

        private void LoadShapes()
        {
            var sr = _manager.GetSelectedShapes();
            if (sr != null)
            {
                for (int i = 1; i <= sr.Count; i++)
                {
                    var s = sr[i];
                    _shapes.Add(new ShapeData
                    {
                        Shape = s,
                        OrigCenterX = s.Left + (s.Width / 2f),
                        OrigCenterY = s.Top + (s.Height / 2f),
                        OrigWidth = s.Width,
                        OrigHeight = s.Height
                    });
                }

                if (_shapes.Count >= 2)
                {
                    float dx = 0f;
                    float dy = 0f;
                    float maxDistanceSquared = 0f;
                    for (int i = 0; i < _shapes.Count - 1; i++)
                    {
                        for (int j = i + 1; j < _shapes.Count; j++)
                        {
                            float candidateDx = _shapes[j].OrigCenterX - _shapes[i].OrigCenterX;
                            float candidateDy = _shapes[j].OrigCenterY - _shapes[i].OrigCenterY;
                            float candidateDistanceSquared = candidateDx * candidateDx + candidateDy * candidateDy;
                            if (candidateDistanceSquared > maxDistanceSquared)
                            {
                                maxDistanceSquared = candidateDistanceSquared;
                                dx = candidateDx;
                                dy = candidateDy;
                            }
                        }
                    }

                    // Keep the initial direction stable regardless of PowerPoint's selection order.
                    if (dx < 0f || (Math.Abs(dx) < 0.01f && dy < 0f))
                    {
                        dx = -dx;
                        dy = -dy;
                    }

                    float dist = (float)Math.Sqrt(maxDistanceSquared);
                    float avgSpacing = dist / (_shapes.Count - 1);
                    
                    float angleRad = (float)Math.Atan2(dy, dx);
                    float angleDeg = (float)(angleRad * 180 / Math.PI);

                    _isUpdating = true;
                    AngleSlider.Value = Math.Round(angleDeg);
                    AngleTextBox.Text = Math.Round(angleDeg).ToString();

                    SpacingSlider.Value = avgSpacing;
                    SpacingTextBox.Text = Math.Round(avgSpacing).ToString();
                    _isUpdating = false;
                }
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_isUpdating || AngleTextBox == null || SpacingTextBox == null) return;
            _isUpdating = true;

            if (sender == AngleSlider) AngleTextBox.Text = Math.Round(AngleSlider.Value).ToString();
            if (sender == SpacingSlider) SpacingTextBox.Text = Math.Round(SpacingSlider.Value).ToString();

            ApplyAlignment();
            _isUpdating = false;
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdating || AngleSlider == null || SpacingSlider == null) return;
            _isUpdating = true;

            if (sender == AngleTextBox && double.TryParse(AngleTextBox.Text, out double a)) AngleSlider.Value = a;
            if (sender == SpacingTextBox && double.TryParse(SpacingTextBox.Text, out double s)) SpacingSlider.Value = s;

            ApplyAlignment();
            _isUpdating = false;
        }
        
        private void BtnPreset_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag != null)
            {
                if (double.TryParse(btn.Tag.ToString(), out double a))
                {
                    AngleSlider.Value = a;
                }
            }
        }
        
        private void CmbAnchor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isUpdating) ApplyAlignment();
        }

        private void ApplyAlignment()
        {
            if (_shapes.Count < 2) return;

            float angleDeg = (float)AngleSlider.Value;
            float spacing = (float)SpacingSlider.Value;

            // Determine Anchor Center Point
            string anchorType = (CmbAnchor?.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "Center";

            // Read every shape's captured bounds, then delegate the placement (projection,
            // ordering and anchor arithmetic) to the host-free policy.
            var centerX = new float[_shapes.Count];
            var centerY = new float[_shapes.Count];
            var width = new float[_shapes.Count];
            var height = new float[_shapes.Count];
            for (int i = 0; i < _shapes.Count; i++)
            {
                centerX[i] = _shapes[i].OrigCenterX;
                centerY[i] = _shapes[i].OrigCenterY;
                width[i] = _shapes[i].OrigWidth;
                height[i] = _shapes[i].OrigHeight;
            }

            DiagonalPlacementPolicy.ComputePlacementsWithSpacing(
                centerX, centerY, width, height,
                angleDeg, spacing, DiagonalPlacementPolicy.ResolveAnchor(anchorType),
                out float[] newLeft, out float[] newTop);

            for (int i = 0; i < _shapes.Count; i++)
            {
                try
                {
                    _shapes[i].Shape.Left = newLeft[i];
                    _shapes[i].Shape.Top = newTop[i];
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "DiagonalAlignToolbar.ApplyAlignment.MoveShape");
                }
            }
        }
        
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
