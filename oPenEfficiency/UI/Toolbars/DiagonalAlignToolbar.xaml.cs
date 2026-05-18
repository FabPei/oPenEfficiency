using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using oPenEfficiency.Services;

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
            var app = _manager.GetApplication();
            if (app?.ActiveWindow?.Selection?.Type == PpSelectionType.ppSelectionShapes)
            {
                var sr = app.ActiveWindow.Selection.ShapeRange;
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
                    // Basic ordering to determine endpoints for initial angle estimation
                    var sortedShapes = _shapes.OrderBy(s => s.OrigCenterX + s.OrigCenterY).ToList();
                    var first = sortedShapes.First();
                    var last = sortedShapes.Last();

                    float dx = last.OrigCenterX - first.OrigCenterX;
                    float dy = last.OrigCenterY - first.OrigCenterY;
                    float dist = (float)Math.Sqrt(dx * dx + dy * dy);
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
            float angleRad = (float)(angleDeg * Math.PI / 180f);

            float dirX = (float)Math.Cos(angleRad);
            float dirY = (float)Math.Sin(angleRad);

            // First project shapes based on the generic center (0,0) or average center to get the order right.
            float avgX = _shapes.Average(s => s.OrigCenterX);
            float avgY = _shapes.Average(s => s.OrigCenterY);

            var projectedShapes = _shapes.Select(s => new
            {
                Data = s,
                Proj = (s.OrigCenterX - avgX) * dirX + (s.OrigCenterY - avgY) * dirY
            }).OrderBy(x => x.Proj).ToList();

            // Determine Anchor Center Point
            string anchorType = (CmbAnchor?.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "Center";
            float anchorOrigX = avgX;
            float anchorOrigY = avgY;
            float anchorIndex = (projectedShapes.Count - 1) / 2f;

            if (anchorType == "First")
            {
                anchorOrigX = projectedShapes.First().Data.OrigCenterX;
                anchorOrigY = projectedShapes.First().Data.OrigCenterY;
                anchorIndex = 0;
            }
            else if (anchorType == "Last")
            {
                anchorOrigX = projectedShapes.Last().Data.OrigCenterX;
                anchorOrigY = projectedShapes.Last().Data.OrigCenterY;
                anchorIndex = projectedShapes.Count - 1;
            }

            for (int i = 0; i < projectedShapes.Count; i++)
            {
                var item = projectedShapes[i];
                float offset = (i - anchorIndex) * spacing;
                
                float newCenterX = anchorOrigX + offset * dirX;
                float newCenterY = anchorOrigY + offset * dirY;

                try
                {
                    item.Data.Shape.Left = newCenterX - (item.Data.OrigWidth / 2f);
                    item.Data.Shape.Top = newCenterY - (item.Data.OrigHeight / 2f);
                }
                catch { } // Ignore if shape was deleted
            }
        }
        
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
