using System.Windows;
using System.Text.RegularExpressions;
using System;
using System.Collections.Generic;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Features;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public partial class SplitShapeWindow : Window
    {
        private PowerPointManager _manager;
        private PowerPoint.Shape _originalShape;
        private List<PowerPoint.Shape> _createdShapes = new List<PowerPoint.Shape>();
        private bool _isClosing = false;

        public SplitShapeWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            this.Loaded += (s, e) => {
                PositionWindow();
                CaptureOriginalShape();
            };
            this.Closed += Window_Closed;
        }

        private void PositionWindow()
        {
            var screenWidth = SystemParameters.PrimaryScreenWidth;
            var screenHeight = SystemParameters.PrimaryScreenHeight;
            this.Left = (screenWidth - this.Width) / 2;
            this.Top = (screenHeight - this.Height) / 2;
        }

        private void CaptureOriginalShape()
        {
            try
            {
                var selection = _manager.GetSelectedShapes();
                if (selection != null && selection.Count == 1)
                {
                    _originalShape = selection[1];
                }
            }
            catch { }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
             if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private void DecimalValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9,.]");
            e.Handled = regex.IsMatch(e.Text);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            CloseCleanup();
            this.Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            CloseCleanup();
            this.Close();
        }

        private void CloseCleanup()
        {
            _isClosing = true;
            if (_originalShape != null && _createdShapes.Count > 0)
            {
                try { _originalShape.Delete(); } catch { }
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            if (!_isClosing)
            {
                CloseCleanup();
            }
        }

        private void BtnSplit_Click(object sender, RoutedEventArgs e)
        {
            if (_originalShape == null)
            {
                CaptureOriginalShape();
                if (_originalShape == null)
                {
                    MessageBox.Show("Please select exactly one shape to split.", "Selection Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            if (int.TryParse(TxtRows.Text, out int rows) &&
                int.TryParse(TxtColumns.Text, out int cols) &&
                double.TryParse(TxtSpacing.Text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double spacing))
            {
                if (rows < 1 || cols < 1)
                {
                    MessageBox.Show("Rows and Columns must be at least 1.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (spacing < 0) spacing = 0;

                bool hLines = ChkHorizontalLines.IsChecked == true;
                bool vLines = ChkVerticalLines.IsChecked == true;

                // Remove previously created shapes
                if (_createdShapes.Count > 0)
                {
                    foreach (var s in _createdShapes)
                    {
                        try { s.Delete(); } catch { }
                    }
                    _createdShapes.Clear();
                }
                
                // Hide original
                try { _originalShape.Visible = Office.MsoTriState.msoFalse; } catch { }

                var result = SplitShapeToGridFeature.Execute(_manager, _originalShape, rows, cols, (float)spacing, hLines, vLines);
                
                if (result != null && result.Count > 0)
                {
                    _createdShapes = result;
                    
                    // Select new shapes
                    try
                    {
                        var app = _manager.GetApplication();
                        if (app != null && app.ActiveWindow != null && app.ActiveWindow.View != null)
                        {
                            var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                            if (slide != null)
                            {
                                var names = new List<string>();
                                foreach (var s in _createdShapes) names.Add(s.Name);
                                slide.Shapes.Range(names.ToArray()).Select();
                            }
                        }
                    }
                    catch { }
                }
                else
                {
                    // Restore original if failed
                    try { _originalShape.Visible = Office.MsoTriState.msoTrue; } catch { }
                    MessageBox.Show("Failed to split shape.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("Please enter valid numbers.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
