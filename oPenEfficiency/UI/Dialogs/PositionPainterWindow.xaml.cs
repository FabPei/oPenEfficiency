using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI
{
    public partial class PositionPainterWindow : Window, INotifyPropertyChanged
    {
        private PowerPointManager _pm;
        private float _left, _top, _width, _height;
        
        private string _leftText;
        public string LeftText { get { return _leftText; } set { _leftText = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("LeftText")); } }
        
        private string _topText;
        public string TopText { get { return _topText; } set { _topText = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("TopText")); } }
        
        private string _widthText;
        public string WidthText { get { return _widthText; } set { _widthText = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("WidthText")); } }
        
        private string _heightText;
        public string HeightText { get { return _heightText; } set { _heightText = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("HeightText")); } }

        public event PropertyChangedEventHandler PropertyChanged;

        public PositionPainterWindow(PowerPointManager pm)
        {
            InitializeComponent();
            _pm = pm;
            DataContext = this;
            Learn();
        }

        private void Learn()
        {
            try
            {
                var app = _pm.GetApplication();
                var selection = app.ActiveWindow.Selection;
                if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count >= 1)
                {
                    var shape = selection.ShapeRange[1];
                    _left = shape.Left;
                    _top = shape.Top;
                    _width = shape.Width;
                    _height = shape.Height;
                    
                    LeftText = $"Left (X): {Math.Round(_left, 2)}";
                    TopText = $"Top (Y): {Math.Round(_top, 2)}";
                    WidthText = $"Width: {Math.Round(_width, 2)}";
                    HeightText = $"Height: {Math.Round(_height, 2)}";
                }
                else
                {
                    MessageBox.Show("Please select a shape first.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error learning shape: " + ex.Message);
                Close();
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e) { if (e.ChangedButton == MouseButton.Left) DragMove(); }
        private void Close_Click(object sender, RoutedEventArgs e) => Close();
        private void Learn_Click(object sender, RoutedEventArgs e) => Learn();

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var app = _pm.GetApplication();
                var selection = app.ActiveWindow.Selection;
                if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    foreach (Shape shape in selection.ShapeRange)
                    {
                        if (CbLeft.IsChecked == true) shape.Left = _left;
                        if (CbTop.IsChecked == true) shape.Top = _top;
                        if (CbWidth.IsChecked == true) shape.Width = _width;
                        if (CbHeight.IsChecked == true) shape.Height = _height;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error applying properties: " + ex.Message);
            }
        }
    }
}