using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    public partial class PositionPainterWindow : Window, INotifyPropertyChanged
    {
        private readonly PowerPointManager _pm;
        private Presentation _sourcePresentation;
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

        public PositionPainterWindow(
            PowerPointManager pm,
            bool applyLeft = true,
            bool applyTop = true,
            bool applyWidth = true,
            bool applyHeight = true)
        {
            InitializeComponent();
            _pm = pm;
            DataContext = this;
            ConfigureProperties(applyLeft, applyTop, applyWidth, applyHeight);
            Learn();
        }

        public void ConfigureProperties(bool applyLeft, bool applyTop, bool applyWidth, bool applyHeight)
        {
            CbLeft.IsChecked = applyLeft;
            CbTop.IsChecked = applyTop;
            CbWidth.IsChecked = applyWidth;
            CbHeight.IsChecked = applyHeight;
        }

        private void Learn()
        {
            try
            {
                var app = _pm.GetApplication();
                var selection = app?.ActiveWindow?.Selection;
                if (selection != null &&
                    selection.Type == PpSelectionType.ppSelectionShapes &&
                    selection.ShapeRange.Count == 1)
                {
                    var shape = selection.ShapeRange[1];
                    _sourcePresentation = app.ActivePresentation;
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
                    MessageBox.Show("Select exactly one source shape.", "Copy/Paste Position & Size", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PositionPainterWindow.Learn");
                MessageBox.Show("The source shape could not be read.", "Copy/Paste Position & Size", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e) { if (e.ChangedButton == MouseButton.Left) DragMove(); }
        private void Close_Click(object sender, RoutedEventArgs e) => Close();
        private void Learn_Click(object sender, RoutedEventArgs e) => Learn();

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            bool applyLeft = CbLeft.IsChecked == true;
            bool applyTop = CbTop.IsChecked == true;
            bool applyWidth = CbWidth.IsChecked == true;
            bool applyHeight = CbHeight.IsChecked == true;
            if (!applyLeft && !applyTop && !applyWidth && !applyHeight) return;

            var states = new List<ShapeState>();
            try
            {
                var app = _pm.GetApplication();
                var activePresentation = app?.ActivePresentation;
                if (_sourcePresentation == null ||
                    activePresentation == null ||
                    !IsSameComObject(activePresentation, _sourcePresentation))
                {
                    MessageBox.Show(
                        "The active presentation changed. Return to the source presentation or learn a new source shape.",
                        "Copy/Paste Position & Size",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                var selection = app.ActiveWindow?.Selection;
                if (selection == null ||
                    selection.Type != PpSelectionType.ppSelectionShapes ||
                    selection.ShapeRange.Count < 1)
                {
                    MessageBox.Show("Select one or more target shapes.", "Copy/Paste Position & Size", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                foreach (Shape shape in selection.ShapeRange)
                    states.Add(new ShapeState(shape));

                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "PositionPainterWindow.StartNewUndoEntry"); }

                foreach (var state in states)
                {
                    var shape = state.Shape;
                    if (applyWidth || applyHeight)
                        shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                    if (applyWidth) shape.Width = _width;
                    if (applyHeight) shape.Height = _height;
                    if (applyLeft) shape.Left = _left;
                    if (applyTop) shape.Top = _top;
                    if (applyWidth || applyHeight)
                        shape.LockAspectRatio = state.LockAspectRatio;
                }
            }
            catch (Exception ex)
            {
                foreach (var state in states)
                {
                    try { state.Restore(); }
                    catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "PositionPainterWindow.Rollback"); }
                }
                ExceptionLogger.Log(ex, "PositionPainterWindow.Apply");
                MessageBox.Show("The properties could not be applied. Any partial changes were rolled back where possible.",
                    "Copy/Paste Position & Size", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static bool IsSameComObject(object first, object second)
        {
            IntPtr firstPointer = IntPtr.Zero;
            IntPtr secondPointer = IntPtr.Zero;
            try
            {
                firstPointer = Marshal.GetIUnknownForObject(first);
                secondPointer = Marshal.GetIUnknownForObject(second);
                return firstPointer == secondPointer;
            }
            finally
            {
                if (firstPointer != IntPtr.Zero) Marshal.Release(firstPointer);
                if (secondPointer != IntPtr.Zero) Marshal.Release(secondPointer);
            }
        }

        private sealed class ShapeState
        {
            public ShapeState(Shape shape)
            {
                Shape = shape;
                Left = shape.Left;
                Top = shape.Top;
                Width = shape.Width;
                Height = shape.Height;
                LockAspectRatio = shape.LockAspectRatio;
            }

            public Shape Shape { get; }
            public float Left { get; }
            public float Top { get; }
            public float Width { get; }
            public float Height { get; }
            public Office.MsoTriState LockAspectRatio { get; }

            public void Restore()
            {
                Shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                Shape.Width = Width;
                Shape.Height = Height;
                Shape.Left = Left;
                Shape.Top = Top;
                Shape.LockAspectRatio = LockAspectRatio;
            }
        }
    }
}
