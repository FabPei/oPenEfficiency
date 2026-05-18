using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public partial class FontSizeControl : UserControl
    {
        private readonly PowerPointManager _manager;
        private bool _isUpdating = false;

        public FontSizeControl(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            LoadSizes();
        }

        private void LoadSizes()
        {
            _isUpdating = true;
            try
            {
                float[] commonSizes = { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96 };
                foreach (var s in commonSizes) ComboFontSize.Items.Add(s.ToString());
            }
            finally
            {
                _isUpdating = false;
            }
        }

        private void ComboFontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdating || _manager == null) return;
            if (ComboFontSize.SelectedItem is string sizeStr)
            {
                if (float.TryParse(sizeStr, out float val))
                {
                    ApplyFontSize(val);
                }
            }
        }

        private void ComboFontSize_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (float.TryParse(ComboFontSize.Text, out float val))
                {
                    ApplyFontSize(val);
                    e.Handled = true;
                    // Move focus away to confirm
                    Keyboard.ClearFocus();
                }
            }
        }

        private void ComboFontSize_LostFocus(object sender, RoutedEventArgs e)
        {
            if (float.TryParse(ComboFontSize.Text, out float val))
            {
                ApplyFontSize(val);
            }
        }

        private void ApplyFontSize(float size)
        {
            try
            {
                var app = _manager.GetApplication();
                if (app.ActiveWindow == null) return;
                var selection = app.ActiveWindow.Selection;

                if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    selection.TextRange.Font.Size = size;
                }
                else if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    foreach (Shape shape in selection.ShapeRange)
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            shape.TextFrame.TextRange.Font.Size = size;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error applying font size: " + ex.Message);
            }
        }
    }
}
