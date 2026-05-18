using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public partial class FontPanel : UserControl
    {
        private readonly PowerPointManager _manager;
        private bool _isUpdating = false;

        public FontPanel(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            LoadData();
        }

        public void OnSelectionChange(Selection sel)
        {
            if (_isUpdating || sel == null) return;
            
            _isUpdating = true;
            try
            {
                string fontName = null;
                float fontSize = 0;

                // 1. Determine target range
                if (sel.Type == PpSelectionType.ppSelectionText)
                {
                    fontName = sel.TextRange.Font.Name;
                    fontSize = sel.TextRange.Font.Size;
                }
                else if (sel.Type == PpSelectionType.ppSelectionShapes || sel.Type == PpSelectionType.ppSelectionNone)
                {
                    ShapeRange range = null;
                    try { if (sel.HasChildShapeRange) range = sel.ChildShapeRange; } catch { }
                    if (range == null) try { range = sel.ShapeRange; } catch { }

                    if (range != null && range.Count > 0)
                    {
                        var shape = range[1];
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            fontName = shape.TextFrame.TextRange.Font.Name;
                            fontSize = shape.TextFrame.TextRange.Font.Size;
                        }
                    }
                }

                // 2. Update UI
                if (!string.IsNullOrEmpty(fontName) && fontName != "Mixed")
                    ComboFont.Text = fontName;
                else
                    ComboFont.Text = "";

                if (fontSize > 0 && fontSize < 1000) // 999999 means mixed
                    ComboFontSize.Text = fontSize.ToString();
                else
                    ComboFontSize.Text = "";
            }
            catch { }
            finally { _isUpdating = false; }
        }

        private void LoadData()
        {
            _isUpdating = true;
            try
            {
                // Load Fonts
                var fontNames = System.Windows.Media.Fonts.SystemFontFamilies
                                     .Select(f => f.Source)
                                     .OrderBy(s => s)
                                     .ToList();
                if (fontNames.Count == 0) fontNames = new List<string> { "Arial", "Calibri", "Segoe UI" };
                ComboFont.ItemsSource = fontNames;

                // Load Sizes
                float[] commonSizes = { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96 };
                foreach (var s in commonSizes) ComboFontSize.Items.Add(s.ToString());
            }
            finally { _isUpdating = false; }
        }

        private void ComboFont_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdating || _manager == null) return;
            if (ComboFont.SelectedItem is string fontName) ApplyProperty("Name", fontName);
        }

        private void ComboFontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdating || _manager == null) return;
            if (ComboFontSize.SelectedItem is string s && float.TryParse(s, out float v)) ApplyProperty("Size", v);
        }

        private void ComboFontSize_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (float.TryParse(ComboFontSize.Text, out float val))
                {
                    ApplyProperty("Size", val);
                    e.Handled = true;
                    Keyboard.ClearFocus();
                }
            }
        }

        private void ComboFontSize_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_isUpdating) return;
            if (float.TryParse(ComboFontSize.Text, out float val)) ApplyProperty("Size", val);
        }

        private void ApplyProperty(string prop, object val)
        {
            try
            {
                var app = _manager.GetApplication();
                if (app.ActiveWindow == null) return;
                var selection = app.ActiveWindow.Selection;

                if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    if (prop == "Name") selection.TextRange.Font.Name = (string)val;
                    else if (prop == "Size") selection.TextRange.Font.Size = (float)val;
                }
                else if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    foreach (Shape shape in selection.ShapeRange)
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            if (prop == "Name") shape.TextFrame.TextRange.Font.Name = (string)val;
                            else if (prop == "Size") shape.TextFrame.TextRange.Font.Size = (float)val;
                        }
                    }
                }
            }
            catch { }
        }
    }
}
