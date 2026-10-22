using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    public partial class FontPanel : UserControl
    {
        private readonly PowerPointManager _manager;
        private bool _isUpdating = false;
        private string _displayedFontName;
        private string _displayedFontSize;

        public FontPanel(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            LoadData();
        }

        public void FocusFontFamily()
        {
            FocusComboBox(ComboFont);
        }

        public void FocusFontSize()
        {
            FocusComboBox(ComboFontSize);
        }

        private void FocusComboBox(ComboBox comboBox)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                comboBox.Focus();
                Keyboard.Focus(comboBox);
                if (comboBox.Template.FindName("PART_EditableTextBox", comboBox) is TextBox editor)
                {
                    editor.Focus();
                    editor.SelectAll();
                }
            }), System.Windows.Threading.DispatcherPriority.Input);
        }

        public void OnSelectionChange(Selection sel)
        {
            if (_isUpdating || sel == null) return;
            
            _isUpdating = true;
            try
            {
                var ranges = GetTextRanges(sel);
                IsEnabled = ranges.Count > 0;
                string fontName = ranges.Count > 0 ? ranges[0].Font.Name : null;
                float fontSize = ranges.Count > 0 ? ranges[0].Font.Size : 0;
                bool mixedName = ranges.Skip(1).Any(range =>
                    !string.Equals(range.Font.Name, fontName, StringComparison.OrdinalIgnoreCase));
                bool mixedSize = ranges.Skip(1).Any(range =>
                    Math.Abs(range.Font.Size - fontSize) > 0.01f);

                _displayedFontName = !mixedName && !string.IsNullOrEmpty(fontName) && fontName != "Mixed"
                    ? fontName
                    : "";
                _displayedFontSize = !mixedSize && fontSize > 0 && fontSize < 4000
                    ? fontSize.ToString(CultureInfo.CurrentCulture)
                    : "";
                ComboFont.Text = _displayedFontName;
                ComboFontSize.Text = _displayedFontSize;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FontPanel.OnSelectionChange");
                IsEnabled = false;
                ClearDisplay();
            }
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
            if (ComboFont.SelectedItem is string fontName) ApplyFontName(fontName);
        }

        private void ComboFontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdating || _manager == null) return;
            if (ComboFontSize.SelectedItem is string s && Core.Services.FontEditPolicy.TryParseSize(s, out float value))
                ApplyFontSize(value);
        }

        private void ComboFont_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                ApplyFontName(ComboFont.Text);
                e.Handled = true;
                Keyboard.ClearFocus();
            }
            else if (e.Key == Key.Escape)
            {
                RefreshFromSelection();
                e.Handled = true;
            }
        }

        private void ComboFont_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!_isUpdating) ApplyFontName(ComboFont.Text);
        }

        private void ComboFontSize_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                ApplySizeFromText();
                e.Handled = true;
                Keyboard.ClearFocus();
            }
            else if (e.Key == Key.Escape)
            {
                RefreshFromSelection();
                e.Handled = true;
            }
        }

        private void ComboFontSize_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_isUpdating) return;
            ApplySizeFromText();
        }

        private void ApplyFontName(string fontName)
        {
            fontName = fontName?.Trim();
            if (string.IsNullOrEmpty(fontName) ||
                string.Equals(fontName, _displayedFontName, StringComparison.Ordinal)) return;
            ApplyProperty("Name", fontName);
        }

        private void ApplySizeFromText()
        {
            if (string.Equals(ComboFontSize.Text, _displayedFontSize, StringComparison.Ordinal)) return;
            if (Core.Services.FontEditPolicy.TryParseSize(ComboFontSize.Text, out float value))
                ApplyFontSize(value);
            else
                RefreshFromSelection();
        }

        private void ApplyFontSize(float value)
        {
            if (!Core.Services.FontEditPolicy.IsValidFontSize(value))
            {
                RefreshFromSelection();
                return;
            }
            ApplyProperty("Size", value);
        }

        private bool ApplyProperty(string property, object value)
        {
            var states = new List<FontState>();
            try
            {
                var app = _manager.GetApplication();
                if (app.ActiveWindow?.Selection == null) return false;
                var selection = app.ActiveWindow.Selection;
                var ranges = GetTextRanges(selection);
                if (ranges.Count == 0) return false;

                foreach (var range in ranges)
                {
                    var state = new FontState(range);
                    bool changes = property == "Name"
                        ? Core.Services.FontEditPolicy.NeedsNameChange(state.Name, (string)value)
                        : Core.Services.FontEditPolicy.NeedsSizeChange(state.Size, (float)value);
                    if (changes) states.Add(state);
                }
                if (states.Count == 0)
                {
                    RefreshFromSelection();
                    return false;
                }

                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "FontPanel.StartNewUndoEntry"); }

                foreach (var state in states)
                {
                    if (property == "Name") state.Range.Font.Name = (string)value;
                    else state.Range.Font.Size = (float)value;
                }
                RefreshFromSelection();
                return true;
            }
            catch (Exception ex)
            {
                foreach (var state in states)
                {
                    try
                    {
                        state.Range.Font.Name = state.Name;
                        state.Range.Font.Size = state.Size;
                    }
                    catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "FontPanel.Rollback"); }
                }
                ExceptionLogger.Log(ex, "FontPanel.ApplyProperty");
                RefreshFromSelection();
                return false;
            }
        }

        private static List<TextRange> GetTextRanges(Selection selection)
        {
            var ranges = new List<TextRange>();
            if (selection == null) return ranges;
            if (selection.Type == PpSelectionType.ppSelectionText)
            {
                if (selection.TextRange != null) ranges.Add(selection.TextRange);
                return ranges;
            }
            if (selection.Type != PpSelectionType.ppSelectionShapes) return ranges;

            ShapeRange shapes = null;
            try { if (selection.HasChildShapeRange) shapes = selection.ChildShapeRange; }
            catch { /* Non-group selections do not expose ChildShapeRange. */ }
            if (shapes == null)
            {
                try { shapes = selection.ShapeRange; }
                catch { /* Selection types without ShapeRange yield no font targets. */ }
            }
            if (shapes == null) return ranges;
            for (int index = 1; index <= shapes.Count; index++)
            {
                var shape = shapes[index];
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    ranges.Add(shape.TextFrame.TextRange);
            }
            return ranges;
        }

        private void RefreshFromSelection()
        {
            try { OnSelectionChange(_manager?.GetApplication()?.ActiveWindow?.Selection); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "FontPanel.RefreshFromSelection"); }
        }

        private void ClearDisplay()
        {
            _isUpdating = true;
            _displayedFontName = "";
            _displayedFontSize = "";
            ComboFont.Text = "";
            ComboFontSize.Text = "";
            _isUpdating = false;
        }

        private sealed class FontState
        {
            public FontState(TextRange range)
            {
                Range = range;
                Name = range.Font.Name;
                Size = range.Font.Size;
            }

            public TextRange Range { get; }
            public string Name { get; }
            public float Size { get; }
        }
    }
}
