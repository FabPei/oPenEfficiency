using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI
{
    public partial class FontControl : UserControl
    {
        private readonly PowerPointManager _manager;
        private bool _isUpdating = false;

        public FontControl(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            LoadFonts();
        }

        private void LoadFonts()
        {
            _isUpdating = true;
            try
            {
                var fontNames = System.Windows.Media.Fonts.SystemFontFamilies
                                     .Select(f => f.Source)
                                     .OrderBy(s => s)
                                     .ToList();

                if (fontNames.Count == 0)
                {
                    fontNames = new List<string> { "Arial", "Calibri", "Segoe UI", "Tahoma", "Times New Roman", "Verdana" };
                }

                ComboFont.ItemsSource = fontNames;
            }
            finally
            {
                _isUpdating = false;
            }
        }

        private void ComboFont_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdating || _manager == null) return;
            if (ComboFont.SelectedItem is string fontName)
            {
                ApplyFont(fontName);
            }
        }

        private void ApplyFont(string fontName)
        {
            try
            {
                var app = _manager.GetApplication();
                if (app.ActiveWindow == null) return;
                var selection = app.ActiveWindow.Selection;

                if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    selection.TextRange.Font.Name = fontName;
                }
                else if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    foreach (Shape shape in selection.ShapeRange)
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            shape.TextFrame.TextRange.Font.Name = fontName;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Silently fail if PowerPoint is busy
                System.Diagnostics.Debug.WriteLine("Error applying font: " + ex.Message);
            }
        }
    }
}
