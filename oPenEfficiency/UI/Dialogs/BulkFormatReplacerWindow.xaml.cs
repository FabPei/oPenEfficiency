using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using oPenEfficiency.Features;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI
{
    public partial class BulkFormatReplacerWindow : Window
    {
        private readonly PowerPointManager _manager;
        private BulkFormatReplacerFeature.FormatProps _currentProps;

        public BulkFormatReplacerWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            this.Loaded += Window_Loaded;
            LoadData();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                bool isFilled;
                // Text Pipettes
                var textFindGeom = IconService.GetIconGeometry("SharedPipetteText", out isFilled);
                if (textFindGeom != null) PathTextColorFind.Data = textFindGeom;

                var textReplaceGeom = IconService.GetIconGeometry("SharedPipetteText", true, out isFilled); // Hover variant for replace
                if (textReplaceGeom != null) PathTextColorReplace.Data = textReplaceGeom;

                // Fill Pipettes
                var fillFindGeom = IconService.GetIconGeometry("SharedPipetteFill", out isFilled);
                if (fillFindGeom != null) PathFillColorFind.Data = fillFindGeom;

                var fillReplaceGeom = IconService.GetIconGeometry("SharedPipetteFill", true, out isFilled); // Hover variant for replace
                if (fillReplaceGeom != null) PathFillColorReplace.Data = fillReplaceGeom;

                // Line Pipettes (Reuse Fill for now or use the generic one)
                var lineFindGeom = IconService.GetIconGeometry("SharedPipetteFill", out isFilled);
                if (lineFindGeom != null) PathLineColorFind.Data = lineFindGeom;

                var lineReplaceGeom = IconService.GetIconGeometry("SharedPipetteFill", true, out isFilled);
                if (lineReplaceGeom != null) PathLineColorReplace.Data = lineReplaceGeom;
            }
            catch { }
        }

        private void LoadData()
        {
            _currentProps = BulkFormatReplacerFeature.GetPropsInSelectedSlides(_manager);
            
            // Populate "As-Is" Combos
            ComboFontFind.Items.Clear();
            foreach (var f in _currentProps.Fonts.OrderBy(x => x)) ComboFontFind.Items.Add(f);
            
            ComboSizeFind.Items.Clear();
            foreach (var s in _currentProps.Sizes.OrderBy(x => x)) ComboSizeFind.Items.Add(s.ToString());
            
            ComboTextColorFind.Items.Clear();
            foreach (var c in _currentProps.TextColors) ComboTextColorFind.Items.Add(RgbToHex(c));
            
            ComboFillColorFind.Items.Clear();
            foreach (var c in _currentProps.FillColors) ComboFillColorFind.Items.Add(RgbToHex(c));
            
            ComboLineColorFind.Items.Clear();
            foreach (var c in _currentProps.LineColors) ComboLineColorFind.Items.Add(RgbToHex(c));

            // Populate "To-Be" Combos with common defaults
            PopulateToBeCombos();

            if (ComboFontFind.Items.Count > 0) ComboFontFind.SelectedIndex = 0;
            if (ComboSizeFind.Items.Count > 0) ComboSizeFind.SelectedIndex = 0;
        }

        private void PopulateToBeCombos()
        {
            // Fonts
            var fontNames = Fonts.SystemFontFamilies.Select(f => f.Source).OrderBy(s => s).ToList();
            if (fontNames.Count == 0) fontNames = new List<string> { "Arial", "Calibri", "Segoe UI", "Tahoma" };
            foreach (var name in fontNames) ComboFontReplace.Items.Add(name);

            // Sizes
            float[] commonSizes = { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96 };
            foreach (var s in commonSizes) ComboSizeReplace.Items.Add(s.ToString());

            // Colors
            AddCommonHexColors(ComboTextColorReplace);
            AddCommonHexColors(ComboFillColorReplace);
            AddCommonHexColors(ComboLineColorReplace);
        }

        private void Replace_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string fontFind = CheckFont.IsChecked == true ? ComboFontFind.Text : null;
                string fontReplace = CheckFont.IsChecked == true ? ComboFontReplace.Text : null;

                float? sizeFind = null;
                float? sizeReplace = null;
                if (CheckSize.IsChecked == true && float.TryParse(ComboSizeFind.Text, out float sf) && float.TryParse(ComboSizeReplace.Text, out float sr))
                {
                    sizeFind = sf;
                    sizeReplace = sr;
                }

                int? textColorFind = CheckTextColor.IsChecked == true ? HexToRgb(ComboTextColorFind.Text) : (int?)null;
                int? textColorReplace = CheckTextColor.IsChecked == true ? HexToRgb(ComboTextColorReplace.Text) : (int?)null;

                int? fillColorFind = CheckFillColor.IsChecked == true ? HexToRgb(ComboFillColorFind.Text) : (int?)null;
                int? fillColorReplace = CheckFillColor.IsChecked == true ? HexToRgb(ComboFillColorReplace.Text) : (int?)null;

                int? lineColorFind = CheckLineColor.IsChecked == true ? HexToRgb(ComboLineColorFind.Text) : (int?)null;
                int? lineColorReplace = CheckLineColor.IsChecked == true ? HexToRgb(ComboLineColorReplace.Text) : (int?)null;

                // Determine Scope
                var scope = BulkFormatReplacerFeature.ReplacementScope.SelectedSlides;
                if (RadioScopeCurrent.IsChecked == true) scope = BulkFormatReplacerFeature.ReplacementScope.CurrentSlide;
                else if (RadioScopeEntire.IsChecked == true) scope = BulkFormatReplacerFeature.ReplacementScope.EntirePresentation;

                BulkFormatReplacerFeature.ReplaceFormats(_manager, 
                    fontFind, fontReplace, 
                    sizeFind, sizeReplace, 
                    textColorFind, textColorReplace, 
                    fillColorFind, fillColorReplace, 
                    lineColorFind, lineColorReplace,
                    scope);

                MessageBox.Show("Batch replacement completed successfully.");
                LoadData(); // Refresh list
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during replacement: " + ex.Message);
            }
        }

        private void Pick_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            string targetName = btn.Tag as string;
            ComboBox targetCombo = this.FindName(targetName) as ComboBox;

            Window overlay = new Window
            {
                WindowStyle = WindowStyle.None,
                AllowsTransparency = true,
                Background = new SolidColorBrush(Color.FromArgb(1, 255, 255, 255)),
                Cursor = Cursors.Cross,
                Topmost = true,
                Left = SystemParameters.VirtualScreenLeft,
                Top = SystemParameters.VirtualScreenTop,
                Width = SystemParameters.VirtualScreenWidth,
                Height = SystemParameters.VirtualScreenHeight,
                ShowInTaskbar = false
            };

            overlay.MouseDown += (s, ev) =>
            {
                try
                {
                    var point = System.Windows.Forms.Control.MousePosition;
                    using (var bmp = new System.Drawing.Bitmap(1, 1))
                    {
                        using (var g = System.Drawing.Graphics.FromImage(bmp))
                        {
                            g.CopyFromScreen(point, new System.Drawing.Point(0, 0), new System.Drawing.Size(1, 1));
                        }
                        var color = bmp.GetPixel(0, 0);
                        int rgb = color.R | (color.G << 8) | (color.B << 16);
                        targetCombo.Text = RgbToHex(rgb);
                    }
                }
                catch { }
                overlay.Close();
            };

            overlay.KeyDown += (s, ev) => { if (ev.Key == Key.Escape) overlay.Close(); };
            overlay.ShowDialog();
        }

        private void Assign_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            string targetName = btn.Tag as string;
            string propertyType = btn.Uid as string;
            ComboBox targetCombo = this.FindName(targetName) as ComboBox;

            try
            {
                var selection = _manager.GetApplication().ActiveWindow.Selection;
                Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = null;
                
                if (selection != null && selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    try { if (selection.HasChildShapeRange) shapes = selection.ChildShapeRange; } catch { }
                    if (shapes == null) shapes = selection.ShapeRange;
                }

                if (shapes == null || shapes.Count == 0)
                {
                    MessageBox.Show("Please select a shape first.");
                    return;
                }

                string value = null;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    value = GetPropertyValueFromShape(shapes[i], propertyType);
                    if (value != null) break;
                }

                if (value != null) targetCombo.Text = value;
                else MessageBox.Show("Selection doesn't have the requested property (ensure it contains text for Font/Size/TextColor).");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not get property from selection: " + ex.Message);
            }
        }

        private string GetPropertyValueFromShape(Microsoft.Office.Interop.PowerPoint.Shape shape, string propertyType)
        {
            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape child in shape.GroupItems)
                {
                    string val = GetPropertyValueFromShape(child, propertyType);
                    if (val != null) return val;
                }
                return null;
            }

            try
            {
                switch (propertyType)
                {
                    case "Font":
                        if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                            return shape.TextFrame.TextRange.Font.Name;
                        break;
                    case "Size":
                        if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                            return shape.TextFrame.TextRange.Font.Size.ToString();
                        break;
                    case "TextColor":
                        if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                            return RgbToHex(shape.TextFrame.TextRange.Font.Color.RGB);
                        break;
                    case "FillColor":
                        if (shape.Fill.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                            return RgbToHex(shape.Fill.ForeColor.RGB);
                        break;
                    case "LineColor":
                        if (shape.Line.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                            return RgbToHex(shape.Line.ForeColor.RGB);
                        break;
                }
            }
            catch { }
            return null;
        }

        private void Close_Click(object sender, RoutedEventArgs e) => this.Close();

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) this.DragMove();
        }

        private string RgbToHex(int rgb)
        {
            byte r = (byte)(rgb & 0xFF);
            byte g = (byte)((rgb >> 8) & 0xFF);
            byte b = (byte)((rgb >> 16) & 0xFF);
            return $"#{r:X2}{g:X2}{b:X2}";
        }

        private int? HexToRgb(string hex)
        {
            if (string.IsNullOrWhiteSpace(hex)) return null;
            try
            {
                if (hex.StartsWith("#")) hex = hex.Substring(1);
                if (hex.Length == 6)
                {
                    byte r = Convert.ToByte(hex.Substring(0, 2), 16);
                    byte g = Convert.ToByte(hex.Substring(2, 2), 16);
                    byte b = Convert.ToByte(hex.Substring(4, 2), 16);
                    return (b << 16) | (g << 8) | r;
                }
            }
            catch { }
            return null;
        }

        private void AddCommonHexColors(ComboBox combo)
        {
            string[] colors = { "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#00FFFF", "#FF00FF", "#C0C0C0", "#808080" };
            foreach (var c in colors) combo.Items.Add(c);
        }
    }
}
