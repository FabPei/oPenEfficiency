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

        public BulkFormatReplacerWindow(
            PowerPointManager manager,
            BulkFormatReplacerFeature.ReplacementScope? initialScope = null)
        {
            InitializeComponent();
            _manager = manager;
            if (initialScope.HasValue)
            {
                RadioScopeCurrent.IsChecked = initialScope.Value == BulkFormatReplacerFeature.ReplacementScope.CurrentSlide;
                RadioScopeSelected.IsChecked = initialScope.Value == BulkFormatReplacerFeature.ReplacementScope.SelectedSlides;
                RadioScopeEntire.IsChecked = initialScope.Value == BulkFormatReplacerFeature.ReplacementScope.EntirePresentation;
            }
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

                var textReplaceGeom = IconService.GetIconGeometry("SharedPipetteText", out isFilled);
                if (textReplaceGeom != null) PathTextColorReplace.Data = textReplaceGeom;

                // Fill Pipettes
                var fillFindGeom = IconService.GetIconGeometry("SharedPipetteFill", out isFilled);
                if (fillFindGeom != null) PathFillColorFind.Data = fillFindGeom;

                var fillReplaceGeom = IconService.GetIconGeometry("SharedPipetteFill", out isFilled);
                if (fillReplaceGeom != null) PathFillColorReplace.Data = fillReplaceGeom;

                // Line Pipettes (Reuse Fill for now or use the generic one)
                var lineFindGeom = IconService.GetIconGeometry("SharedPipetteFill", out isFilled);
                if (lineFindGeom != null) PathLineColorFind.Data = lineFindGeom;

                var lineReplaceGeom = IconService.GetIconGeometry("SharedPipetteFill", out isFilled);
                if (lineReplaceGeom != null) PathLineColorReplace.Data = lineReplaceGeom;
            }
            catch { /* The dialog can still open when initial selection properties are unavailable. */ }
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
                if (CheckFont.IsChecked != true &&
                    CheckSize.IsChecked != true &&
                    CheckTextColor.IsChecked != true &&
                    CheckFillColor.IsChecked != true &&
                    CheckLineColor.IsChecked != true)
                {
                    MessageBox.Show(this, "Select at least one format property to replace.", "Replace Formatting", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                string fontFind = CheckFont.IsChecked == true ? ComboFontFind.Text : null;
                string fontReplace = CheckFont.IsChecked == true ? ComboFontReplace.Text : null;
                if (CheckFont.IsChecked == true &&
                    (string.IsNullOrWhiteSpace(fontFind) || string.IsNullOrWhiteSpace(fontReplace)))
                {
                    MessageBox.Show(this, "Enter both font names.", "Replace Formatting", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                float? sizeFind = null;
                float? sizeReplace = null;
                if (CheckSize.IsChecked == true)
                {
                    if (!float.TryParse(ComboSizeFind.Text, out float sf) ||
                        !float.TryParse(ComboSizeReplace.Text, out float sr) ||
                        sf <= 0 || sr <= 0)
                    {
                        MessageBox.Show(this, "Enter two positive font sizes.", "Replace Formatting", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    sizeFind = sf;
                    sizeReplace = sr;
                }

                int? textColorFind = CheckTextColor.IsChecked == true ? HexToRgb(ComboTextColorFind.Text) : (int?)null;
                int? textColorReplace = CheckTextColor.IsChecked == true ? HexToRgb(ComboTextColorReplace.Text) : (int?)null;

                int? fillColorFind = CheckFillColor.IsChecked == true ? HexToRgb(ComboFillColorFind.Text) : (int?)null;
                int? fillColorReplace = CheckFillColor.IsChecked == true ? HexToRgb(ComboFillColorReplace.Text) : (int?)null;

                int? lineColorFind = CheckLineColor.IsChecked == true ? HexToRgb(ComboLineColorFind.Text) : (int?)null;
                int? lineColorReplace = CheckLineColor.IsChecked == true ? HexToRgb(ComboLineColorReplace.Text) : (int?)null;

                if ((CheckTextColor.IsChecked == true && (!textColorFind.HasValue || !textColorReplace.HasValue)) ||
                    (CheckFillColor.IsChecked == true && (!fillColorFind.HasValue || !fillColorReplace.HasValue)) ||
                    (CheckLineColor.IsChecked == true && (!lineColorFind.HasValue || !lineColorReplace.HasValue)))
                {
                    MessageBox.Show(this, "Enter colors as six-digit hex values, for example #1A2B3C.", "Replace Formatting", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Determine Scope
                var scope = BulkFormatReplacerFeature.ReplacementScope.SelectedSlides;
                if (RadioScopeCurrent.IsChecked == true) scope = BulkFormatReplacerFeature.ReplacementScope.CurrentSlide;
                else if (RadioScopeEntire.IsChecked == true) scope = BulkFormatReplacerFeature.ReplacementScope.EntirePresentation;

                int replacementCount = BulkFormatReplacerFeature.ReplaceFormats(_manager,
                    fontFind, fontReplace, 
                    sizeFind, sizeReplace, 
                    textColorFind, textColorReplace, 
                    fillColorFind, fillColorReplace, 
                    lineColorFind, lineColorReplace,
                    scope);

                if (replacementCount < 0)
                {
                    MessageBox.Show(this, "The replacement failed and completed changes were rolled back. See the openEfficiency error log for details.", "Replace Formatting", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                if (replacementCount == 0)
                {
                    MessageBox.Show(this, "No matching format properties were found in the selected scope.", "Replace Formatting", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                MessageBox.Show(this, $"Replaced {replacementCount} matching format properties.", "Replace Formatting", MessageBoxButton.OK, MessageBoxImage.Information);
                LoadData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error during replacement: " + ex.Message, "Replace Formatting", MessageBoxButton.OK, MessageBoxImage.Error);
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
                catch { /* Overlay cleanup continues when a preview shape is already gone. */ }
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
                var shapes = _manager.GetSelectedShapes();
                if (shapes == null || shapes.Count == 0)
                {
                    MessageBox.Show("Please select a shape first.");
                    return;
                }

                var shape = shapes[1];
                string value = null;

                switch (propertyType)
                {
                    case "Font":
                        if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            value = shape.TextFrame.TextRange.Font.Name;
                        break;
                    case "Size":
                        if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            value = shape.TextFrame.TextRange.Font.Size.ToString();
                        break;
                    case "TextColor":
                        if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            value = RgbToHex(shape.TextFrame.TextRange.Font.Color.RGB);
                        break;
                    case "FillColor":
                        if (shape.Fill.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                            value = RgbToHex(shape.Fill.ForeColor.RGB);
                        break;
                    case "LineColor":
                        if (shape.Line.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                            value = RgbToHex(shape.Line.ForeColor.RGB);
                        break;
                }

                if (value != null) targetCombo.Text = value;
                else MessageBox.Show("Selection doesn't have the requested property (ensure it contains text for Font/Size/TextColor).");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not get property from selection: " + ex.Message);
            }
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
            catch { /* A stale shape is treated as no match while locating the replacement target. */ }
            return null;
        }

        private void AddCommonHexColors(ComboBox combo)
        {
            string[] colors = { "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#00FFFF", "#FF00FF", "#C0C0C0", "#808080" };
            foreach (var c in colors) combo.Items.Add(c);
        }
    }
}
