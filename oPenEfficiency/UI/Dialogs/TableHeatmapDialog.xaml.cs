using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using Forms = System.Windows.Forms;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class TableHeatmapDialog : Window
    {
        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                this.DragMove();
        }
        private PowerPointManager _manager;
        private Microsoft.Office.Interop.PowerPoint.Shape _tableShape;
        private Table _table;

        // Custom colors default to standard RGBs
        private System.Drawing.Color _lowColor = System.Drawing.Color.FromArgb(248, 113, 113); // Red
        private System.Drawing.Color? _midColor = System.Drawing.Color.FromArgb(251, 191, 36);  // Yellow
        private System.Drawing.Color _highColor = System.Drawing.Color.FromArgb(52, 211, 153); // Green

        public TableHeatmapDialog(PowerPointManager manager, Microsoft.Office.Interop.PowerPoint.Shape tableShape, Table table)
        {
            InitializeComponent();
            _manager = manager;
            _tableShape = tableShape;
            _table = table;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ApplyHeatmap();
                this.Close();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableHeatmapDialog.ApplyButton_Click");
                Forms.MessageBox.Show($"Error applying heatmap: {ex.Message}", "Error", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Error);
            }
        }

        private void ColorPresetsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CustomColorPanel == null) return;

            if (ColorPresetsComboBox.SelectedIndex == 12) // "Custom..."
            {
                CustomColorPanel.Visibility = Visibility.Visible;
            }
            else
            {
                CustomColorPanel.Visibility = Visibility.Collapsed;
            }
        }

        private void ApplyType_Changed(object sender, RoutedEventArgs e)
        {
            if (ColorPresetsComboBox == null || IconPresetsComboBox == null) return;
            bool isIconType = ApplyIconsRadio.IsChecked == true;
            ColorPresetsComboBox.Visibility = isIconType ? Visibility.Collapsed : Visibility.Visible;
            IconPresetsComboBox.Visibility = isIconType ? Visibility.Visible : Visibility.Collapsed;
            
            if (isIconType)
            {
                CustomColorPanel.Visibility = Visibility.Collapsed;
            }
            else
            {
                ColorPresetsComboBox_SelectionChanged(null, null);
            }
        }

        private void BtnLowColor_Click(object sender, RoutedEventArgs e)
        {
            var color = PickColor(_lowColor);
            if (color.HasValue)
            {
                _lowColor = color.Value;
                RectLowColor.Background = new SolidColorBrush(Color.FromRgb(_lowColor.R, _lowColor.G, _lowColor.B));
            }
        }

        private void BtnMidColor_Click(object sender, RoutedEventArgs e)
        {
            var color = PickColor(_midColor ?? System.Drawing.Color.White);
            if (color.HasValue)
            {
                _midColor = color.Value;
                RectMidColor.Background = new SolidColorBrush(Color.FromRgb(color.Value.R, color.Value.G, color.Value.B));
            }
        }

        private void BtnClearMidColor_Click(object sender, RoutedEventArgs e)
        {
            _midColor = null;
            RectMidColor.Background = Brushes.Transparent;
        }

        private void BtnHighColor_Click(object sender, RoutedEventArgs e)
        {
            var color = PickColor(_highColor);
            if (color.HasValue)
            {
                _highColor = color.Value;
                RectHighColor.Background = new SolidColorBrush(Color.FromRgb(_highColor.R, _highColor.G, _highColor.B));
            }
        }

        private System.Drawing.Color? PickColor(System.Drawing.Color initialColor)
        {
            using (var dialog = new Forms.ColorDialog())
            {
                dialog.Color = initialColor;
                dialog.FullOpen = true;
                if (dialog.ShowDialog() == Forms.DialogResult.OK)
                {
                    return dialog.Color;
                }
            }
            return null;
        }

        private void ApplyHeatmap()
        {
            // Gather options
            bool isRowScope = ScopeRowRadio.IsChecked == true;
            bool isColScope = ScopeColRadio.IsChecked == true;
            bool isTableScope = ScopeTableRadio.IsChecked == true;
            
            bool applyToFill = ApplyFillRadio.IsChecked == true;
            bool applyIcons = ApplyIconsRadio.IsChecked == true;
            bool excludeHeader = ExcludeHeaderCheckBox.IsChecked == true;
            bool ignoreEmpty = IgnoreEmptyCheckBox.IsChecked == true;

            int presetIndex = ColorPresetsComboBox.SelectedIndex;
            int iconPresetIndex = IconPresetsComboBox.SelectedIndex;

            System.Drawing.Color low, high;
            System.Drawing.Color? mid = null;

            System.Drawing.Color green = System.Drawing.Color.FromArgb(99, 190, 123);
            System.Drawing.Color yellow = System.Drawing.Color.FromArgb(255, 235, 132);
            System.Drawing.Color red = System.Drawing.Color.FromArgb(248, 105, 107);
            System.Drawing.Color white = System.Drawing.Color.FromArgb(255, 255, 255);
            System.Drawing.Color blue = System.Drawing.Color.FromArgb(90, 138, 198);

            switch (presetIndex)
            {
                case 0: // 3-Color: Red to Yellow to Green
                    low = red; mid = yellow; high = green; break;
                case 1: // 3-Color: Green to Yellow to Red
                    low = green; mid = yellow; high = red; break;
                case 2: // 3-Color: Red to White to Green
                    low = red; mid = white; high = green; break;
                case 3: // 3-Color: Green to White to Red
                    low = green; mid = white; high = red; break;
                case 4: // 3-Color: Red to White to Blue
                    low = red; mid = white; high = blue; break;
                case 5: // 3-Color: Blue to White to Red
                    low = blue; mid = white; high = red; break;
                case 6: // 2-Color: White to Red
                    low = white; high = red; break;
                case 7: // 2-Color: Red to White
                    low = red; high = white; break;
                case 8: // 2-Color: White to Green
                    low = white; high = green; break;
                case 9: // 2-Color: Green to White
                    low = green; high = white; break;
                case 10: // 2-Color: Yellow to Green
                    low = yellow; high = green; break;
                case 11: // 2-Color: Green to Yellow
                    low = green; high = yellow; break;
                case 12: // Custom
                default:
                    low = _lowColor;
                    mid = _midColor;
                    high = _highColor;
                    break;
            }

            int rows = _table.Rows.Count;
            int cols = _table.Columns.Count;
            int startRow = excludeHeader && rows > 1 ? 2 : 1;

            if (isTableScope)
            {
                // Find Global Min/Max
                double globalMin = double.MaxValue;
                double globalMax = double.MinValue;
                bool foundAny = false;

                for (int r = startRow; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        if (TryParseCell(_table.Cell(r, c), out double val))
                        {
                            if (val < globalMin) globalMin = val;
                            if (val > globalMax) globalMax = val;
                            foundAny = true;
                        }
                    }
                }

                if (foundAny)
                {
                    for (int r = startRow; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            ApplyFormattingToCell(_table.Cell(r, c), globalMin, globalMax, low, mid, high, applyToFill, applyIcons, iconPresetIndex, ignoreEmpty);
                        }
                    }
                }
            }
            else if (isRowScope)
            {
                for (int r = startRow; r <= rows; r++)
                {
                    double rowMin = double.MaxValue;
                    double rowMax = double.MinValue;
                    bool foundAny = false;

                    for (int c = 1; c <= cols; c++)
                    {
                        if (TryParseCell(_table.Cell(r, c), out double val))
                        {
                            if (val < rowMin) rowMin = val;
                            if (val > rowMax) rowMax = val;
                            foundAny = true;
                        }
                    }

                    if (foundAny)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            ApplyFormattingToCell(_table.Cell(r, c), rowMin, rowMax, low, mid, high, applyToFill, applyIcons, iconPresetIndex, ignoreEmpty);
                        }
                    }
                }
            }
            else if (isColScope)
            {
                for (int c = 1; c <= cols; c++)
                {
                    double colMin = double.MaxValue;
                    double colMax = double.MinValue;
                    bool foundAny = false;

                    for (int r = startRow; r <= rows; r++)
                    {
                        if (TryParseCell(_table.Cell(r, c), out double val))
                        {
                            if (val < colMin) colMin = val;
                            if (val > colMax) colMax = val;
                            foundAny = true;
                        }
                    }

                    if (foundAny)
                    {
                        for (int r = startRow; r <= rows; r++)
                        {
                            ApplyFormattingToCell(_table.Cell(r, c), colMin, colMax, low, mid, high, applyToFill, applyIcons, iconPresetIndex, ignoreEmpty);
                        }
                    }
                }
            }
        }

        private bool TryParseCell(Cell cell, out double value)
        {
            value = 0;
            try
            {
                var shape = cell.Shape;
                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    string text = shape.TextFrame.TextRange.Text.Trim();
                    // Clean up common formatting (currency, percent)
                    text = text.Replace("$", "").Replace("€", "").Replace("%", "").Trim();
                    // Try parsing with any number style in current culture or invariant
                    if (double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value)) return true;
                    if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value)) return true;
                }
            }
            catch { }
            return false;
        }

        private void ApplyFormattingToCell(Cell cell, double min, double max, System.Drawing.Color lowColor, System.Drawing.Color? midColor, System.Drawing.Color highColor, bool applyToFill, bool applyIcons, int iconPresetIndex, bool ignoreEmpty)
        {
            try
            {
                if (!TryParseCell(cell, out double value))
                {
                    if (ignoreEmpty) return;
                    value = min; // Default fallback if not ignored
                }

                double normalized = 0;
                if (max > min)
                {
                    normalized = (value - min) / (max - min);
                }

                if (applyIcons)
                {
                    ApplyIconToCell(cell, normalized, iconPresetIndex);
                    return;
                }

                System.Drawing.Color finalColor;

                if (midColor.HasValue)
                {
                    if (normalized < 0.5)
                    {
                        double n = normalized * 2; // scale 0-0.5 to 0-1
                        finalColor = InterpolateColor(lowColor, midColor.Value, n);
                    }
                    else
                    {
                        double n = (normalized - 0.5) * 2; // scale 0.5-1 to 0-1
                        finalColor = InterpolateColor(midColor.Value, highColor, n);
                    }
                }
                else
                {
                    finalColor = InterpolateColor(lowColor, highColor, normalized);
                }

                int bgr = System.Drawing.ColorTranslator.ToOle(finalColor);

                if (applyToFill)
                {
                    cell.Shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                    cell.Shape.Fill.ForeColor.RGB = bgr;
                }
                else
                {
                    if (cell.Shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && cell.Shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        cell.Shape.TextFrame.TextRange.Font.Color.RGB = bgr;
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableHeatmapDialog.ApplyColorToCell");
            }
        }

        private System.Drawing.Color InterpolateColor(System.Drawing.Color c1, System.Drawing.Color c2, double fraction)
        {
            fraction = Math.Max(0, Math.Min(1, fraction));
            int r = (int)(c1.R + (c2.R - c1.R) * fraction);
            int g = (int)(c1.G + (c2.G - c1.G) * fraction);
            int b = (int)(c1.B + (c2.B - c1.B) * fraction);
            return System.Drawing.Color.FromArgb(r, g, b);
        }

        private void ApplyIconToCell(Cell cell, double normalized, int presetIndex)
        {
            try
            {
                string icon = "";
                System.Drawing.Color color = System.Drawing.Color.Black;

                switch (presetIndex)
                {
                    case 0: // 3 Arrows (▲ ▶ ▼)
                        if (normalized >= 0.66) { icon = "▲"; color = System.Drawing.Color.FromArgb(99, 190, 123); } // Green
                        else if (normalized >= 0.33) { icon = "▶"; color = System.Drawing.Color.FromArgb(255, 235, 132); } // Yellow
                        else { icon = "▼"; color = System.Drawing.Color.FromArgb(248, 105, 107); } // Red
                        break;
                    case 1: // 4 Arrows (▲ ↗ ↘ ▼)
                        if (normalized >= 0.75) { icon = "▲"; color = System.Drawing.Color.FromArgb(99, 190, 123); } 
                        else if (normalized >= 0.50) { icon = "↗"; color = System.Drawing.Color.FromArgb(255, 235, 132); } 
                        else if (normalized >= 0.25) { icon = "↘"; color = System.Drawing.Color.FromArgb(248, 105, 107); } 
                        else { icon = "▼"; color = System.Drawing.Color.FromArgb(248, 105, 107); } 
                        break;
                    case 2: // 3 Traffic Lights (● ● ●)
                        if (normalized >= 0.66) { icon = "●"; color = System.Drawing.Color.FromArgb(99, 190, 123); }
                        else if (normalized >= 0.33) { icon = "●"; color = System.Drawing.Color.FromArgb(255, 235, 132); }
                        else { icon = "●"; color = System.Drawing.Color.FromArgb(248, 105, 107); }
                        break;
                    case 3: // 3 Indicators (✔ ❗ ✖)
                        if (normalized >= 0.66) { icon = "✔"; color = System.Drawing.Color.FromArgb(99, 190, 123); }
                        else if (normalized >= 0.33) { icon = "❗"; color = System.Drawing.Color.FromArgb(255, 235, 132); }
                        else { icon = "✖"; color = System.Drawing.Color.FromArgb(248, 105, 107); }
                        break;
                    case 4: // 3 Ratings (★)
                        if (normalized >= 0.66) { icon = "★"; color = System.Drawing.Color.FromArgb(255, 235, 132); }
                        else if (normalized >= 0.33) { icon = "★"; color = System.Drawing.Color.Silver; }
                        else { icon = "★"; color = System.Drawing.Color.SandyBrown; }
                        break;
                }

                if (cell.Shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && cell.Shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    string oldText = cell.Shape.TextFrame.TextRange.Text;
                    
                    // Clean existing icons
                    string[] knownIcons = { "▲", "▶", "▼", "↗", "↘", "●", "✔", "❗", "✖", "★", "☆" };
                    foreach (var k in knownIcons)
                    {
                        if (oldText.StartsWith(k + " ")) oldText = oldText.Substring(k.Length + 1);
                        else if (oldText.StartsWith(k)) oldText = oldText.Substring(k.Length);
                    }
                    
                    cell.Shape.TextFrame.TextRange.Text = icon + " " + oldText;
                    var firstChar = cell.Shape.TextFrame.TextRange.Characters(1, 1);
                    firstChar.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(color);
                    firstChar.Font.Name = "Segoe UI Symbol";
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableHeatmapDialog.ApplyIconToCell");
            }
        }
    }
}
