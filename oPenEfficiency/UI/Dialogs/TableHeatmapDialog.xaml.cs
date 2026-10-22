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
    public enum TableHeatmapScope
    {
        Row,
        Column,
        EntireTable
    }

    public partial class TableHeatmapDialog : Window
    {
        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left &&
                e.ButtonState == System.Windows.Input.MouseButtonState.Pressed)
                this.DragMove();
        }
        private void FitToOwnerWorkArea()
        {
            var helper = new System.Windows.Interop.WindowInteropHelper(this);
            var screen = Forms.Screen.FromHandle(helper.Owner);
            var dpi = VisualTreeHelper.GetDpi(this);
            MaxWidth = Math.Max(1, (screen.WorkingArea.Width / dpi.DpiScaleX) - 32);
            MaxHeight = Math.Max(1, (screen.WorkingArea.Height / dpi.DpiScaleY) - 32);
            Width = Math.Min(Width, MaxWidth);
        }

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape)
            {
                CancelButton_Click(sender, e);
                e.Handled = true;
            }
        }

        private PowerPointManager _manager;
        private Microsoft.Office.Interop.PowerPoint.Shape _tableShape;
        private Table _table;


        // Custom colors default to standard RGBs
        private System.Drawing.Color _lowColor = System.Drawing.Color.FromArgb(248, 113, 113); // Red
        private System.Drawing.Color? _midColor = System.Drawing.Color.FromArgb(251, 191, 36);  // Yellow
        private System.Drawing.Color _highColor = System.Drawing.Color.FromArgb(52, 211, 153); // Green

        public TableHeatmapDialog(
            PowerPointManager manager,
            Microsoft.Office.Interop.PowerPoint.Shape tableShape,
            Table table,
            TableHeatmapScope initialScope = TableHeatmapScope.Row)
        {
            InitializeComponent();
            _manager = manager;
            SourceInitialized += (sender, args) => FitToOwnerWorkArea();
            _tableShape = tableShape;
            _table = table;
            ScopeRowRadio.IsChecked = initialScope == TableHeatmapScope.Row;
            ScopeColRadio.IsChecked = initialScope == TableHeatmapScope.Column;
            ScopeTableRadio.IsChecked = initialScope == TableHeatmapScope.EntireTable;
            UpdatePreview();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ApplyHeatmap()) this.Close();
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
            UpdatePreview();
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
            UpdatePreview();
        }

        private void BtnLowColor_Click(object sender, RoutedEventArgs e)
        {
            var color = PickColor(_lowColor);
            if (color.HasValue)
            {
                _lowColor = color.Value;
                RectLowColor.Background = new SolidColorBrush(Color.FromRgb(_lowColor.R, _lowColor.G, _lowColor.B));
            }
            UpdatePreview();
        }

        private void BtnMidColor_Click(object sender, RoutedEventArgs e)
        {
            var color = PickColor(_midColor ?? System.Drawing.Color.White);
            if (color.HasValue)
            {
                _midColor = color.Value;
                RectMidColor.Background = new SolidColorBrush(Color.FromRgb(color.Value.R, color.Value.G, color.Value.B));
            }
            UpdatePreview();
        }

        private void BtnClearMidColor_Click(object sender, RoutedEventArgs e)
        {
            _midColor = null;
            RectMidColor.Background = Brushes.Transparent;
            UpdatePreview();
        }

        private void BtnHighColor_Click(object sender, RoutedEventArgs e)
        {
            var color = PickColor(_highColor);
            if (color.HasValue)
            {
                _highColor = color.Value;
                RectHighColor.Background = new SolidColorBrush(Color.FromRgb(_highColor.R, _highColor.G, _highColor.B));
            }
            UpdatePreview();
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

        // REG-TABLEHEATMAP-PREVIEW-001: sample data for the live preview. The top row doubles as
        // the header row demo for the Exclude-header option; the last cell is blank to show the
        // empty-cell option. Ranges differ per row so scope normalization is visible at a glance.
        private static readonly double?[,] PreviewValues =
        {
            { 100, 200, 300, 400 },
            { 10, 20, 30, 40 },
            { 5, 15, 25, 35 },
            { 60, 70, 80, null }
        };

        private void Scope_Changed(object sender, RoutedEventArgs e) => UpdatePreview();

        private void PreviewOption_Changed(object sender, RoutedEventArgs e) => UpdatePreview();

        /// <summary>
        /// Renders the sample grid with the exact color/icon math the applier applies to real cells
        /// (TableHeatmapApplier.ComputeColor / ComputeIcon), including the applier's grouping,
        /// header and empty-cell semantics, so the preview can never drift from the result.
        /// </summary>
        private void UpdatePreview()
        {
            if (PreviewGrid == null || ColorPresetsComboBox == null) return;

            var options = new oPenEfficiency.Features.Tables.TableHeatmapOptions
            {
                ApplyToFill = ApplyFillRadio.IsChecked == true,
                ApplyIcons = ApplyIconsRadio.IsChecked == true,
                ExcludeHeader = ExcludeHeaderCheckBox.IsChecked == true,
                IgnoreEmpty = IgnoreEmptyCheckBox.IsChecked == true,
                ColorPresetIndex = ColorPresetsComboBox.SelectedIndex,
                IconPresetIndex = IconPresetsComboBox.SelectedIndex,
                CustomLow = _lowColor,
                CustomMid = _midColor,
                CustomHigh = _highColor
            };

            int rows = PreviewValues.GetLength(0);
            int cols = PreviewValues.GetLength(1);
            int startRow = options.ExcludeHeader ? 1 : 0;

            // Mirror TableHeatmapApplier grouping: row, column, or entire-table min/max groups.
            var groups = new List<List<int[]>>();
            if (ScopeTableRadio.IsChecked == true)
            {
                var cells = new List<int[]>();
                for (int r = startRow; r < rows; r++) for (int c = 0; c < cols; c++) cells.Add(new[] { r, c });
                groups.Add(cells);
            }
            else if (ScopeColRadio.IsChecked == true)
            {
                for (int c = 0; c < cols; c++)
                {
                    var cells = new List<int[]>();
                    for (int r = startRow; r < rows; r++) cells.Add(new[] { r, c });
                    groups.Add(cells);
                }
            }
            else
            {
                for (int r = startRow; r < rows; r++)
                {
                    var cells = new List<int[]>();
                    for (int c = 0; c < cols; c++) cells.Add(new[] { r, c });
                    groups.Add(cells);
                }
            }

            double?[,] normalized = new double?[rows, cols];
            foreach (var group in groups)
            {
                double minimum = double.MaxValue;
                double maximum = double.MinValue;
                bool found = false;
                foreach (var cell in group)
                {
                    double? sample = PreviewValues[cell[0], cell[1]];
                    if (!sample.HasValue) continue;
                    minimum = Math.Min(minimum, sample.Value);
                    maximum = Math.Max(maximum, sample.Value);
                    found = true;
                }
                if (!found) continue;
                foreach (var cell in group)
                {
                    int r = cell[0];
                    int c = cell[1];
                    bool numeric = PreviewValues[r, c].HasValue;
                    if (!numeric && options.IgnoreEmpty) continue;
                    double value = numeric ? PreviewValues[r, c].Value : minimum;
                    normalized[r, c] = maximum > minimum ? (value - minimum) / (maximum - minimum) : 0;
                }
            }

            PreviewGrid.Children.Clear();
            PreviewGrid.RowDefinitions.Clear();
            PreviewGrid.ColumnDefinitions.Clear();
            for (int c = 0; c < cols; c++)
                PreviewGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            for (int r = 0; r < rows; r++)
                PreviewGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    bool headerCell = r == 0 && options.ExcludeHeader;
                    double? norm = normalized[r, c];
                    string sampleText = PreviewValues[r, c]?.ToString(CultureInfo.InvariantCulture) ?? "–";

                    var text = new TextBlock
                    {
                        Text = sampleText,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(0, 2, 0, 2)
                    };
                    var cell = new Border
                    {
                        Padding = new Thickness(6, 4, 6, 4),
                        Margin = new Thickness(1),
                        CornerRadius = new CornerRadius(2),
                        BorderBrush = (Brush)FindResource("BorderBrush"),
                        BorderThickness = new Thickness(1),
                        Child = text
                    };

                    if (!norm.HasValue || headerCell)
                    {
                        text.Foreground = (Brush)FindResource("TextSecondaryBrush");
                    }
                    else if (options.ApplyIcons)
                    {
                        oPenEfficiency.Features.Tables.TableHeatmapApplier.ComputeIcon(
                            norm.Value, options.IconPresetIndex, out string icon, out System.Drawing.Color iconColor);
                        var iconRun = new System.Windows.Documents.Run(icon + " ");
                        iconRun.Foreground = new SolidColorBrush(Color.FromRgb(iconColor.R, iconColor.G, iconColor.B));
                        text.Inlines.Clear();
                        text.Inlines.Add(iconRun);
                        text.Inlines.Add(new System.Windows.Documents.Run(sampleText));
                        text.Foreground = (Brush)FindResource("TextPrimaryBrush");
                    }
                    else
                    {
                        System.Drawing.Color finalColor = oPenEfficiency.Features.Tables.TableHeatmapApplier.ComputeColor(
                            norm.Value, options.ColorPresetIndex, options);
                        var brush = new SolidColorBrush(Color.FromRgb(finalColor.R, finalColor.G, finalColor.B));
                        if (options.ApplyToFill)
                        {
                            cell.Background = brush;
                            text.Foreground = ContrastBrush(finalColor);
                        }
                        else
                        {
                            text.Foreground = brush;
                        }
                    }

                    Grid.SetRow(cell, r);
                    Grid.SetColumn(cell, c);
                    PreviewGrid.Children.Add(cell);
                }
            }

            string scopeText = ScopeTableRadio.IsChecked == true ? "Entire table: one normalization for all values."
                : ScopeColRadio.IsChecked == true ? "Column-wise: each column normalized separately."
                : "Row-wise: each row normalized separately.";
            string emptyText = options.IgnoreEmpty ? " Blank cells are left unchanged." : " Blank cells receive the group's lowest color.";
            PreviewCaption.Text = scopeText + emptyText + " Real cells follow the same rules with your table's values.";
        }

        private static Brush ContrastBrush(System.Drawing.Color background)
        {
            double luminance = (0.299 * background.R + 0.587 * background.G + 0.114 * background.B) / 255;
            return luminance > 0.6 ? Brushes.Black : Brushes.White;
        }

        private bool ApplyHeatmap()
        {
            var options = new oPenEfficiency.Features.Tables.TableHeatmapOptions
            {
                Scope = ScopeTableRadio.IsChecked == true ? TableHeatmapScope.EntireTable
                    : ScopeColRadio.IsChecked == true ? TableHeatmapScope.Column
                    : TableHeatmapScope.Row,
                ApplyToFill = ApplyFillRadio.IsChecked == true,
                ApplyIcons = ApplyIconsRadio.IsChecked == true,
                ExcludeHeader = ExcludeHeaderCheckBox.IsChecked == true,
                IgnoreEmpty = IgnoreEmptyCheckBox.IsChecked == true,
                ColorPresetIndex = ColorPresetsComboBox.SelectedIndex,
                IconPresetIndex = IconPresetsComboBox.SelectedIndex,
                CustomLow = _lowColor,
                CustomMid = _midColor,
                CustomHigh = _highColor
            };

            // The application core (scoping, normalization, color and icon math) moved verbatim to
            // TableHeatmapApplier so the identical behavior runs for this dialog and for the
            // dialog-free automation entry; every user-facing message stays here.
            var result = oPenEfficiency.Features.Tables.TableHeatmapApplier.Apply(_manager, _table, options);
            if (result.NoDataRows)
            {
                MessageBox.Show(this, "No data rows remain after excluding the header.", "Apply Table Heatmap", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }
            if (result.NoNumericValues)
            {
                MessageBox.Show(this, "No numeric values were found in the selected heatmap scope.", "Apply Table Heatmap", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }
            if (result.Failed > 0)
            {
                MessageBox.Show(this, $"Formatted {result.Formatted} cell(s), but {result.Failed} cell(s) failed. Use PowerPoint Undo to restore the previous formatting before retrying.", "Apply Table Heatmap", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            return result.Formatted > 0;
        }
    }
}
