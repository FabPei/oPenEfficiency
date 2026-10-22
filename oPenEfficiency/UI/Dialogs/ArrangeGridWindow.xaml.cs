using System;
using System.Globalization;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Input;
using System.Windows.Media;
using System.Text.RegularExpressions;
using oPenEfficiency.Features;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    public partial class ArrangeGridWindow : Window
    {
        private PowerPointManager _manager;
        private int? _selectedShapeCount;

        public ArrangeGridWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            RefreshSelectedShapeCount();
            UpdatePreview();
            ThemeManager.ThemeChanged += ThemeChanged_RefreshPreview;
            Activated += ArrangeGridWindow_Activated;
            Closed += (sender, args) =>
            {
                ThemeManager.ThemeChanged -= ThemeChanged_RefreshPreview;
                Activated -= ArrangeGridWindow_Activated;
            };
            Loaded += (sender, args) =>
            {
                TxtColumns.Focus();
                TxtColumns.SelectAll();
                UpdatePrimaryActionState();
            };
        }

        private void ThemeChanged_RefreshPreview(object sender, string themeName)
        {
            UpdatePreview();
        }

        // The window is modeless, so the slide selection can change while it stays
        // open; refresh the cached count when the user comes back to the window.
        private void ArrangeGridWindow_Activated(object sender, EventArgs e)
        {
            RefreshSelectedShapeCount();
            UpdatePreview();
        }

        private void RefreshSelectedShapeCount()
        {
            try
            {
                var shapes = _manager?.GetSelectedShapes();
                _selectedShapeCount = shapes == null ? 0 : shapes.Count;
            }
            catch (Exception ex)
            {
                _selectedShapeCount = null;
                ExceptionLogger.Log(ex, "ArrangeGridWindow.RefreshSelectedShapeCount");
            }
        }

        private void DragSurface_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed) DragMove();
        }

        private void BtnArrange_Click(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(TxtRows.Text, out int rows) && rows > 0 &&
                int.TryParse(TxtColumns.Text, out int cols) && cols > 0 &&
                float.TryParse(TxtSpacing.Text.Replace(",", "."), NumberStyles.Float, CultureInfo.InvariantCulture, out float spacing) && spacing >= 0)
            {
                var shapes = _manager?.GetSelectedShapes();
                int shapeCount = shapes?.Count ?? 0;
                if ((long)rows * cols < shapeCount)
                {
                    MessageBox.Show($"The grid has space for {rows * cols} shape(s), but {shapeCount} are selected. Increase rows or columns.", "Grid Too Small", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (!ArrangeGridFeature.Execute(_manager, rows, cols, spacing))
                {
                    MessageBox.Show("Select at least two shapes and use positive rows and columns with a non-negative spacing.", "Unable to Arrange Grid", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            else
            {
                MessageBox.Show("Enter positive rows and columns and a non-negative spacing.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        // Numeric Control Handlers
        private void BtnIncCols_Click(object sender, RoutedEventArgs e) => AdjustValue(TxtColumns, 1);
        private void BtnDecCols_Click(object sender, RoutedEventArgs e) => AdjustValue(TxtColumns, -1);
        private void BtnIncRows_Click(object sender, RoutedEventArgs e) => AdjustValue(TxtRows, 1);
        private void BtnDecRows_Click(object sender, RoutedEventArgs e) => AdjustValue(TxtRows, -1);

        private void AdjustValue(System.Windows.Controls.TextBox tb, int delta)
        {
            if (int.TryParse(tb.Text, out int val))
            {
                int newVal = Math.Max(1, val + delta);
                tb.Text = newVal.ToString();
            }
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private void DecimalValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9,.]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private bool InputsAreValid()
        {
            return int.TryParse(TxtRows?.Text, out int rows) && rows > 0 &&
                   int.TryParse(TxtColumns?.Text, out int columns) && columns > 0 &&
                   float.TryParse((TxtSpacing?.Text ?? string.Empty).Replace(",", "."),
                       NumberStyles.Float, CultureInfo.InvariantCulture, out float spacing) && spacing >= 0;
        }

        private void Input_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdatePrimaryActionState();
            UpdatePreview();
        }

        private void UpdatePrimaryActionState()
        {
            if (BtnArrange != null) BtnArrange.IsEnabled = InputsAreValid();
        }

        private void UpdatePreview()
        {
            if (PreviewCells == null || PreviewDimensionLabel == null || PreviewHost == null) return;
            GridPreviewState state = GridPreviewPlanner.Create(TxtColumns?.Text, TxtRows?.Text, _selectedShapeCount);
            PreviewCells.Children.Clear();
            if (state.IsValid)
            {
                PreviewCells.Columns = state.VisibleColumns;
                PreviewCells.Rows = state.VisibleRows;
                Brush cellBrush = TryFindResource("AccentBrush") as Brush;
                int arrangedVisibleCells = Math.Min(state.ArrangedCellCount, state.CellCount);
                for (int i = 0; i < state.CellCount; i++)
                {
                    PreviewCells.Children.Add(new System.Windows.Controls.Border
                    {
                        Background = cellBrush,
                        Opacity = i < arrangedVisibleCells ? 0.65 : 0.18,
                        CornerRadius = new CornerRadius(1),
                        Margin = new Thickness(1)
                    });
                }
            }
            PreviewDimensionLabel.Text = state.DimensionLabel;
            AutomationProperties.SetName(PreviewHost, state.AutomationName);
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                e.Handled = true;
                Close();
            }
            else if (e.Key == Key.Enter && InputsAreValid())
            {
                e.Handled = true;
                BtnArrange_Click(BtnArrange, new RoutedEventArgs());
            }
        }
    }
}
