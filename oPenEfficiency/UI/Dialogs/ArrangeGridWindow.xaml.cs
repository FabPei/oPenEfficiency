using System;
using System.Windows;
using System.Windows.Input;
using System.Text.RegularExpressions;
using oPenEfficiency.Features;

namespace oPenEfficiency.UI
{
    public partial class ArrangeGridWindow : Window
    {
        private PowerPointManager _manager;

        public ArrangeGridWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        private void BtnArrange_Click(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(TxtRows.Text, out int rows) && 
                int.TryParse(TxtColumns.Text, out int cols) && 
                float.TryParse(TxtSpacing.Text.Replace(",", "."), out float spacing))
            {
                ArrangeGridFeature.Execute(_manager, rows, cols, spacing);
                // Window stays open as requested
            }
            else
            {
                MessageBox.Show("Please enter valid numbers for rows, columns, and spacing.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
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
    }
}
