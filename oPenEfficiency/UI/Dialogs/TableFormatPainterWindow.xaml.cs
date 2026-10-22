using System;
using System.Windows;
using oPenEfficiency.Features;

namespace oPenEfficiency.UI
{
    public partial class TableFormatPainterWindow : Window
    {
        private readonly PowerPointManager _manager;

        public TableFormatPainterWindow() : this(null) { }

        public TableFormatPainterWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            bool up = ChkUp.IsChecked == true;
            bool down = ChkDown.IsChecked == true;
            bool left = ChkLeft.IsChecked == true;
            bool right = ChkRight.IsChecked == true;
            bool text = ChkText.IsChecked == true;
            bool cell = ChkCell.IsChecked == true;

            if (!up && !down && !left && !right)
            {
                MessageBox.Show(this, "Choose at least one paint direction.", "Table Format Painter", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (!text && !cell)
            {
                MessageBox.Show(this, "Choose Text Formatting, Cell Fill, or both.", "Table Format Painter", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int skip;
            if (!int.TryParse(TxtSkip.Text, out skip) || skip < 0)
            {
                MessageBox.Show(this, "Enter a whole number of 0 or greater for cells to skip.", "Table Format Painter", MessageBoxButton.OK, MessageBoxImage.Information);
                TxtSkip.Focus();
                return;
            }

            if (TableFormatPainterFeature.Paint(_manager, up, down, left, right, text, cell, skip))
            {
                Close();
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left) DragMove();
        }
    }
}
