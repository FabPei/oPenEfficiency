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

            int skip = 0;
            int.TryParse(TxtSkip.Text, out skip);

            if (TableFormatPainterFeature.Paint(_manager, up, down, left, right, text, cell, skip))
            {
                Close();
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
