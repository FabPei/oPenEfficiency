using System.Collections.Generic;
using System.Windows;

namespace oPenEfficiency.UI.Dialogs
{
    public class SlideSizeItem
    {
        public int SlideNumber { get; set; }
        public double SizeMB { get; set; }
        public string Details { get; set; }
    }

    public partial class SlideSizeWindow : Window
    {
        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                this.DragMove();
        }

        public bool CreateLabelsRequested { get; private set; }
        public SlideSizeWindow(List<SlideSizeItem> items)
        {
            InitializeComponent();
            DgSlideSizes.ItemsSource = items;
        }

        private void BtnShowLabels_Click(object sender, RoutedEventArgs e)
        {
            CreateLabelsRequested = true;
            this.DialogResult = true;
            this.Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
