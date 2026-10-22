using System.Collections.Generic;
using System.Windows;

namespace oPenEfficiency.UI.Dialogs
{
    public enum SlideSizeLabelAction { None, CreateOrRefresh, Remove }

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

        public SlideSizeLabelAction LabelAction { get; private set; }
        public SlideSizeWindow(List<SlideSizeItem> items, bool hasExistingLabels)
        {
            InitializeComponent();
            DgSlideSizes.ItemsSource = items;
            BtnRemoveLabels.IsEnabled = hasExistingLabels;
        }

        private void BtnShowLabels_Click(object sender, RoutedEventArgs e)
        {
            LabelAction = SlideSizeLabelAction.CreateOrRefresh;
            this.DialogResult = true;
            this.Close();
        }

        private void BtnRemoveLabels_Click(object sender, RoutedEventArgs e)
        {
            LabelAction = SlideSizeLabelAction.Remove;
            DialogResult = true;
            Close();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
