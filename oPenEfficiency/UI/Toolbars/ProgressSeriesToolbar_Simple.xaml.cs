using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace oPenEfficiency.UI
{
    public partial class ProgressSeriesToolbarSimple : Window
    {
        public ProgressSeriesToolbarSimple()
        {
            InitializeComponent();
        }

        public void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }

        public void Input_KeyDown(object sender, KeyEventArgs e) { }
        public void Input_LostFocus(object sender, RoutedEventArgs e) { }
        public void SliderActive_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e) { }
        public void Setting_Changed(object sender, SelectionChangedEventArgs e) { }
        public void BtnOrientation_Click(object sender, RoutedEventArgs e) { }
        public void BtnColorWell_Click(object sender, RoutedEventArgs e) { }
        public void BtnReset_Click(object sender, RoutedEventArgs e) { }
    }
}
