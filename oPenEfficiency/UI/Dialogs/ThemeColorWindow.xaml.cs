using System;
using System.Windows;
using System.Windows.Input;

namespace oPenEfficiency.UI
{
    public partial class ThemeColorWindow : Window
    {
        public ThemeColorWindow(PowerPointManager manager)
        {
            InitializeComponent();
            ColorControl.Initialize(manager);
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed) DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
