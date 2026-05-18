using System;
using System.Windows;
using System.Windows.Input;
using oPenEfficiency.Features;

namespace oPenEfficiency.UI
{
    public partial class ObjectConnectorWindow : Window
    {
        private readonly PowerPointManager _manager;

        public ObjectConnectorWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) this.DragMove();
        }

        private void Window_Deactivated(object sender, EventArgs e)
        {
            try { this.Close(); } catch { }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnConnect_Click(object sender, RoutedEventArgs e)
        {
            string side1 = (ComboSide1.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Tag?.ToString();
            string side2 = (ComboSide2.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Tag?.ToString();
            bool auto = ChkAuto.IsChecked ?? false;
            bool smart = ChkSmart.IsChecked ?? false;

            if (ObjectConnectorFeature.Execute(_manager, side1, side2, auto, smart))
            {
                this.Close();
            }
        }
    }
}
