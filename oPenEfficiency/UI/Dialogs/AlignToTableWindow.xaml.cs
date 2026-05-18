using System;
using System.Windows;
using System.Windows.Input;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    public partial class AlignToTableWindow : Window
    {
        private readonly PowerPointManager _powerPointManager;

        public AlignToTableWindow(PowerPointManager powerPointManager)
        {
            InitializeComponent();
            _powerPointManager = powerPointManager;

            // Position near selection
            _powerPointManager.GetSelectionScreenLocation(out int x, out int y);

            // Convert physical pixels (PowerPoint) to logical pixels (WPF)
            double dpiScaleX = 1.0;
            double dpiScaleY = 1.0;
            var source = PresentationSource.FromVisual(this);
            if (source?.CompositionTarget != null)
            {
                dpiScaleX = source.CompositionTarget.TransformToDevice.M11;
                dpiScaleY = source.CompositionTarget.TransformToDevice.M22;
            }

            this.Left = (x / dpiScaleX) - (this.Width / 2);
            this.Top = (y / dpiScaleY) - (this.Height / 2);
        }

        private void Align_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as System.Windows.Controls.Button;
            if (btn != null && btn.Tag != null)
            {
                try
                {
                    oPenEfficiency.Features.AlignToTableCellFeature.Align(_powerPointManager, btn.Tag.ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error aligning shapes: " + ex.Message);
                }
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }
    }
}
