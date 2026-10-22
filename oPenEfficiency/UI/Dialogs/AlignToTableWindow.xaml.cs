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

            Loaded += (s, e) => PositionNearSelection();
        }

        private void PositionNearSelection()
        {
            try
            {
                _powerPointManager.GetSelectionScreenLocation(out int x, out int y);

                double dpiScaleX = 1.0;
                double dpiScaleY = 1.0;
                var source = PresentationSource.FromVisual(this);
                if (source?.CompositionTarget != null)
                {
                    dpiScaleX = source.CompositionTarget.TransformToDevice.M11;
                    dpiScaleY = source.CompositionTarget.TransformToDevice.M22;
                }

                Left = Math.Max(SystemParameters.WorkArea.Left, Math.Min((x / dpiScaleX) - (ActualWidth / 2), SystemParameters.WorkArea.Right - ActualWidth));
                Top = Math.Max(SystemParameters.WorkArea.Top, Math.Min((y / dpiScaleY) - (ActualHeight / 2), SystemParameters.WorkArea.Bottom - ActualHeight));
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "AlignToTableWindow.PositionNearSelection"); }
        }

        private void Align_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as System.Windows.Controls.Button;
            if (btn != null && btn.Tag != null)
            {
                try
                {
                    if (!oPenEfficiency.Features.AlignToTableCellFeature.Align(_powerPointManager, btn.Tag.ToString()))
                        MessageBox.Show(this, "No selected shape center is inside a visible table cell.", "Align to Cell", MessageBoxButton.OK, MessageBoxImage.Information);
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

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
                e.Handled = true;
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && e.ButtonState == MouseButtonState.Pressed)
                this.DragMove();
        }
    }
}
