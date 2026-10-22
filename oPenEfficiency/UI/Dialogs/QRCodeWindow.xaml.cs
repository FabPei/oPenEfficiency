using System;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using oPenEfficiency.Features;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    public partial class QRCodeWindow : Window
    {
        private PowerPointManager _manager;
        private CancellationTokenSource _generationCancellation;

        public QRCodeWindow(PowerPointManager manager, string initialText = null)
        {
            InitializeComponent();
            _manager = manager;
            SourceInitialized += (sender, args) => FitToOwnerWorkArea();
            if (!string.IsNullOrWhiteSpace(initialText)) TxtUrl.Text = initialText;
            TxtUrl.Focus();
            TxtUrl.SelectAll();
        }

        public void SetInitialText(string text)
        {
            TxtUrl.Text = text ?? string.Empty;
            TxtUrl.Focus();
            TxtUrl.SelectAll();
        }

        private void FitToOwnerWorkArea()
        {
            var helper = new System.Windows.Interop.WindowInteropHelper(this);
            var screen = System.Windows.Forms.Screen.FromHandle(helper.Owner);
            var dpi = System.Windows.Media.VisualTreeHelper.GetDpi(this);
            MaxWidth = Math.Max(1, (screen.WorkingArea.Width / dpi.DpiScaleX) - 32);
            MaxHeight = Math.Max(1, (screen.WorkingArea.Height / dpi.DpiScaleY) - 32);
            Width = Math.Min(Width, MaxWidth);
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                BtnCancel_Click(sender, e);
                e.Handled = true;
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && e.ButtonState == MouseButtonState.Pressed)
                this.DragMove();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            _generationCancellation?.Cancel();
            this.Close();
        }

        private async void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            string url = TxtUrl.Text.Trim();
            if (string.IsNullOrEmpty(url))
            {
                MessageBox.Show("Please enter a valid URL or text.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                string qrCodeApiUrl = $"https://api.qrserver.com/v1/create-qr-code/?size=300x300&data={Uri.EscapeDataString(url)}";
                BtnOk.IsEnabled = false;
                BtnOk.Content = "Generating...";
                _generationCancellation = new CancellationTokenSource();
                bool success = await QRCodeFeature.InsertImageFromUrlAsync(
                    _manager,
                    qrCodeApiUrl,
                    _generationCancellation.Token);
                
                if (success)
                {
                    this.Close();
                }
                else if (IsLoaded && !_generationCancellation.IsCancellationRequested)
                {
                    MessageBox.Show("Failed to generate or insert the QR code. Please check your internet connection and try again.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "QRCodeWindow.BtnOk_Click");
                MessageBox.Show("The QR code could not be generated or inserted.", "QR Code", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _generationCancellation?.Dispose();
                _generationCancellation = null;
                if (IsLoaded)
                {
                    BtnOk.IsEnabled = true;
                    BtnOk.Content = "Generate & Insert";
                }
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            _generationCancellation?.Cancel();
            base.OnClosed(e);
        }

        private void TxtUrl_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (BtnOk.IsEnabled) BtnOk_Click(this, new RoutedEventArgs());
                e.Handled = true;
            }
        }
    }
}
