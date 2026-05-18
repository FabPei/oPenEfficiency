using System;
using System.Windows;
using System.Windows.Input;
using oPenEfficiency.Features;

namespace oPenEfficiency.UI
{
    public partial class QRCodeWindow : Window
    {
        private PowerPointManager _manager;

        public QRCodeWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            TxtUrl.Focus();
            TxtUrl.SelectAll();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            string url = TxtUrl.Text.Trim();
            if (string.IsNullOrEmpty(url))
            {
                MessageBox.Show("Please enter a valid URL or text.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                // Generate QR Code URL (using api.qrserver.com)
                string qrCodeApiUrl = $"https://api.qrserver.com/v1/create-qr-code/?size=300x300&data={Uri.EscapeDataString(url)}";
                bool success = QRCodeFeature.InsertImageFromUrl(_manager, qrCodeApiUrl);
                
                if (success)
                {
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Failed to generate or insert the QR code. Please check your internet connection and try again.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error generating QR code: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TxtUrl_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                BtnOk_Click(this, new RoutedEventArgs());
            }
        }
    }
}
