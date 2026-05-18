using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using oPenEfficiency.Utils;
using oPenEfficiency.Features;

namespace oPenEfficiency.UI
{
    public partial class FormatNumbersWindow : Window
    {
        private readonly PowerPointManager _powerPointManager;

        public FormatNumbersWindow(PowerPointManager powerPointManager)
        {
            InitializeComponent();
            _powerPointManager = powerPointManager;

            // Center on screen and ensure window appears in front of PowerPoint
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.Topmost = true;
        }

        private void Format_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as System.Windows.Controls.Button;
            if (btn == null) return;

            string tag = btn.Tag.ToString();
            string tSep = ",";
            string dSep = ".";
            string groupingType = "Standard";

            switch (tag)
            {
                case "COMMA":
                    tSep = ","; dSep = "."; break;
                case "PERIOD":
                    tSep = "."; dSep = ","; break;
                case "SPACE":
                    tSep = " "; dSep = ","; break;
                case "APOSTROPHE":
                    tSep = "'"; dSep = "."; break;
                case "INDIAN":
                    tSep = ","; dSep = "."; groupingType = "Indian"; break;
            }

            string currencyPlacement = "InPlace";
            var radioPrefix = this.FindName("RadioPrefix") as RadioButton;
            var radioSuffix = this.FindName("RadioSuffix") as RadioButton;
            var radioRemove = this.FindName("RadioRemove") as RadioButton;

            if (radioPrefix != null && radioPrefix.IsChecked == true) currencyPlacement = "Prefix";
            else if (radioSuffix != null && radioSuffix.IsChecked == true) currencyPlacement = "Suffix";
            else if (radioRemove != null && radioRemove.IsChecked == true) currencyPlacement = "Remove";

            string targetCurrency = "";
            var txtCurrency = this.FindName("TxtCurrency") as TextBox;
            if (txtCurrency != null) targetCurrency = txtCurrency.Text;

            try
            {
                FormatNumbersFeature.Execute(_powerPointManager, tSep, dSep, groupingType, currencyPlacement, targetCurrency);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reformatting numbers: " + ex.Message);
            }
        }

        private void Placement_Changed(object sender, RoutedEventArgs e)
        {
            var txtPlacementInfo = this.FindName("TxtPlacementInfo") as TextBlock;
            if (txtPlacementInfo == null) return;

            var rb = sender as System.Windows.Controls.RadioButton;
            if (rb == null) return;

            switch (rb.Name)
            {
                case "RadioPrefix":
                    txtPlacementInfo.Text = "Common in English-speaking countries ($100.00).";
                    break;
                case "RadioSuffix":
                    txtPlacementInfo.Text = "Common in many European and African countries (100,00 € or 100 kr).";
                    break;
                case "RadioInPlace":
                    txtPlacementInfo.Text = "Specific contexts (e.g. Cape Verdean Escudo): the symbol acts as decimal separator (100$00).";
                    break;
                case "RadioRemove":
                    txtPlacementInfo.Text = "Strips all detected currency symbols from the numbers.";
                    break;
            }
        }

        private void QuickCurrency_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as System.Windows.Controls.Button;
            var txtCurrency = this.FindName("TxtCurrency") as TextBox;
            if (btn != null && txtCurrency != null)
            {
                txtCurrency.Text = btn.Content.ToString();
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
