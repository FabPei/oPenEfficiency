using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using oPenEfficiency.Features;

namespace oPenEfficiency.UI
{
    public partial class TrafficLightMenu : Window
    {
        private PowerPointManager _manager;
        private TrafficLightFeature.TrafficLightState _currentState;
        private bool _isHorizontal;
        private bool _hasBorder;
        private float _borderWeight;

        private bool _isDialogOpened = false;

        public TrafficLightMenu(PowerPointManager manager, string altText)
        {
            InitializeComponent();
            _manager = manager;

            // Parse state from the latest persisted metadata.
            // Tag format: oPE_TrafficLight|stateInt|orientation(H/V)|hasBorder(0/1)|borderWeight
            var parts = altText.Split('|');
            _currentState = TrafficLightFeature.TrafficLightState.Red;
            _isHorizontal = false;
            _hasBorder = true;
            _borderWeight = 1.5f;

            if (parts.Length >= 5)
            {
                if (int.TryParse(parts[1], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int stateInt) &&
                    Enum.IsDefined(typeof(TrafficLightFeature.TrafficLightState), stateInt))
                    _currentState = (TrafficLightFeature.TrafficLightState)stateInt;
                _isHorizontal = parts[2] == "H";
                _hasBorder = parts[3] == "1";
                string normalizedWeight = parts[4].Replace(',', '.');
                if (float.TryParse(normalizedWeight, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out float weight))
                    _borderWeight = weight;
            }

            UpdateButtonStates();
            Loaded += (sender, args) => GetCurrentStateButton().Focus();
        }

        private void UpdateButtonStates()
        {
            // Dim all non-selected
            IconRed.Opacity = _currentState == TrafficLightFeature.TrafficLightState.Red ? 1.0 : 0.3;
            IconYellow.Opacity = _currentState == TrafficLightFeature.TrafficLightState.Yellow ? 1.0 : 0.3;
            IconGreen.Opacity = _currentState == TrafficLightFeature.TrafficLightState.Green ? 1.0 : 0.3;
            IconOff.Opacity = _currentState == TrafficLightFeature.TrafficLightState.Off ? 1.0 : 0.3;

            TxtOrientation.Text = _isHorizontal ? "Horiz." : "Vert.";
            TxtBorder.SetResourceReference(
                TextBlock.ForegroundProperty,
                _hasBorder ? "TextPrimaryBrush" : "TextMutedBrush");

            ComboWeight.Text = _borderWeight.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private void ExecuteUpdate()
        {
            try
            {
                TrafficLightFeature.Execute(_manager, _currentState, _isHorizontal, _hasBorder, _borderWeight);
            }
            catch { /* Initial selection probing is optional; commands validate the target again. */ }
        }

        private void BtnRed_Click(object sender, RoutedEventArgs e)
        {
            _currentState = TrafficLightFeature.TrafficLightState.Red;
            ExecuteUpdate();
            this.Close();
        }

        private void BtnYellow_Click(object sender, RoutedEventArgs e)
        {
            _currentState = TrafficLightFeature.TrafficLightState.Yellow;
            ExecuteUpdate();
            this.Close();
        }

        private void BtnGreen_Click(object sender, RoutedEventArgs e)
        {
            _currentState = TrafficLightFeature.TrafficLightState.Green;
            ExecuteUpdate();
            this.Close();
        }

        private void BtnOff_Click(object sender, RoutedEventArgs e)
        {
            _currentState = TrafficLightFeature.TrafficLightState.Off;
            ExecuteUpdate();
            this.Close();
        }

        private void BtnOrientation_Click(object sender, RoutedEventArgs e)
        {
            _isHorizontal = !_isHorizontal;
            ExecuteUpdate();
            this.Close();
        }

        private void BtnBorder_Click(object sender, RoutedEventArgs e)
        {
            _hasBorder = !_hasBorder;
            ExecuteUpdate();
            this.Close();
        }

        private void ComboWeight_DropDownClosed(object sender, EventArgs e)
        {
            if (ComboWeight.SelectedItem != null && ComboWeight.SelectedItem is ComboBoxItem item)
            {
                if (float.TryParse(item.Content.ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out float newWeight))
                {
                    _borderWeight = newWeight;
                    ExecuteUpdate();
                }
            }
            // Give focus back so we don't hold the dropdown open weirdly
            this.Focus();
        }

        private Button GetCurrentStateButton()
        {
            switch (_currentState)
            {
                case TrafficLightFeature.TrafficLightState.Yellow: return BtnYellow;
                case TrafficLightFeature.TrafficLightState.Green: return BtnGreen;
                case TrafficLightFeature.TrafficLightState.Off: return BtnOff;
                default: return BtnRed;
            }
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape) return;
            e.Handled = true;
            Close();
        }

        private void Window_Deactivated(object sender, EventArgs e)
        {
            if (!_isDialogOpened)
            {
                try { this.Close(); } catch { /* Closing an already-closed context window is idempotent. */ }
            }
        }
    }
}
