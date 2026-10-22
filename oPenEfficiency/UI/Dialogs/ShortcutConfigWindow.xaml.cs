using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using oPenEfficiency.Services.Shortcuts;

namespace oPenEfficiency
{
    public partial class ShortcutConfigWindow : Window
    {
        private readonly string _featureId;
        private readonly string _featureName;
        private readonly List<string> _capturedKeys = new List<string>();
        private readonly HashSet<Key> _heldKeys = new HashSet<Key>();
        private readonly IDisposable _modalCaptureScope;
        private bool _capturing;
        private ShortcutValidationResult _validation;

        public ShortcutConfigWindow(string featureId = null)
        {
            InitializeComponent();
            _featureId = featureId;
            _featureName = UI.FeatureLibrary.AllFeatures
                .FirstOrDefault(feature => feature.Id == featureId)?.Name ??
                featureId ??
                "Command";

            Title = "Configure Shortcut for " + _featureName;
            FeatureNameText.Text = _featureName;
            _modalCaptureScope = ShortcutManager.SuspendActivationForModalCapture();
            Closed += (sender, args) => _modalCaptureScope.Dispose();
            PreviewKeyDown += Window_PreviewKeyDown;
            PreviewKeyUp += Window_PreviewKeyUp;

            LoadSavedShortcut();
        }

        private void LoadSavedShortcut()
        {
            var assignment = ShortcutManager.GetAssignment(_featureId);
            if (assignment != null)
            {
                _capturedKeys.AddRange(
                    ShortcutPolicy.SplitStorageValue(assignment.RawValue)
                        .Where(key => !string.IsNullOrWhiteSpace(key) &&
                                      !string.Equals(
                                          key.Trim(),
                                          "None",
                                          StringComparison.OrdinalIgnoreCase)));
                if (assignment.NeedsAttention)
                {
                    AttentionText.Text =
                        "Requires attention: " + ResolveFeatureNames(assignment.AttentionReason);
                    AttentionText.Visibility = Visibility.Visible;
                }
            }

            UpdateValidation();
        }

        private void BtnCapture_Click(object sender, RoutedEventArgs e)
        {
            _capturedKeys.Clear();
            _heldKeys.Clear();
            _capturing = true;
            AttentionText.Visibility = Visibility.Collapsed;
            BtnCapture.Content = "Hold the complete chord, then release";
            BtnSave.IsEnabled = false;
            ChordText.Text = "Listening…";
            BtnCapture.Focus();
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!_capturing) return;

            Key key = ResolveKey(e);
            if (key != Key.None && _heldKeys.Add(key))
                _capturedKeys.Add(key.ToString());

            UpdateValidation();
            e.Handled = true;
        }

        private void Window_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (!_capturing) return;

            Key key = ResolveKey(e);
            if (key != Key.None)
                _heldKeys.Remove(key);
            e.Handled = true;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_heldKeys.Count == 0 && Keyboard.Modifiers == ModifierKeys.None)
                {
                    _capturing = false;
                    BtnCapture.Content = "Capture again";
                    UpdateValidation();
                }
            }));
        }

        private void UpdateValidation()
        {
            _validation = ShortcutManager.ValidateShortcut(
                _featureId,
                _capturedKeys,
                riskConfirmed: false);

            ChordText.Text = _validation.Chord?.DisplayText ??
                (_capturedKeys.Count > 0
                    ? ShortcutPolicy.FormatLegacyValue(string.Join(",", _capturedKeys))
                    : "No chord captured");

            if (_validation.RequiresRiskConfirmation)
            {
                ValidationText.Text = _validation.Warning +
                    " Saving requires an explicit confirmation.";
                ValidationText.Foreground =
                    (Brush)FindResource("WarningBrush");
                BtnSave.IsEnabled = !_capturing;
            }
            else if (!_validation.IsValid)
            {
                string validationError = _validation.Error;
                if (!string.IsNullOrWhiteSpace(_validation.ConflictingFeatureId))
                {
                    validationError += " Conflicts with " +
                        GetFeatureName(_validation.ConflictingFeatureId) + ".";
                }
                ValidationText.Text = ResolveFeatureNames(validationError);
                ValidationText.Foreground =
                    (Brush)FindResource("ErrorBrush");
                BtnSave.IsEnabled = false;
            }
            else
            {
                ValidationText.Text =
                    "Chord is valid and available for this command.";
                ValidationText.Foreground =
                    (Brush)FindResource("TextSecondaryBrush");
                BtnSave.IsEnabled = !_capturing;
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_featureId) || _capturing) return;

            bool riskConfirmed = false;
            if (_validation != null && _validation.RequiresRiskConfirmation)
            {
                var answer = MessageBox.Show(
                    _validation.Warning +
                    "\n\nThis can run a command while you are typing or navigating in PowerPoint. Save it anyway?",
                    "Confirm Shortcut Activation Risk",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);
                if (answer != MessageBoxResult.Yes) return;
                riskConfirmed = true;
            }

            if (!ShortcutManager.TryReplaceShortcut(
                _featureId,
                _capturedKeys,
                riskConfirmed,
                out var validation,
                out string error))
            {
                _validation = validation;
                MessageBox.Show(
                    ResolveFeatureNames(error ?? "The shortcut could not be saved."),
                    "Shortcut Not Saved",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                UpdateValidation();
                return;
            }

            DialogResult = true;
            Close();
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_featureId)) return;
            if (!ShortcutManager.ClearShortcut(_featureId))
            {
                MessageBox.Show(
                    "The assignment could not be cleared. The previous assignment remains unchanged.",
                    "Shortcut Not Cleared",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e) => Close();

        private string ResolveFeatureNames(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;
            foreach (var feature in UI.FeatureLibrary.AllFeatures)
            {
                text = text.Replace(feature.Id, feature.Name);
            }
            return text;
        }

        private static string GetFeatureName(string featureId) =>
            UI.FeatureLibrary.AllFeatures
                .FirstOrDefault(feature => feature.Id == featureId)?.Name ??
            featureId;

        private static Key ResolveKey(KeyEventArgs e) =>
            e.Key == Key.System ? e.SystemKey : e.Key;
    }
}
