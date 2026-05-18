using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace oPenEfficiency
{
    public partial class ShortcutConfigWindow : Window
    {
        private string _featureId;
        private string _configPath;

        public ShortcutConfigWindow(string featureId = null)
        {
            InitializeComponent();
            _featureId = featureId;
            // Store config in AppData
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var folder = System.IO.Path.Combine(appData, "oPenEfficiency");
            if (!System.IO.Directory.Exists(folder)) System.IO.Directory.CreateDirectory(folder);
            _configPath = System.IO.Path.Combine(folder, "shortcuts.cfg");

            LoadKeys();
            
            if (!string.IsNullOrEmpty(_featureId))
            {
                this.Title = $"Configure Shortcuts for {_featureId}";
                LoadSavedShortcut();
            }

            this.PreviewKeyDown += Window_PreviewKeyDown;
        }

        private void LoadKeys()
        {
            var keys = Enum.GetValues(typeof(Key))
                .Cast<Key>()
                .Distinct()
                .OrderBy(k => 
                {
                    if (k == Key.None) return 0;
                    if (k == Key.LeftCtrl || k == Key.RightCtrl) return 1;
                    if (k == Key.LeftAlt || k == Key.RightAlt) return 2;
                    if (k == Key.LeftShift || k == Key.RightShift) return 3;
                    if (k == Key.System) return 4;
                    if (k >= Key.A && k <= Key.Z) return 5;
                    if (k >= Key.D0 && k <= Key.D9) return 6;
                    return 7;
                })
                .ThenBy(k => k.ToString())
                .ToList();
                
            ComboKey1.ItemsSource = keys;
            ComboKey2.ItemsSource = keys;
            ComboKey3.ItemsSource = keys;
            ComboKey4.ItemsSource = keys;
            
            ComboKey1.SelectedItem = Key.None;
            ComboKey2.SelectedItem = Key.None;
            ComboKey3.SelectedItem = Key.None;
            ComboKey4.SelectedItem = Key.None;
        }

        private void LoadSavedShortcut()
        {
            if (!System.IO.File.Exists(_configPath)) return;
            
            var lines = System.IO.File.ReadAllLines(_configPath);
            foreach (var line in lines)
            {
                var parts = line.Split('=');
                if (parts.Length == 2 && parts[0] == _featureId)
                {
                    var keys = parts[1].Split(',');
                    if (keys.Length >= 1) ComboKey1.SelectedItem = (Key)Enum.Parse(typeof(Key), keys[0]);
                    if (keys.Length >= 2) ComboKey2.SelectedItem = (Key)Enum.Parse(typeof(Key), keys[1]);
                    if (keys.Length >= 3) ComboKey3.SelectedItem = (Key)Enum.Parse(typeof(Key), keys[2]);
                    if (keys.Length >= 4) ComboKey4.SelectedItem = (Key)Enum.Parse(typeof(Key), keys[3]);
                    break;
                }
            }
        }

        private int _listeningCombo = 0;

        private void ResetButtons()
        {
            BtnAssign1.Content = "Assign";
            BtnAssign2.Content = "Assign";
            BtnAssign3.Content = "Assign";
            BtnAssign4.Content = "Assign";
        }

        private void BtnAssign1_Click(object sender, RoutedEventArgs e) { ResetButtons(); _listeningCombo = 1; BtnAssign1.Content = "Press..."; }
        private void BtnAssign2_Click(object sender, RoutedEventArgs e) { ResetButtons(); _listeningCombo = 2; BtnAssign2.Content = "Press..."; }
        private void BtnAssign3_Click(object sender, RoutedEventArgs e) { ResetButtons(); _listeningCombo = 3; BtnAssign3.Content = "Press..."; }
        private void BtnAssign4_Click(object sender, RoutedEventArgs e) { ResetButtons(); _listeningCombo = 4; BtnAssign4.Content = "Press..."; }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (_listeningCombo > 0)
            {
                Key key = e.Key;
                if (key == Key.System) key = e.SystemKey; 
                
                switch (_listeningCombo)
                {
                    case 1: ComboKey1.SelectedItem = key; break;
                    case 2: ComboKey2.SelectedItem = key; break;
                    case 3: ComboKey3.SelectedItem = key; break;
                    case 4: ComboKey4.SelectedItem = key; break;
                }
                
                ResetButtons();
                _listeningCombo = 0;
                e.Handled = true;
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_featureId)) return;

            var k1 = ComboKey1.SelectedItem?.ToString() ?? "None";
            var k2 = ComboKey2.SelectedItem?.ToString() ?? "None";
            var k3 = ComboKey3.SelectedItem?.ToString() ?? "None";
            var k4 = ComboKey4.SelectedItem?.ToString() ?? "None";
            var value = $"{k1},{k2},{k3},{k4}";

            var lines = System.IO.File.Exists(_configPath) ? System.IO.File.ReadAllLines(_configPath).ToList() : new System.Collections.Generic.List<string>();
            var updated = false;

            for (int i = 0; i < lines.Count; i++)
            {
                if (lines[i].StartsWith(_featureId + "="))
                {
                    lines[i] = $"{_featureId}={value}";
                    updated = true;
                    break;
                }
            }

            if (!updated)
            {
                lines.Add($"{_featureId}={value}");
            }

            System.IO.File.WriteAllLines(_configPath, lines);
            
            ShortcutManager.Reload();
            MessageBox.Show($"Shortcut saved for {_featureId}: {k1} + {k2} + {k3} + {k4}");
            this.Close();
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            ResetButtons();
            _listeningCombo = 0;
            ComboKey1.SelectedItem = Key.None;
            ComboKey2.SelectedItem = Key.None;
            ComboKey3.SelectedItem = Key.None;
            ComboKey4.SelectedItem = Key.None;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
