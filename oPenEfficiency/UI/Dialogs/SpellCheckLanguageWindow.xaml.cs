using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Core;

namespace oPenEfficiency.UI.Dialogs
{
    public enum SpellCheckScope
    {
        Presentation,
        SelectedSlides,
        SelectedShapes
    }

    public partial class SpellCheckLanguageWindow : Window
    {
        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                this.DragMove();
        }
        
        private void BtnClose_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        public MsoLanguageID SelectedLanguage { get; private set; }
        public SpellCheckScope SelectedScope { get; private set; }

        private List<KeyValuePair<string, MsoLanguageID>> _allLanguages;
        private MsoLanguageID? _initialLanguage;
        
        public SpellCheckLanguageWindow(MsoLanguageID? initialLanguage = null)
        {
            _initialLanguage = initialLanguage;
            InitializeComponent();
            PopulateLanguages();
        }

        private void PopulateLanguages()
        {
            _allLanguages = new List<KeyValuePair<string, MsoLanguageID>>();

            // Add standard common languages first
            var common = new Dictionary<string, MsoLanguageID>
            {
                { "English (US)", MsoLanguageID.msoLanguageIDEnglishUS },
                { "English (UK)", MsoLanguageID.msoLanguageIDEnglishUK },
                { "German", MsoLanguageID.msoLanguageIDGerman },
                { "French", MsoLanguageID.msoLanguageIDFrench },
                { "Spanish", MsoLanguageID.msoLanguageIDSpanish },
                { "Italian", MsoLanguageID.msoLanguageIDItalian },
                { "Dutch", MsoLanguageID.msoLanguageIDDutch },
                { "Portuguese (Brazil)", MsoLanguageID.msoLanguageIDBrazilianPortuguese },
                { "Japanese", MsoLanguageID.msoLanguageIDJapanese },
                { "Chinese (Simplified)", MsoLanguageID.msoLanguageIDSimplifiedChinese }
            };

            foreach (var lang in common) _allLanguages.Add(lang);

            // Add all other MsoLanguageID values via reflection for completeness
            var allValues = Enum.GetValues(typeof(MsoLanguageID)).Cast<MsoLanguageID>();
            foreach (var val in allValues)
            {
                if (common.Values.Contains(val)) continue;
                
                string name = val.ToString().Replace("msoLanguageID", "");
                if (string.IsNullOrEmpty(name)) continue;
                
                // Add spaces before capitals for readability
                name = System.Text.RegularExpressions.Regex.Replace(name, "(\\B[A-Z])", " $1");
                _allLanguages.Add(new KeyValuePair<string, MsoLanguageID>(name, val));
            }

            CmbLanguages.ItemsSource = _allLanguages.OrderBy(x => x.Key).ToList();
            
            if (_initialLanguage.HasValue)
            {
                CmbLanguages.SelectedIndex = _allLanguages.FindIndex(x => x.Value == _initialLanguage.Value);
            }
            if (CmbLanguages.SelectedIndex == -1)
            {
                CmbLanguages.SelectedIndex = _allLanguages.FindIndex(x => x.Value == MsoLanguageID.msoLanguageIDGerman);
            }
            if (CmbLanguages.SelectedIndex == -1) CmbLanguages.SelectedIndex = 0;
        }

        private void CmbLanguages_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (CmbLanguages.IsDropDownOpen)
            {
                string filter = CmbLanguages.Text.ToLower();
                if (string.IsNullOrEmpty(filter))
                {
                    CmbLanguages.ItemsSource = _allLanguages.OrderBy(x => x.Key).ToList();
                }
                else
                {
                    CmbLanguages.ItemsSource = _allLanguages
                        .Where(x => x.Key.ToLower().Contains(filter))
                        .OrderBy(x => x.Key)
                        .ToList();
                }
            }
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            if (CmbLanguages.SelectedItem == null) return;

            SelectedLanguage = (MsoLanguageID)((KeyValuePair<string, MsoLanguageID>)CmbLanguages.SelectedItem).Value;
            
            if (RadioWholePres.IsChecked == true) SelectedScope = SpellCheckScope.Presentation;
            else if (RadioSelectedSlides.IsChecked == true) SelectedScope = SpellCheckScope.SelectedSlides;
            else if (RadioSelectedShapes.IsChecked == true) SelectedScope = SpellCheckScope.SelectedShapes;

            this.DialogResult = true;
            this.Close();
        }
    }
}
