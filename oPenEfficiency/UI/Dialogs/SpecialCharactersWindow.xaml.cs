using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using oPenEfficiency.Features;

namespace oPenEfficiency.UI
{
    public partial class SpecialCharactersWindow : Window
    {
        private readonly PowerPointManager _ppManager;

        public static readonly Dictionary<string, string> LanguageData = new Dictionary<string, string>
        {
            { "German", "äöüßÄÖÜ" },
            { "French", "àâçéèêëîïôûùÀÂÇÉÈÊËÎÏÔÛÙ" },
            { "Spanish", "áéíóúüñÁÉÍÓÚÜÑ¿¡" },
            { "Italian", "àèéìòùÀÈÉÌÒÙ" },
            { "Portuguese", "áâãàçéêíóôõúÁÂÃÀÇÉÊÍÓÔÕÚ" },
            { "Nordic", "åæøÅÆØðþÐÞ" },
            { "Polish", "ąćęłńóśźżĄĆĘŁŃÓŚŹŻ" },
            { "Turkish", "çğışöüÇĞİŞÖÜ" },
            { "Czech/Slovak", "áčďéěíňóřšťúůýžÁČĎÉĚÍŇÓŘŠŤÚŮÝŽ" },
            { "Hungarian", "áéíóöőúüűÁÉÍÓÖŐÚÜŰ" },
            { "Romanian", "ăâîşţĂÂÎŞŢ" }
        };

        public static readonly Dictionary<string, string> SymbolData = new Dictionary<string, string>
        {
            { "Mathematical", "¼½¾±≈≠≤≥∞∑∆∏π√∫‰ØΔμΣΩ" },
            { "Business", "©®™£¥€¢¤§¶" },
            { "Arrows & Process", "➔⇨⇒➔↔↕↖↗↘↙↺↻⟳➞➠" },
            { "Status & Indicators", "⚠⚐🏁🎯🚀⌛⏳✅❌❓❗💡📌🔍🔗" },
            { "Harvey Balls", "○◔◑◕●" },
            { "Punctuation", "…—–•·«»“”" },
            { "Miscellaneous", "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳✓✕✘⚡" }
        };

        public SpecialCharactersWindow(PowerPointManager ppManager)
        {
            InitializeComponent();
            _ppManager = ppManager;
            this.Closing += (s, e) => { e.Cancel = true; this.Hide(); };
            PopulateCategories();
        }

        private void PopulateCategories()
        {
            if (CategoryList == null) return;
            CategoryList.Items.Clear();

            // Symbols first
            foreach (var category in SymbolData)
            {
                CategoryList.Items.Add(new CategoryItem { Header = category.Key, Characters = category.Value });
            }

            // Languages
            foreach (var lang in LanguageData)
            {
                CategoryList.Items.Add(new CategoryItem { Header = lang.Key, Characters = lang.Value });
            }

            if (CategoryList.Items.Count > 0)
                CategoryList.SelectedIndex = 0;
        }

        private void CategoryList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CharactersPanel == null) return;
            if (CategoryList.SelectedItem is CategoryItem item)
            {
                CharactersPanel.Children.Clear();
                foreach (char c in item.Characters)
                {
                    var btn = new Button
                    {
                        Content = c.ToString(),
                        Style = (Style)FindResource("CharButtonStyle"),
                        Tag = c.ToString()
                    };
                    btn.Click += CharButton_Click;
                    CharactersPanel.Children.Add(btn);
                }
            }
        }

        private void CharButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string charToInsert)
            {
                InsertTextAtCursorFeature.Execute(_ppManager, charToInsert);
            }
        }

        public void SelectLanguage(string language)
        {
            if (CategoryList == null) return;
            foreach (var obj in CategoryList.Items)
            {
                if (obj is CategoryItem item && item.Header.Equals(language, StringComparison.OrdinalIgnoreCase))
                {
                    CategoryList.SelectedItem = item;
                    break;
                }
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try { this.DragMove(); } catch { }
        }

        protected override void OnDeactivated(EventArgs e)
        {
            base.OnDeactivated(e);
        }
    }

    public class CategoryItem
    {
        public string Header { get; set; }
        public string Characters { get; set; }
        public override string ToString() => Header;
    }
}
