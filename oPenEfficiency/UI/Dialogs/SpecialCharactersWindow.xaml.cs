using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using oPenEfficiency.Features;
using oPenEfficiency.Models;

namespace oPenEfficiency.UI
{
    public partial class SpecialCharactersWindow : Window
    {
        private readonly PowerPointManager _ppManager;

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
            foreach (var category in SpecialCharacterCatalog.Symbols)
            {
                CategoryList.Items.Add(new CategoryItem { Header = category.Header, Characters = category.Characters });
            }

            // Languages
            foreach (var category in SpecialCharacterCatalog.Languages)
            {
                CategoryList.Items.Add(new CategoryItem { Header = category.Header, Characters = category.Characters });
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
                foreach (string glyph in SpecialCharacterCatalog.EnumerateGlyphs(item.Characters))
                {
                    var btn = new Button
                    {
                        Content = glyph,
                        Style = (Style)FindResource("CharButtonStyle"),
                        Tag = glyph
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
                if (obj is CategoryItem item && SpecialCharacterCatalog.MatchesHeader(item.Header, language))
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
            try { this.DragMove(); } catch { /* DragMove can be rejected when the mouse button is released mid-event. */ }
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
