using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI.Dialogs
{
    internal sealed class CountryFlagPickerWindow : Window
    {
        private readonly CountryFlagInsertionService _insertion;
        private readonly TextBox _search = new TextBox { MinHeight = 30, Margin = new Thickness(0, 0, 0, 8) };
        private readonly ListBox _results = new ListBox { MinHeight = 340 };

        public CountryFlagPickerWindow(PowerPointManager manager)
        {
            _insertion = new CountryFlagInsertionService(manager);
            Title = "Country Flags";
            Width = 470;
            Height = 520;
            MinWidth = 380;
            MinHeight = 400;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ShowInTaskbar = false;
            Content = BuildLayout();
            _search.TextChanged += (sender, args) => RefreshResults();
            _search.PreviewKeyDown += SearchKeyDown;
            _results.MouseDoubleClick += (sender, args) => InsertSelected();
            _results.PreviewKeyDown += ResultsKeyDown;
            PreviewKeyDown += (sender, args) => { if (args.Key == Key.Escape) Close(); };
            Loaded += (sender, args) => { RefreshResults(); _search.Focus(); };
        }

        private UIElement BuildLayout()
        {
            var root = new DockPanel { Margin = new Thickness(14) };
            var footer = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 10, 0, 0) };
            var cancel = Button("Cancel", (s, e) => Close());
            var insert = Button("Insert flag", (s, e) => InsertSelected());
            footer.Children.Add(cancel); footer.Children.Add(insert);
            DockPanel.SetDock(footer, Dock.Bottom); root.Children.Add(footer);

            var body = new DockPanel();
            DockPanel.SetDock(_search, Dock.Top); body.Children.Add(_search);
            VirtualizingStackPanel.SetIsVirtualizing(_results, true);
            VirtualizingStackPanel.SetVirtualizationMode(_results, VirtualizationMode.Recycling);
            body.Children.Add(_results);
            root.Children.Add(body);
            return root;
        }

        private void RefreshResults()
        {
            IEnumerable<CountryFlagEntry> entries;
            if (string.IsNullOrWhiteSpace(_search.Text))
            {
                var config = JsonService.Load<SidebarConfig>("sidebar_config.json") ?? new SidebarConfig();
                var recent = CountryFlagCatalog.ResolveRecent(config.RecentCountryFlagIds);
                entries = recent.Concat(CountryFlagCatalog.Entries.Where(entry => !recent.Contains(entry)));
            }
            else entries = CountryFlagCatalog.Search(_search.Text);

            _results.Items.Clear();
            foreach (var entry in entries.Take(249)) _results.Items.Add(CreateItem(entry));
            if (_results.Items.Count > 0) _results.SelectedIndex = 0;
        }

        private static ListBoxItem CreateItem(CountryFlagEntry entry)
        {
            var image = new Image { Width = 48, Height = 36, Margin = new Thickness(0, 0, 10, 0) };
            string thumbnail = CountryFlagInsertionService.GetThumbnailPath(entry.AssetId);
            if (File.Exists(thumbnail)) image.Source = new BitmapImage(new Uri(thumbnail, UriKind.Absolute));
            var text = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            text.Children.Add(new TextBlock { Text = entry.Name, FontWeight = FontWeights.SemiBold });
            text.Children.Add(new TextBlock { Text = entry.Alpha2 + " · " + entry.Alpha3, Opacity = 0.7 });
            var panel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(4, 3, 4, 3) };
            panel.Children.Add(image); panel.Children.Add(text);
            var item = new ListBoxItem { Content = panel, Tag = entry, HorizontalContentAlignment = HorizontalAlignment.Stretch };
            System.Windows.Automation.AutomationProperties.SetName(item, entry.Name + " flag");
            return item;
        }

        private void SearchKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down && _results.Items.Count > 0) { _results.Focus(); _results.SelectedIndex = 0; e.Handled = true; }
            else if (e.Key == Key.Enter) { InsertSelected(); e.Handled = true; }
        }

        private void ResultsKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) { InsertSelected(); e.Handled = true; }
            else if (e.Key == Key.Escape) { Close(); e.Handled = true; }
            else if (e.Key == Key.Back && string.IsNullOrEmpty(_search.Text)) { _search.Focus(); }
        }

        private void InsertSelected()
        {
            if (!(_results.SelectedItem is ListBoxItem item) || !(item.Tag is CountryFlagEntry entry)) return;
            if (_insertion.Insert(entry)) Close();
        }

        private static Button Button(string text, RoutedEventHandler click)
        {
            var button = new Button { Content = text, MinWidth = 88, MinHeight = 30, Margin = new Thickness(4) };
            button.Click += click;
            return button;
        }
    }
}
