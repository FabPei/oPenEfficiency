using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using oPenEfficiency.Models;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.UI.Dialogs
{
    internal sealed class QuickContextMenuConfigWindow : Window
    {
        private readonly ListBox _available = new ListBox { MinWidth = 230, MinHeight = 300 };
        private readonly ListBox _selected = new ListBox { MinWidth = 230, MinHeight = 300 };
        private readonly Dictionary<string, FeatureCatalogRecord> _catalog;

        public QuickContextMenuConfigWindow()
        {
            Title = "Configure Quick Context Menu";
            Width = 620;
            Height = 460;
            MinWidth = 560;
            MinHeight = 400;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ShowInTaskbar = false;
            _catalog = FeatureLibrary.Catalog
                .Where(item => item.UserVisible && !item.IsObsolete &&
                    item.ExecutionKind == FeatureExecutionKind.Command &&
                    item.Id != "BtnQuickContextMenu")
                .ToDictionary(item => item.Id, StringComparer.Ordinal);

            var config = JsonService.Load<SidebarConfig>("sidebar_config.json") ?? new SidebarConfig();
            IEnumerable<string> configuredIds = config.QuickContextMenuFeatureIds?.Count > 0
                ? (IEnumerable<string>)config.QuickContextMenuFeatureIds
                : QuickContextMenuSettings.DefaultFeatureIds;
            var ids = QuickContextMenuSettings.Normalize(
                configuredIds,
                new HashSet<string>(_catalog.Keys, StringComparer.Ordinal));
            RefreshLists(ids);
            Content = BuildLayout();
        }

        private UIElement BuildLayout()
        {
            var root = new DockPanel { Margin = new Thickness(16) };
            var intro = new TextBlock
            {
                Text = "Choose and order up to 12 commands. Unavailable commands stay visible but disabled for the current selection.",
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 12)
            };
            DockPanel.SetDock(intro, Dock.Top);
            root.Children.Add(intro);

            var footer = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
            var reset = MakeButton("Reset defaults", (s, e) => RefreshLists(QuickContextMenuSettings.DefaultFeatureIds));
            var cancel = MakeButton("Cancel", (s, e) => Close());
            var save = MakeButton("Save", Save);
            footer.Children.Add(reset); footer.Children.Add(cancel); footer.Children.Add(save);
            DockPanel.SetDock(footer, Dock.Bottom);
            root.Children.Add(footer);

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            grid.Children.Add(Wrap("Available commands", _available, 0));
            grid.Children.Add(Wrap("Quick Menu order", _selected, 2));
            var controls = new StackPanel { VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(10) };
            controls.Children.Add(MakeButton("Add →", Add));
            controls.Children.Add(MakeButton("← Remove", Remove));
            controls.Children.Add(MakeButton("Move up", MoveUp));
            controls.Children.Add(MakeButton("Move down", MoveDown));
            Grid.SetColumn(controls, 1);
            grid.Children.Add(controls);
            root.Children.Add(grid);
            return root;
        }

        private static FrameworkElement Wrap(string heading, ListBox list, int column)
        {
            var panel = new DockPanel();
            var label = new TextBlock { Text = heading, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 6) };
            DockPanel.SetDock(label, Dock.Top);
            panel.Children.Add(label);
            panel.Children.Add(list);
            Grid.SetColumn(panel, column);
            return panel;
        }

        private static Button MakeButton(string text, RoutedEventHandler handler)
        {
            var button = new Button { Content = text, MinWidth = 86, MinHeight = 28, Margin = new Thickness(4) };
            button.Click += handler;
            return button;
        }

        private void Add(object sender, RoutedEventArgs e)
        {
            if (!(_available.SelectedItem is FeatureChoice choice) || _selected.Items.Count >= QuickContextMenuSettings.MaximumItems) return;
            _selected.Items.Add(choice);
            _available.Items.Remove(choice);
        }

        private void Remove(object sender, RoutedEventArgs e)
        {
            if (!(_selected.SelectedItem is FeatureChoice choice)) return;
            _selected.Items.Remove(choice);
            InsertAvailable(choice);
        }

        private void MoveUp(object sender, RoutedEventArgs e) => MoveSelected(-1);
        private void MoveDown(object sender, RoutedEventArgs e) => MoveSelected(1);

        private void MoveSelected(int delta)
        {
            int index = _selected.SelectedIndex;
            int target = index + delta;
            if (index < 0 || target < 0 || target >= _selected.Items.Count) return;
            object item = _selected.Items[index];
            _selected.Items.RemoveAt(index);
            _selected.Items.Insert(target, item);
            _selected.SelectedIndex = target;
        }

        private void Save(object sender, RoutedEventArgs e)
        {
            if (_selected.Items.Count < QuickContextMenuSettings.MinimumItems)
            {
                MessageBox.Show(this, "Select at least one command.", Title, MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            var config = JsonService.Load<SidebarConfig>("sidebar_config.json") ?? new SidebarConfig();
            config.QuickContextMenuFeatureIds = _selected.Items.Cast<FeatureChoice>().Select(item => item.Id).ToList();
            if (!JsonService.TrySave(config, "sidebar_config.json"))
            {
                MessageBox.Show(this, "The configuration could not be saved. See the local diagnostics log.", Title, MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            DialogResult = true;
        }

        private void RefreshLists(IEnumerable<string> selectedIds)
        {
            _available.Items.Clear();
            _selected.Items.Clear();
            var selected = QuickContextMenuSettings.Normalize(selectedIds, new HashSet<string>(_catalog.Keys, StringComparer.Ordinal));
            foreach (string id in selected) _selected.Items.Add(new FeatureChoice(_catalog[id]));
            foreach (var feature in _catalog.Values.Where(item => !selected.Contains(item.Id)).OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase))
                _available.Items.Add(new FeatureChoice(feature));
        }

        private void InsertAvailable(FeatureChoice choice)
        {
            int index = 0;
            while (index < _available.Items.Count && string.Compare(((FeatureChoice)_available.Items[index]).Name, choice.Name, StringComparison.CurrentCultureIgnoreCase) < 0) index++;
            _available.Items.Insert(index, choice);
        }

        private sealed class FeatureChoice
        {
            public FeatureChoice(FeatureCatalogRecord feature) { Id = feature.Id; Name = feature.Name; }
            public string Id { get; }
            public string Name { get; }
            public override string ToString() => Name;
        }
    }
}
