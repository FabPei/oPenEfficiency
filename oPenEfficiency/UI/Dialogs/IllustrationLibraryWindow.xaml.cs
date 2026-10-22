using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI.Dialogs
{
    /// <summary>
    /// Two views in one window: the picker that places the bundled and imported figures on a slide, and the
    /// source list that shows where the artwork comes from. Importing is deliberately manual - the user picks
    /// the files, the window screens them and records which license applied.
    /// </summary>
    public partial class IllustrationLibraryWindow : Window
    {
        private readonly IllustrationImportService _importService = new IllustrationImportService();
        private readonly IllustrationCatalogService _catalogService = new IllustrationCatalogService();
        private readonly IllustrationInsertionService _insertionService = new IllustrationInsertionService();
        private readonly PowerPointManager _manager;
        private readonly List<IllustrationSource> _sources;
        private readonly List<FigureRow> _figureRows = new List<FigureRow>();
        private readonly List<PaletteToggle> _paletteToggles = new List<PaletteToggle>();
        private readonly List<PartToggle> _partToggles = new List<PartToggle>();
        private string _figuresSubtitle = string.Empty;

        public IllustrationLibraryWindow()
            : this(null)
        {
        }

        public IllustrationLibraryWindow(PowerPointManager manager)
        {
            InitializeComponent();

            _manager = manager;
            _sources = IllustrationSourceCatalog.All.ToList();

            NoticeText.Text = IllustrationSourceCatalog.ResponsibilityNotice;
            SourceList.ItemsSource = _sources;
            SourceList.SelectedIndex = 0;

            LoadFigures();
            SetActiveView(true);
            SetStatus("Imported files are stored per user in " + _importService.LibraryRoot);
        }

        private IllustrationSource SelectedSource
        {
            get { return SourceList.SelectedItem as IllustrationSource; }
        }

        private FigureRow SelectedFigure
        {
            get { return FigureList.SelectedItem as FigureRow; }
        }

        // ---------------------------------------------------------------- figures

        private void LoadFigures()
        {
            _figureRows.Clear();

            IReadOnlyList<IllustrationItem> items;
            try
            {
                items = _catalogService.LoadItems();
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationLibraryWindow.LoadFigures");
                InsertFeedback.Text = "The illustration catalog could not be read: " + ex.Message;
                return;
            }

            foreach (var item in items)
            {
                _figureRows.Add(new FigureRow
                {
                    Item = item,
                    Name = item.Name,
                    DetailLine = DescribeFigure(item),
                    ThumbnailPath = _catalogService.ResolveItemPath(item)
                });
            }

            var collections = items
                .Select(item => item.CollectionId)
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(id => _catalogService.DescribeCollection(id), StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var previous = GetSelectedCollectionId();
            CollectionFilterBox.Items.Clear();
            CollectionFilterBox.Items.Add(new ComboBoxItem { Content = "All collections", Tag = string.Empty });
            foreach (var collectionId in collections)
            {
                CollectionFilterBox.Items.Add(new ComboBoxItem
                {
                    Content = _catalogService.DescribeCollection(collectionId),
                    Tag = collectionId
                });
            }

            CollectionFilterBox.SelectedIndex = 0;
            if (!string.IsNullOrWhiteSpace(previous))
            {
                for (int i = 0; i < CollectionFilterBox.Items.Count; i++)
                {
                    var candidate = CollectionFilterBox.Items[i] as ComboBoxItem;
                    if (candidate != null && string.Equals(candidate.Tag as string, previous, StringComparison.Ordinal))
                    {
                        CollectionFilterBox.SelectedIndex = i;
                        break;
                    }
                }
            }

            int imported = _figureRows.Count(row => row.Item.Tags != null && row.Item.Tags.Contains("imported"));
            _figuresSubtitle = string.Format(
                "{0} figures are ready to insert ({1} bundled, {2} imported); {3} further sources can be opened and imported from here.",
                _figureRows.Count,
                _figureRows.Count - imported,
                imported,
                _sources.Count - IllustrationSourceCatalog.Bundled.Count);
            HeaderSubtitle.Text = _figuresSubtitle;

            ApplyFigureFilter();
        }

        private string DescribeFigure(IllustrationItem item)
        {
            var plan = _catalogService.LoadPlan(item);
            int partCount = plan == null ? 0 : plan.Parts.Count(part => !part.IsPlaceholder);
            int colorCount = plan == null ? 0 : plan.Palette.Count;

            var pieces = new List<string>
            {
                _catalogService.DescribeCollection(item.CollectionId),
                partCount > 0 ? partCount + " named parts" : "one object",
                colorCount + (colorCount == 1 ? " colour" : " colours")
            };

            if (item.Tags != null && item.Tags.Contains("imported")) pieces.Add("imported by you");

            return string.Join(" - ", pieces);
        }

        private string GetSelectedCollectionId()
        {
            var selected = CollectionFilterBox.SelectedItem as ComboBoxItem;
            return selected == null ? string.Empty : (selected.Tag as string ?? string.Empty);
        }

        private void ApplyFigureFilter()
        {
            string query = FigureSearchBox.Text;
            string collectionId = GetSelectedCollectionId();

            var visible = _figureRows
                .Where(row => string.IsNullOrWhiteSpace(collectionId) ||
                              string.Equals(row.Item.CollectionId, collectionId, StringComparison.Ordinal))
                .Where(row => MatchesQuery(row, query))
                .ToList();

            FigureList.ItemsSource = visible;
            if (visible.Count > 0) FigureList.SelectedIndex = 0;

            HeaderSubtitle.Text = visible.Count == _figureRows.Count
                ? _figuresSubtitle
                : string.Format("{0} of {1} figures shown.", visible.Count, _figureRows.Count);
        }

        private static bool MatchesQuery(FigureRow row, string query)
        {
            if (string.IsNullOrWhiteSpace(query)) return true;

            var terms = query
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(term => term.Trim())
                .Where(term => term.Length > 0)
                .ToList();
            if (terms.Count == 0) return true;

            var haystack = ((row.Name ?? string.Empty) + " " + (row.DetailLine ?? string.Empty) + " " +
                            string.Join(" ", row.Item.Tags ?? new List<string>())).ToLowerInvariant();

            return terms.All(term => haystack.Contains(term.ToLowerInvariant()));
        }

        private void FigureSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!IsLoaded && FigureList == null) return;
            ApplyFigureFilter();
        }

        private void FigureSearchBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter || FigureList.Items.Count == 0) return;

            FigureList.SelectedIndex = Math.Max(0, FigureList.SelectedIndex);
            FigureList.Focus();
            FigureList.ScrollIntoView(FigureList.SelectedItem);
            e.Handled = true;
        }

        private void FigureList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter || SelectedFigure == null) return;

            InsertSelectedFigure();
            e.Handled = true;
        }

        private void FigureList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (SelectedFigure == null) return;
            InsertSelectedFigure();
            e.Handled = true;
        }

        private void CollectionFilterBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyFigureFilter();
        }

        private void FigureList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ShowFigureDetails();
        }

        private void ShowFigureDetails()
        {
            var row = SelectedFigure;
            PalettePanel.Children.Clear();
            PartsPanel.Children.Clear();
            _paletteToggles.Clear();
            _partToggles.Clear();
            InsertFeedback.Text = string.Empty;

            if (row == null)
            {
                FigureName.Text = string.Empty;
                FigureMeta.Text = string.Empty;
                FigureLicense.Text = string.Empty;
                PartsHint.Text = string.Empty;
                InsertButton.IsEnabled = false;
                return;
            }

            var item = row.Item;
            var plan = _catalogService.LoadPlan(item);

            FigureName.Text = item.Name;
            FigureMeta.Text = row.DetailLine;
            FigureLicense.Text = DescribeLicense(item);

            if (plan == null)
            {
                PartsHint.Text = "The artwork file could not be read.";
                InsertButton.IsEnabled = false;
                return;
            }

            // Colours
            if (plan.Palette.Count == 0)
            {
                PalettePanel.Children.Add(BuildHint("This artwork carries no colours to replace."));
            }
            else
            {
                foreach (var hex in plan.Palette)
                {
                    PalettePanel.Children.Add(BuildPaletteRow(hex));
                }
            }

            // Parts
            var parts = IllustrationInsertionPlanner.SelectableParts(plan);
            if (parts.Count == 0)
            {
                PartsHint.Text = "This artwork is one object; it cannot be split into parts.";
            }
            else
            {
                PartsHint.Text = "Untick a part to leave it out of the insertion.";
                foreach (var part in parts)
                {
                    var toggle = new CheckBox
                    {
                        Content = part.Id,
                        IsChecked = true,
                        Margin = new Thickness(0, 2, 0, 2),
                        Foreground = (Brush)TryFindResource("TextSecondaryBrush")
                    };
                    _partToggles.Add(new PartToggle { PartId = part.Id, Box = toggle });
                    PartsPanel.Children.Add(toggle);
                }
            }

            InsertModeBox.SelectedIndex = 0;
            InsertSizeBox.SelectedIndex = 1;
            InsertButton.IsEnabled = true;
        }

        private string DescribeLicense(IllustrationItem item)
        {
            string collectionId = item.CollectionId ?? string.Empty;
            string importedPrefix = "imported-from-";

            if (collectionId.StartsWith(importedPrefix, StringComparison.Ordinal))
            {
                var importedSource = IllustrationSourceCatalog.Find(collectionId.Substring(importedPrefix.Length));
                string license = importedSource == null
                    ? "The license of this source is recorded next to the source list."
                    : importedSource.LicenseSummary;
                return "Imported by you from " + (importedSource == null ? "an external source" : importedSource.Name) +
                       ". " + license + " You keep the responsibility for this file.";
            }

            var source = FindSourceForCollection(collectionId);
            if (source != null)
            {
                string text = "Bundled with oPenEfficiency - " + source.LicenseSummary;
                return source.AttributionRequired
                    ? text + " Credit is required wherever this artwork appears."
                    : text + " No credit is required.";
            }

            return "Bundled with oPenEfficiency under CC0 1.0 Universal.";
        }

        private IllustrationSource FindSourceForCollection(string collectionId)
        {
            if (string.IsNullOrWhiteSpace(collectionId)) return null;

            var exact = IllustrationSourceCatalog.Find(collectionId);
            if (exact != null) return exact;

            return IllustrationSourceCatalog.All
                .Where(source => source != null && !string.IsNullOrWhiteSpace(source.Id))
                .Where(source => collectionId.StartsWith(source.Id + "-", StringComparison.Ordinal))
                .OrderByDescending(source => source.Id.Length)
                .FirstOrDefault();
        }

        private UIElement BuildHint(string text)
        {
            return new TextBlock
            {
                Text = text,
                TextWrapping = TextWrapping.Wrap,
                Foreground = (Brush)TryFindResource("TextSecondaryBrush")
            };
        }

        private UIElement BuildPaletteRow(string originalHex)
        {
            var grid = new Grid { Margin = new Thickness(0, 2, 0, 2) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var swatch = new Border
            {
                Width = 22,
                Height = 22,
                CornerRadius = new CornerRadius(3),
                BorderThickness = new Thickness(1),
                BorderBrush = (Brush)TryFindResource("BorderBrush"),
                Background = BuildBrush(originalHex)
            };
            Grid.SetColumn(swatch, 0);

            var label = new TextBlock
            {
                Text = originalHex,
                Margin = new Thickness(8, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = (Brush)TryFindResource("TextPrimaryBrush")
            };
            Grid.SetColumn(label, 1);

            var button = new Button
            {
                Content = "Change",
                Height = 24,
                Padding = new Thickness(10, 0, 10, 0),
                Style = (Style)TryFindResource("SecondaryButton"),
                Tag = originalHex
            };
            button.Click += PaletteChange_Click;
            Grid.SetColumn(button, 2);

            grid.Children.Add(swatch);
            grid.Children.Add(label);
            grid.Children.Add(button);

            _paletteToggles.Add(new PaletteToggle { OriginalHex = originalHex, Swatch = swatch, Label = label });
            return grid;
        }

        private void PaletteChange_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button == null) return;

            string originalHex = button.Tag as string;
            if (string.IsNullOrWhiteSpace(originalHex)) return;

            string initial = originalHex;
            var existing = _paletteToggles.FirstOrDefault(toggle => string.Equals(toggle.OriginalHex, originalHex, StringComparison.OrdinalIgnoreCase));
            if (existing != null && !string.IsNullOrWhiteSpace(existing.ReplacementHex)) initial = existing.ReplacementHex;

            var dialog = new WpfColorPickerDialog(initial) { Owner = this };
            if (dialog.ShowDialog() != true) return;

            string replacement = string.Format("#{0:X2}{1:X2}{2:X2}", dialog.SelectedColor.R, dialog.SelectedColor.G, dialog.SelectedColor.B);

            if (existing != null)
            {
                existing.ReplacementHex = replacement;
                existing.Swatch.Background = BuildBrush(replacement);
                existing.Label.Text = originalHex + "  " + (char)0x2192 + "  " + replacement;
            }
        }

        private static Brush BuildBrush(string hex)
        {
            try
            {
                return new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
            }
            catch (Exception)
            {
                return Brushes.Transparent;
            }
        }

        private void InsertButton_Click(object sender, RoutedEventArgs e)
        {
            InsertSelectedFigure();
        }

        private void InsertSelectedFigure()
        {
            var row = SelectedFigure;
            if (row == null) return;

            var item = row.Item;
            string svg = _catalogService.LoadSvg(item);
            if (string.IsNullOrWhiteSpace(svg))
            {
                InsertFeedback.Text = "The artwork file could not be read.";
                return;
            }

            var options = new IllustrationInsertOptions
            {
                DropPlaceholders = true,
                Decompose = InsertModeBox.SelectedIndex == 1,
                DesiredWidth = ResolveInsertWidth()
            };

            foreach (var toggle in _paletteToggles)
            {
                if (string.IsNullOrWhiteSpace(toggle.ReplacementHex)) continue;
                if (string.Equals(toggle.OriginalHex, toggle.ReplacementHex, StringComparison.OrdinalIgnoreCase)) continue;
                options.ColorMap[toggle.OriginalHex] = toggle.ReplacementHex;
            }

            foreach (var toggle in _partToggles)
            {
                if (toggle.Box.IsChecked != true) options.HiddenPartIds.Add(toggle.PartId);
            }

            try
            {
                var result = _insertionService.Insert(_manager, item, svg, options, _catalogService.ResolveItemPath(item));
                InsertFeedback.Text = result.Message;
                if (!result.Inserted)
                {
                    MessageBox.Show(this, result.Message, "Insert illustration", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationLibraryWindow.Insert");
                InsertFeedback.Text = "Insertion failed: " + ex.Message;
            }
        }

        private float ResolveInsertWidth()
        {
            switch (InsertSizeBox.SelectedIndex)
            {
                case 0: return 160f;
                case 2: return 360f;
                default: return 240f;
            }
        }

        // ---------------------------------------------------------------- sources

        private void SourceList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ShowDetails();
        }

        private void ShowDetails()
        {
            var source = SelectedSource;
            if (source == null)
            {
                DetailName.Text = string.Empty;
                DetailStatus.Text = string.Empty;
                DetailLicense.Text = string.Empty;
                DetailAttribution.Text = string.Empty;
                DetailAcquisition.Text = string.Empty;
                DetailTechnical.Text = string.Empty;
                ImportButton.IsEnabled = false;
                OpenSiteButton.IsEnabled = false;
                OpenLicenseButton.IsEnabled = false;
                return;
            }

            DetailName.Text = source.Name;
            DetailStatus.Text = source.IsBundled
                ? source.StatusLabel + " - " + source.BundledFigureCount + " figures"
                : source.StatusLabel;
            DetailLicense.Text = "License: " + source.LicenseSummary;
            DetailAttribution.Text = source.AttributionRequired
                ? "Credit is required: keep the source's attribution wherever this artwork appears."
                : "No credit is required by the license.";
            DetailAcquisition.Text = source.AcquisitionNote;
            DetailTechnical.Text = source.TechnicalNote;

            OpenSiteButton.IsEnabled = !string.IsNullOrWhiteSpace(source.HomePageUrl);
            OpenLicenseButton.IsEnabled = !string.IsNullOrWhiteSpace(source.LicenseUrl);
            ImportButton.IsEnabled = true;
            ImportHint.Text = source.IsBundled
                ? "Files from this source already ship with the add-in; you can still import your own copies or a newer version."
                : "Import screens every file, refuses scripts, embedded images and external references, and records the source and license in " +
                  IllustrationImportService.CatalogFileName + ".";
        }

        private void OpenSiteButton_Click(object sender, RoutedEventArgs e)
        {
            OpenLink(SelectedSource == null ? null : SelectedSource.HomePageUrl, "Open website");
        }

        private void OpenLicenseButton_Click(object sender, RoutedEventArgs e)
        {
            OpenLink(SelectedSource == null ? null : SelectedSource.LicenseUrl, "Open license");
        }

        private void OpenLink(string url, string caption)
        {
            if (string.IsNullOrWhiteSpace(url)) return;

            try
            {
                IllustrationImportService.OpenUrl(url);
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationLibraryWindow.OpenLink");
                MessageBox.Show(this, "The link could not be opened: " + ex.Message, caption, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Directory.CreateDirectory(_importService.LibraryRoot);
                IllustrationImportService.OpenUrl(_importService.LibraryRoot);
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationLibraryWindow.OpenFolder");
                MessageBox.Show(this, "The import folder could not be opened: " + ex.Message, "Import folder", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            var source = SelectedSource;
            if (source == null) return;

            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Import illustrations from " + source.Name,
                Filter = "SVG files (*.svg)|*.svg|All files (*.*)|*.*",
                Multiselect = true,
                CheckFileExists = true
            };

            if (dialog.ShowDialog(this) != true) return;

            IllustrationImportResult result;
            try
            {
                result = _importService.Import(dialog.FileNames, source);
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "IllustrationLibraryWindow.Import");
                MessageBox.Show(this, "The import failed: " + ex.Message, "Import", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            SetStatus(string.Format(
                "Imported {0} of {1} file(s) from {2} into {3}{4}",
                result.Accepted,
                result.Total,
                source.Name,
                result.Folder,
                result.Skipped > 0 ? " (" + result.Skipped + " already present)" : string.Empty));

            _catalogService.Invalidate();
            LoadFigures();

            if (result.Messages.Count == 0) return;

            MessageBox.Show(
                this,
                "Some files were not imported:" + Environment.NewLine + Environment.NewLine + string.Join(Environment.NewLine, result.Messages),
                "Import finished",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        // ---------------------------------------------------------------- shell

        private void FiguresViewButton_Click(object sender, RoutedEventArgs e)
        {
            SetActiveView(true);
        }

        private void SourcesViewButton_Click(object sender, RoutedEventArgs e)
        {
            SetActiveView(false);
        }

        private void SetActiveView(bool figures)
        {
            FiguresView.Visibility = figures ? Visibility.Visible : Visibility.Collapsed;
            SourcesView.Visibility = figures ? Visibility.Collapsed : Visibility.Visible;
            NoticePanel.Visibility = Visibility.Visible;

            FiguresViewButton.Style = (Style)TryFindResource(figures ? "PrimaryButton" : "SecondaryButton");
            SourcesViewButton.Style = (Style)TryFindResource(figures ? "SecondaryButton" : "PrimaryButton");
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left) return;

            try
            {
                DragMove();
            }
            catch (InvalidOperationException)
            {
                // The button was released before the drag started.
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            FigureSearchBox.Focus();
            Keyboard.Focus(FigureSearchBox);
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.F && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                SetActiveView(true);
                FigureSearchBox.Focus();
                FigureSearchBox.SelectAll();
                e.Handled = true;
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void SetStatus(string message)
        {
            StatusText.Text = message;
        }

        /// <summary>List row of the picker; the template binds Name and DetailLine.</summary>
        private sealed class FigureRow
        {
            private ImageSource _thumbnail;
            private bool _thumbnailLoaded;

            public IllustrationItem Item { get; set; }

            public string Name { get; set; }

            public string DetailLine { get; set; }

            public string ThumbnailPath { get; set; }

            public ImageSource Thumbnail
            {
                get
                {
                    if (_thumbnailLoaded) return _thumbnail;
                    _thumbnailLoaded = true;
                    _thumbnail = IllustrationThumbnailRenderer.Create(ThumbnailPath);
                    return _thumbnail;
                }
            }
        }

        private sealed class PaletteToggle
        {
            public string OriginalHex { get; set; }

            public string ReplacementHex { get; set; }

            public Border Swatch { get; set; }

            public TextBlock Label { get; set; }
        }

        private sealed class PartToggle
        {
            public string PartId { get; set; }

            public CheckBox Box { get; set; }
        }
    }
}
