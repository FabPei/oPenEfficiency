using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Features.Utilities;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI
{
    public enum ExcelLinkManagerStartAction
    {
        ShowAll,
        ShowBroken,
        RefreshAll,
        RelinkSource,
        OpenSource
    }

    public partial class ExcelLinkManagerWindow : Window
    {
        private readonly PowerPointManager _pm;
        public ObservableCollection<ExcelLinkItem> Links { get; } = new ObservableCollection<ExcelLinkItem>();
        /// <summary>
        /// Automation constructor: builds the window without a host so the selection contract can be
        /// driven on fixture rows (DL-LINKMGR-001). <see cref="LoadLinks"/> finds no application and
        /// leaves the collection empty.
        /// </summary>
        public ExcelLinkManagerWindow() : this(null) { }

        public ExcelLinkManagerWindow(
            PowerPointManager pm,
            ExcelLinkManagerStartAction startAction = ExcelLinkManagerStartAction.ShowAll)
        {
            InitializeComponent();
            _pm = pm;
            DataContext = this;
            LoadLinks();
            Loaded += (sender, args) => ApplyStartAction(startAction);
        }

        private void ApplyStartAction(ExcelLinkManagerStartAction action)
        {
            if (action == ExcelLinkManagerStartAction.ShowBroken)
            {
                ApplyBrokenFilter();
                return;
            }
            if (action == ExcelLinkManagerStartAction.RefreshAll)
            {
                RefreshLinks(Links.ToArray());
                return;
            }

            var firstNative = Links.FirstOrDefault(link => link.IsNativeTable);
            if (firstNative == null) return;
            firstNative.IsSelected = true;
            LinksGrid.SelectedItem = firstNative;
            if (action == ExcelLinkManagerStartAction.RelinkSource)
                Relink_Click(this, new RoutedEventArgs());
            else if (action == ExcelLinkManagerStartAction.OpenSource)
                OpenSource_Click(this, new RoutedEventArgs());
        }

        private void LoadLinks()
        {
            Links.Clear();
            try
            {
                var pres = _pm.GetApplication().ActivePresentation;
                foreach (Slide slide in pres.Slides)
                foreach (Shape shape in slide.Shapes)
                {
                    if (TableLink.TryParse(SafeAltText(shape), out var tableLink))
                    {
                        Links.Add(new ExcelLinkItem
                        {
                            SlideNumber = slide.SlideNumber, ShapeName = shape.Name, Kind = "Table",
                            SourcePath = tableLink.Path, Sheet = tableLink.Sheet, Range = tableLink.Range,
                            Mode = (tableLink.Formatted ? "Formatted" : "Values") + (tableLink.AutoResize ? " + Resize" : ""),
                            Status = Core.Services.ExcelLinkStatusPolicy.ResolveStatus(!string.IsNullOrWhiteSpace(tableLink.Path), File.Exists(tableLink.Path)), LastRefresh = ReadLastRefresh(shape),
                            ShapeRef = shape, IsNativeTable = true
                        });
                    }
                    else if (shape.Type == Office.MsoShapeType.msoLinkedOLEObject || shape.Type == Office.MsoShapeType.msoLinkedPicture)
                    {
                        string source = SafeLinkSource(shape);
                        bool isChart = shape.HasChart == Office.MsoTriState.msoTrue;
                        Links.Add(new ExcelLinkItem
                        {
                            SlideNumber = slide.SlideNumber,
                            ShapeName = shape.Name,
                            Kind = isChart ? "Linked Chart" : "OLE/Picture",
                            SourcePath = source,
                            Status = Core.Services.ExcelLinkStatusPolicy.ResolveStatus(!string.IsNullOrWhiteSpace(source), File.Exists(source)),
                            LastRefresh = "—",
                            ShapeRef = shape
                        });
                    }
                }
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ExcelLinkManagerWindow.LoadLinks"); }
        }

        private static string SafeAltText(Shape shape) { try { return shape.AlternativeText ?? ""; } catch { return ""; } }
        private static string SafeLinkSource(Shape shape) { try { return shape.LinkFormat?.SourceFullName ?? ""; } catch { return ""; } }
        private static string ReadLastRefresh(Shape shape)
        {
            try
            {
                string raw = shape.Tags[TableLinkFeature.LastRefreshTag];
                return Core.Services.ExcelLinkStatusPolicy.FormatLastRefresh(raw);
            }
            catch { return "—"; }
        }
        private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left) DragMove();
        }
        private void Close_Click(object sender, RoutedEventArgs e) => Close();
        /// <summary>
        /// The links the grid currently presents: the full collection, or the filtered view a start
        /// action installed.
        /// </summary>
        private ExcelLinkItem[] PresentedLinks()
        {
            var presented = LinksGrid.ItemsSource as IEnumerable<ExcelLinkItem>;
            return (presented ?? Links).ToArray();
        }

        /// <summary>
        /// Exact selected scope for every selected-scope command (Go to, Open Source, Relink Source,
        /// Unlink, Refresh Selected). The checkbox is the authoritative selection state and every
        /// row-selection change mirrors into it (<see cref="LinksGrid_SelectionChanged"/>), so a
        /// checked link counts once. Links the grid does not present are excluded: a link hidden by
        /// the "broken links" filter must never be refreshed, unlinked or relinked by a scope command
        /// the user cannot see (DL-LINKMGR-001).
        /// </summary>
        private ExcelLinkItem[] SelectedLinks()
        {
            var presented = PresentedLinks();
            var checkedLinks = presented.Where(l => l.IsSelected).ToArray();
            if (checkedLinks.Length > 0) return checkedLinks;
            return LinksGrid.SelectedItems.Cast<ExcelLinkItem>().Where(presented.Contains).ToArray();
        }

        /// <summary>Installs the "broken links only" view (the same filter the start action uses).</summary>
        private void ApplyBrokenFilter()
        {
            LinksGrid.ItemsSource = Links
                .Where(link => string.Equals(link.Status, "Missing", StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        private void LinksGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            foreach (ExcelLinkItem item in e.RemovedItems.OfType<ExcelLinkItem>())
                item.IsSelected = false;
            foreach (ExcelLinkItem item in e.AddedItems.OfType<ExcelLinkItem>())
                item.IsSelected = true;
        }

        private void LinkSelectionCheckBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var checkBox = sender as CheckBox;
            var item = checkBox?.DataContext as ExcelLinkItem;
            if (item == null) return;

            item.IsSelected = !item.IsSelected;
            if (item.IsSelected)
            {
                if (!LinksGrid.SelectedItems.Contains(item)) LinksGrid.SelectedItems.Add(item);
            }
            else
            {
                if (LinksGrid.SelectedItems.Contains(item)) LinksGrid.SelectedItems.Remove(item);
            }
            e.Handled = true;
        }

        private void Refresh_Click(object sender, RoutedEventArgs e) => RefreshLinks(SelectedLinks());
        private void RefreshAll_Click(object sender, RoutedEventArgs e) => RefreshLinks(Links.ToArray());

        private void RefreshLinks(ExcelLinkItem[] links)
        {
            foreach (var link in links)
            {
                try
                {
                    bool success = true;
                    if (link.IsNativeTable) success = TableLinkFeature.RefreshTable(_pm, link.ShapeRef);
                    else if (link.ShapeRef.Type == Office.MsoShapeType.msoLinkedOLEObject || link.ShapeRef.Type == Office.MsoShapeType.msoLinkedPicture) link.ShapeRef.LinkFormat.Update();
                    else if (link.ShapeRef.HasChart == Office.MsoTriState.msoTrue) link.ShapeRef.Chart.Refresh();
                    link.Status = success ? "Refreshed" : "Error";
                    if (success) link.LastRefresh = DateTime.Now.ToString("g");
                }
                catch (Exception ex) { link.Status = "Error"; ExceptionLogger.Log(ex, "ExcelLinkManagerWindow.RefreshLinks"); }
            }
        }

        private void BreakLink_Click(object sender, RoutedEventArgs e)
        {
            var selected = SelectedLinks().ToList();
            if (selected.Count == 0) return;

            var confirmation = MessageBox.Show(
                $"Remove the source connection from {selected.Count} selected link(s)?\n\nThe current PowerPoint content will be kept.",
                "Unlink Excel Sources",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Warning);
            if (confirmation != MessageBoxResult.OK) return;

            foreach (var link in selected)
            {
                try
                {
                    bool remove = link.IsNativeTable
                        ? TableLinkFeature.Unlink(link.ShapeRef)
                        : link.ShapeRef.Type == Office.MsoShapeType.msoLinkedOLEObject || link.ShapeRef.Type == Office.MsoShapeType.msoLinkedPicture;
                    if (!link.IsNativeTable && remove) link.ShapeRef.LinkFormat.BreakLink();
                    if (remove) Links.Remove(link);
                }
                catch (Exception ex) { link.Status = "Error"; ExceptionLogger.Log(ex, "ExcelLinkManagerWindow.BreakLink"); }
            }
        }

        private void OpenSource_Click(object sender, RoutedEventArgs e)
        {
            foreach (var link in SelectedLinks().Where(l => l.IsNativeTable))
                if (!TableLinkFeature.OpenSource(link.ShapeRef)) link.Status = "Error";
        }

        private void Relink_Click(object sender, RoutedEventArgs e)
        {
            foreach (var link in SelectedLinks().Where(l => l.IsNativeTable))
            {
                if (TableLinkFeature.RelinkSource(_pm, link.ShapeRef, this))
                {
                    if (TableLink.TryParse(SafeAltText(link.ShapeRef), out var updated))
                    {
                        link.SourcePath = updated.Path;
                        link.Sheet = updated.Sheet;
                        link.Range = updated.Range;
                    }
                    link.Status = "Refreshed";
                    link.LastRefresh = DateTime.Now.ToString("g");
                }
            }
        }

        // --- Bounded automation seam for the Link Manager selection contract (DL-LINKMGR-001) ---
        // The seam drives the real WPF machinery: DataGrid selection, the template CheckBox through
        // its toggle automation peer, and the same SelectedLinks() the command handlers use. It never
        // injects synthetic mouse input. Only the link rows are fixtures, because production rows come
        // from a live deck.

        /// <summary>Replaces the rows with fixtures (automation only).</summary>
        public void LoadFixtureLinksForAutomation(IEnumerable<ExcelLinkItem> items)
        {
            Links.Clear();
            foreach (var item in items) Links.Add(item);
            LinksGrid.ItemsSource = Links;
        }

        /// <summary>Selects rows through the DataGrid selection API - the programmatic equivalent of a
        /// row click, including the CurrentCell a real click would set.</summary>
        public void SelectRowsForAutomation(int[] rowIndexes, bool additive)
        {
            if (!additive) LinksGrid.UnselectAll();
            foreach (int index in rowIndexes)
            {
                if (index < 0 || index >= LinksGrid.Items.Count) continue;
                var item = LinksGrid.Items[index];
                LinksGrid.CurrentCell = new DataGridCellInfo(item, LinksGrid.Columns[0]);
                if (!LinksGrid.SelectedItems.Contains(item)) LinksGrid.SelectedItems.Add(item);
            }
            LinksGrid.UpdateLayout();
        }

        public void UnselectRowsForAutomation(int[] rowIndexes)
        {
            foreach (int index in rowIndexes)
            {
                if (index < 0 || index >= LinksGrid.Items.Count) continue;
                var item = LinksGrid.Items[index];
                if (LinksGrid.SelectedItems.Contains(item)) LinksGrid.SelectedItems.Remove(item);
            }
            LinksGrid.UpdateLayout();
        }

        /// <summary>Toggles the row's check box through its toggle automation peer (the path an
        /// accessibility tool uses, without mouse input). Returns the resulting checked state.</summary>
        public bool ToggleCheckboxForAutomation(int rowIndex)
        {
            var checkBox = FindCheckBoxForRow(rowIndex);
            if (checkBox == null) return false;
            var peer = new System.Windows.Automation.Peers.CheckBoxAutomationPeer(checkBox);
            var toggle = peer.GetPattern(System.Windows.Automation.Peers.PatternInterface.Toggle)
                as System.Windows.Automation.Provider.IToggleProvider;
            if (toggle == null) return false;
            toggle.Toggle();
            LinksGrid.UpdateLayout();
            return checkBox.IsChecked == true;
        }

        /// <summary>Extended multi-select over every presented row.</summary>
        public void SelectAllRowsForAutomation()
        {
            LinksGrid.SelectAll();
            LinksGrid.UpdateLayout();
        }

        /// <summary>Installs the broken-links filter view (automation only).</summary>
        public void ApplyBrokenFilterForAutomation() => ApplyBrokenFilter();

        /// <summary>Restores the unfiltered view (automation only).</summary>
        public void ShowAllRowsForAutomation() => LinksGrid.ItemsSource = Links;

        /// <summary>What a selected-scope command would receive, by identity: "index:checked" per row.</summary>
        public string DescribeSelectedScopeForAutomation()
        {
            return string.Join(",", SelectedLinks().Select(link => Links.IndexOf(link) + ":" + link.IsSelected));
        }

        public int PresentedRowCountForAutomation() => LinksGrid.Items.Count;

        public int CheckedButHiddenCountForAutomation()
        {
            var presented = PresentedLinks();
            return Links.Count(link => link.IsSelected && !presented.Contains(link));
        }

        public string DescribePresentedRowsForAutomation()
        {
            return string.Join(",", PresentedLinks().Select(link => Links.IndexOf(link) + ":" + link.ShapeName));
        }

        /// <summary>Runs the production "Refresh Selected" handler and reports the rows it touched by
        /// index, from the status the handler writes on a row it acted on.</summary>
        public string InvokeRefreshSelectedForAutomation()
        {
            var before = Links.Select(link => link.Status).ToArray();
            Refresh_Click(this, new RoutedEventArgs());
            return string.Join(",", Links.Select((link, index) =>
                link.Status != before[index] ? index.ToString() : null).Where(value => value != null));
        }

        private CheckBox FindCheckBoxForRow(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= LinksGrid.Items.Count) return null;
            var row = LinksGrid.ItemContainerGenerator.ContainerFromIndex(rowIndex) as DataGridRow;
            if (row == null)
            {
                LinksGrid.UpdateLayout();
                row = LinksGrid.ItemContainerGenerator.ContainerFromIndex(rowIndex) as DataGridRow;
            }
            return row == null ? null : FindDescendant<CheckBox>(row);
        }

        private static T FindDescendant<T>(System.Windows.DependencyObject parent) where T : System.Windows.DependencyObject
        {
            int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
                if (child is T match) return match;
                var nested = FindDescendant<T>(child);
                if (nested != null) return nested;
            }
            return null;
        }

        private void GoTo_Click(object sender, RoutedEventArgs e) => GoTo(SelectedLinks().FirstOrDefault());
        private void LinksGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e) => GoTo(LinksGrid.SelectedItem as ExcelLinkItem);
        private void GoTo(ExcelLinkItem link)
        {
            if (link == null) return;
            try
            {
                var app = _pm.GetApplication();
                app.ActiveWindow.View.GotoSlide(link.SlideNumber);
                link.ShapeRef.Select();
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ExcelLinkManagerWindow.GoTo"); }
        }
    }

    public class ExcelLinkItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected == value) return;
                _isSelected = value;
                Changed(nameof(IsSelected));
            }
        }
        public int SlideNumber { get; set; }
        public string ShapeName { get; set; }
        public string Kind { get; set; }
        private string _sourcePath;
        public string SourcePath { get => _sourcePath; set { _sourcePath = value; Changed(nameof(SourcePath)); } }
        private string _sheet;
        public string Sheet { get => _sheet; set { _sheet = value; Changed(nameof(Sheet)); } }
        private string _range;
        public string Range { get => _range; set { _range = value; Changed(nameof(Range)); } }
        public string Mode { get; set; }
        private string _status;
        public string Status { get => _status; set { _status = value; Changed(nameof(Status)); } }
        private string _lastRefresh;
        public string LastRefresh { get => _lastRefresh; set { _lastRefresh = value; Changed(nameof(LastRefresh)); } }
        public Shape ShapeRef { get; set; }
        public bool IsNativeTable { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;
        private void Changed(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
