using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using oPenEfficiency.Services;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Linq;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class TagInspectorWindow : Window
    {
        public enum InspectorScope { CurrentSlide, SelectedShapes, Presentation }
        private readonly PowerPointManager _manager;

        public TagInspectorWindow(PowerPointManager manager, InspectorScope scope = InspectorScope.CurrentSlide)
        {
            InitializeComponent();
            _manager = manager;
            SetScope(scope);
            RefreshTags();
        }

        public void SetScope(InspectorScope scope)
        {
            RadioSlide.IsChecked = scope == InspectorScope.CurrentSlide;
            RadioShapes.IsChecked = scope == InspectorScope.SelectedShapes;
            RadioPresentation.IsChecked = scope == InspectorScope.Presentation;
            if (IsLoaded) RefreshTags();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Scope_Changed(object sender, RoutedEventArgs e)
        {
            if (IsLoaded) RefreshTags();
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            RefreshTags();
        }

        private void TagGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit)
            {
                var tagItem = e.Row.Item as TagItem;
                if (tagItem == null) return;

                var el = e.EditingElement as TextBox;
                if (el == null) return;

                string newValue = el.Text;

                try
                {
                    var app = _manager.GetApplication();
                    try { app.StartNewUndoEntry(); }
                    catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "TagInspectorWindow.AddTag.StartUndoEntry"); }
                    if (RadioPresentation.IsChecked == true)
                    {
                        app.ActivePresentation.Tags.Add(tagItem.Name, newValue);
                    }
                    else if (RadioSlide.IsChecked == true)
                    {
                        var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                        slide?.Tags.Add(tagItem.Name, newValue);
                    }
                    else if (RadioShapes.IsChecked == true)
                    {
                        var sel = app.ActiveWindow.Selection;
                        if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                        {
                            for (int i = 1; i <= sel.ShapeRange.Count; i++)
                            {
                                sel.ShapeRange[i].Tags.Add(tagItem.Name, newValue);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "TagInspectorWindow.UpdateTag");
                    MessageBox.Show("Error updating tag: " + ex.Message);
                }
            }
        }

        private void BtnStickyToTags_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var app = _manager.GetApplication();
                var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null) return;

                if (MessageBox.Show("Copy the text of all sticky notes on this slide to new slide tags? The sticky-note shapes are retained.",
                    "Copy Sticky Notes to Tags", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;

                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "TagInspectorWindow.Refresh.StartUndoEntry"); }

                int count = 0;
                int nextIndex = 1;
                var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 1; i <= slide.Tags.Count; i++) existingNames.Add(slide.Tags.Name(i));
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];
                    if (oPenEfficiency.Features.StickyNoteFeature.IsStickyNote(shape))
                    {
                        string text = "";
                        try { text = shape.TextFrame.TextRange.Text; }
                        catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "TagInspectorWindow.Refresh.ShapeText"); }

                        if (!string.IsNullOrEmpty(text))
                        {
                            string tagName;
                            do { tagName = $"oPE_StickyNoteText_{nextIndex++}"; } while (existingNames.Contains(tagName));
                            slide.Tags.Add(tagName, text);
                            existingNames.Add(tagName);
                            count++;
                        }
                    }
                }

                if (count > 0)
                {
                    RefreshTags();
                    MessageBox.Show($"Copied {count} sticky note(s) to new slide tags. The source shapes were retained.", "Copy Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("No sticky notes found on this slide.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TagInspectorWindow.CopyStickyNotesToTags");
                MessageBox.Show("Error converting sticky notes: " + ex.Message);
            }
        }

        private void BtnTagToSticky_Click(object sender, RoutedEventArgs e)
        {
            var selected = TagGrid.SelectedItem as TagItem;
            if (selected == null) return;

            try
            {
                if (oPenEfficiency.Features.StickyNoteFeature.Execute(_manager, selected.Value))
                {
                    // Optionally delete the tag after conversion? User didn't ask, so keep it.
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TagInspectorWindow.TagToStickyNote");
                MessageBox.Show("Error creating sticky note: " + ex.Message);
            }
        }

        private void RefreshTags()
        {
            try
            {
                var app = _manager.GetApplication();
                var tags = new List<TagItem>();

                if (RadioPresentation.IsChecked == true)
                {
                    var pres = app.ActivePresentation;
                    if (pres != null)
                    {
                        for (int i = 1; i <= pres.Tags.Count; i++)
                        {
                            tags.Add(new TagItem { Name = pres.Tags.Name(i), Value = pres.Tags.Value(i) });
                        }
                    }
                }
                else if (RadioSlide.IsChecked == true)
                {
                    var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                    if (slide != null)
                    {
                        for (int i = 1; i <= slide.Tags.Count; i++)
                        {
                            tags.Add(new TagItem { Name = slide.Tags.Name(i), Value = slide.Tags.Value(i) });
                        }
                    }
                }
                else if (RadioShapes.IsChecked == true)
                {
                    var sel = app.ActiveWindow.Selection;
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
                    {
                        var values = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                        for (int shapeIndex = 1; shapeIndex <= sel.ShapeRange.Count; shapeIndex++)
                        {
                            var shape = sel.ShapeRange[shapeIndex];
                            var shapeTags = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            for (int i = 1; i <= shape.Tags.Count; i++) shapeTags[shape.Tags.Name(i)] = shape.Tags.Value(i);
                            foreach (var name in values.Keys.ToList())
                                values[name].Add(shapeTags.TryGetValue(name, out string value) ? value : null);
                            foreach (var pair in shapeTags)
                                if (!values.ContainsKey(pair.Key))
                                    values[pair.Key] = Enumerable.Repeat<string>(null, shapeIndex - 1).Concat(new[] { pair.Value }).ToList();
                        }
                        tags.AddRange(values.OrderBy(pair => pair.Key).Select(pair => new TagItem
                        {
                            Name = pair.Key,
                            Value = pair.Value.Distinct(StringComparer.Ordinal).Count() == 1 && pair.Value[0] != null ? pair.Value[0] : "<mixed>"
                        }));
                    }
                }

                TagGrid.ItemsSource = tags;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TagInspectorWindow.RefreshTags");
                MessageBox.Show("Error reading tags: " + ex.Message);
            }
        }

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            string name = TxtNewTagName.Text.Trim();
            string val = TxtNewTagValue.Text.Trim();

            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show("Tag name cannot be empty.");
                return;
            }

            try
            {
                var app = _manager.GetApplication();
                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "TagInspectorWindow.DeleteTag.StartUndoEntry"); }
                if (RadioPresentation.IsChecked == true)
                {
                    app.ActivePresentation.Tags.Add(name, val);
                }
                else if (RadioSlide.IsChecked == true)
                {
                    var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                    slide?.Tags.Add(name, val);
                }
                else if (RadioShapes.IsChecked == true)
                {
                    var sel = app.ActiveWindow.Selection;
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        for (int i = 1; i <= sel.ShapeRange.Count; i++)
                        {
                            sel.ShapeRange[i].Tags.Add(name, val);
                        }
                    }
                }

                TxtNewTagName.Clear();
                TxtNewTagValue.Clear();
                RefreshTags();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TagInspectorWindow.AddTag");
                MessageBox.Show("Error adding tag: " + ex.Message);
            }
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            var selected = TagGrid.SelectedItem as TagItem;
            if (selected == null) return;

            string scope = RadioShapes.IsChecked == true ? " from every selected shape" : string.Empty;
            string warning = selected.Name.StartsWith("oPE_", StringComparison.OrdinalIgnoreCase)
                ? " This is an openEfficiency system tag; deleting it can disable related feature behavior." : string.Empty;
            if (MessageBox.Show($"Delete tag '{selected.Name}'{scope}?{warning}", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                return;

            try
            {
                var app = _manager.GetApplication();
                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "TagInspectorWindow.EditTag.StartUndoEntry"); }
                if (RadioPresentation.IsChecked == true)
                {
                    app.ActivePresentation.Tags.Delete(selected.Name);
                }
                else if (RadioSlide.IsChecked == true)
                {
                    var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                    slide?.Tags.Delete(selected.Name);
                }
                else if (RadioShapes.IsChecked == true)
                {
                    var sel = app.ActiveWindow.Selection;
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        for (int i = 1; i <= sel.ShapeRange.Count; i++)
                        {
                            try { sel.ShapeRange[i].Tags.Delete(selected.Name); }
                            catch (Exception ex) { ExceptionLogger.Log(ex, "TagInspectorWindow.DeleteShapeTag"); }
                        }
                    }
                }

                RefreshTags();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TagInspectorWindow.DeleteTag");
                MessageBox.Show("Error deleting tag: " + ex.Message);
            }
        }

        public class TagItem
        {
            public string Name { get; set; }
            public string Value { get; set; }
        }
    }
}
