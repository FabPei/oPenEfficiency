using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using oPenEfficiency.Services;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.UI.Dialogs
{
    public partial class TagInspectorWindow : Window
    {
        private readonly PowerPointManager _manager;

        public TagInspectorWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            this.Topmost = true;
            RefreshTags();
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
                    MessageBox.Show("Error updating tag: " + ex.Message);
                }
            }
        }

        private void BtnStickyToTags_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var app = _manager.GetApplication();
                
                var slides = new List<PowerPoint.Slide>();
                PowerPoint.SlideRange slideRange = null;
                try { slideRange = app.ActiveWindow.Selection.SlideRange; } catch { }
                
                if (slideRange != null && slideRange.Count > 0)
                {
                    for (int s = 1; s <= slideRange.Count; s++)
                        slides.Add(slideRange[s]);
                }
                else
                {
                    var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                    if (slide != null) slides.Add(slide);
                }

                if (slides.Count == 0) return;

                int totalCount = 0;
                foreach (var slide in slides)
                {
                    int count = 0;
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        var shape = slide.Shapes[i];
                        if (shape.Name.StartsWith("StickyNote", StringComparison.OrdinalIgnoreCase))
                        {
                            string text = "";
                            try { text = shape.TextFrame.TextRange.Text; } catch { }
                            
                            if (!string.IsNullOrEmpty(text))
                            {
                                count++;
                                slide.Tags.Add($"OE_STICKY_NOTE_{count}", text);
                            }
                        }
                    }
                    totalCount += count;
                }

                if (totalCount > 0)
                {
                    RefreshTags();
                    MessageBox.Show($"Converted {totalCount} sticky note(s) to slide tags.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("No sticky notes found on selected slide(s).", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
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
                    PowerPoint.SlideRange slideRange = null;
                    try { slideRange = app.ActiveWindow.Selection.SlideRange; } catch { }

                    if (slideRange != null && slideRange.Count > 0)
                    {
                        for (int s = 1; s <= slideRange.Count; s++)
                        {
                            var slide = slideRange[s];
                            string sourceStr = slideRange.Count > 1 ? $"Slide {slide.SlideIndex}" : "";
                            for (int i = 1; i <= slide.Tags.Count; i++)
                            {
                                tags.Add(new TagItem { Name = slide.Tags.Name(i), Value = slide.Tags.Value(i), Source = sourceStr, SlideId = slide.SlideID });
                            }
                        }
                    }
                    else
                    {
                        var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                        if (slide != null)
                        {
                            for (int i = 1; i <= slide.Tags.Count; i++)
                            {
                                tags.Add(new TagItem { Name = slide.Tags.Name(i), Value = slide.Tags.Value(i), Source = "", SlideId = slide.SlideID });
                            }
                        }
                    }
                }
                else if (RadioShapes.IsChecked == true)
                {
                    var sel = app.ActiveWindow.Selection;
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
                    {
                        // Use the first selected shape for inspection
                        var shape = sel.ShapeRange[1];
                        for (int i = 1; i <= shape.Tags.Count; i++)
                        {
                            tags.Add(new TagItem { Name = shape.Tags.Name(i), Value = shape.Tags.Value(i) });
                        }
                    }
                }

                TagGrid.ItemsSource = tags;
            }
            catch (Exception ex)
            {
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
                if (RadioPresentation.IsChecked == true)
                {
                    app.ActivePresentation.Tags.Add(name, val);
                }
                else if (RadioSlide.IsChecked == true)
                {
                    PowerPoint.SlideRange slideRange = null;
                    try { slideRange = app.ActiveWindow.Selection.SlideRange; } catch { }
                    
                    if (slideRange != null && slideRange.Count > 0)
                    {
                        for (int s = 1; s <= slideRange.Count; s++)
                            slideRange[s].Tags.Add(name, val);
                    }
                    else
                    {
                        var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                        slide?.Tags.Add(name, val);
                    }
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
                MessageBox.Show("Error adding tag: " + ex.Message);
            }
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            var selected = TagGrid.SelectedItem as TagItem;
            if (selected == null) return;

            if (MessageBox.Show($"Are you sure you want to delete tag '{selected.Name}'?", "Confirm Delete", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                return;

            try
            {
                var app = _manager.GetApplication();
                if (RadioPresentation.IsChecked == true)
                {
                    app.ActivePresentation.Tags.Delete(selected.Name);
                }
                else if (RadioSlide.IsChecked == true)
                {
                    if (selected.SlideId > 0)
                    {
                        var slide = app.ActivePresentation.Slides.FindBySlideID(selected.SlideId);
                        slide?.Tags.Delete(selected.Name);
                    }
                    else
                    {
                        PowerPoint.SlideRange slideRange = null;
                        try { slideRange = app.ActiveWindow.Selection.SlideRange; } catch { }

                        if (slideRange != null && slideRange.Count > 0)
                        {
                            for (int s = 1; s <= slideRange.Count; s++)
                                slideRange[s].Tags.Delete(selected.Name);
                        }
                        else
                        {
                            var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                            slide?.Tags.Delete(selected.Name);
                        }
                    }
                }
                else if (RadioShapes.IsChecked == true)
                {
                    var sel = app.ActiveWindow.Selection;
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        for (int i = 1; i <= sel.ShapeRange.Count; i++)
                        {
                            try { sel.ShapeRange[i].Tags.Delete(selected.Name); } catch { }
                        }
                    }
                }

                RefreshTags();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error deleting tag: " + ex.Message);
            }
        }

        public class TagItem
        {
            public string Name { get; set; }
            public string Value { get; set; }
            public string Source { get; set; }
            public int SlideId { get; set; }
        }
    }
}
