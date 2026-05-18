using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace oPenEfficiency
{
    public partial class AgendaWizard : Window
    {
        public readonly PowerPointManager _manager;
        public ObservableCollection<AgendaItem> Items { get; set; }
        public ObservableCollection<PreviewShape> PreviewShapes { get; set; } = new ObservableCollection<PreviewShape>();
        public ObservableCollection<PreviewShape> HighlightPreviewShapes { get; set; } = new ObservableCollection<PreviewShape>();
        private List<SavedAgenda> _savedAgendas = new List<SavedAgenda>();

        
        private void MovePhysicalSection(string topic, bool movingUp)
        {
            try
            {
                var pres = _manager.GetApplication().ActivePresentation;
                if (pres == null) return;

                int sectionIndex = -1;
                for (int i = 1; i <= pres.SectionProperties.Count; i++)
                {
                    if (pres.SectionProperties.Name(i) == topic)
                    {
                        sectionIndex = i;
                        break;
                    }
                }

                if (sectionIndex != -1)
                {
                    int targetIndex = movingUp ? Math.Max(1, sectionIndex - 1) : Math.Min(pres.SectionProperties.Count, sectionIndex + 1);
                    if (sectionIndex != targetIndex)
                    {
                        pres.SectionProperties.Move(sectionIndex, targetIndex);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MovePhysicalSection Error: " + ex.Message);
            }
        }

        public AgendaWizard(PowerPointManager manager)
        {
            _manager = manager;
            Items = new ObservableCollection<AgendaItem>();
            Items.CollectionChanged += Items_CollectionChanged;
            InitializeComponent();
            this.DataContext = this;
            AgendaListView.ItemsSource = Items;

            // Load saved agendas
            LoadSavedAgendas();

            // Try to pre-fill from existing agenda in presentation, otherwise use dummy data
            if (!TryPreFillFromPresentation())
            {
                Items.Add(new AgendaItem { Level = 0, Topic = "Welcome & Intro", Duration = 10 });
                Items.Add(new AgendaItem { Level = 0, Topic = "Main Strategy", Duration = 30 });
                Items.Add(new AgendaItem { Level = 1, Topic = "Q1 Review", Duration = 15 });
                Items.Add(new AgendaItem { Level = 1, Topic = "Q2 Outlook", Duration = 15 });
                Items.Add(new AgendaItem { Level = 0, Topic = "Wrap-up", Duration = 5 });
            }

            RecalculateAll();

            // Load Custom Layouts
            LoadLayouts();

            // Detect and auto-select active layout
            DetectActiveLayout();
        }

        private bool TryPreFillFromPresentation()
        {
            try
            {
                var app = _manager.GetApplication();
                var pres = app.ActivePresentation;
                if (pres == null) return false;

                // 0. Check for Efficient Elements Agenda
                if (TryExtractFromEfficientElements(pres)) return true;

                // 1. Check active slide
                var activeSlide = app.ActiveWindow.View.Slide as Slide;
                if (ExtractItemsFromSlide(activeSlide)) return true;

                // 2. If active slide has no agenda, scan for ANY agenda slide in the presentation
                for (int i = 1; i <= pres.Slides.Count; i++)
                {
                    var slide = pres.Slides[i];
                    if (!string.IsNullOrEmpty(slide.Tags["EE4P_AGENDA_DIVIDER"]) || !string.IsNullOrEmpty(slide.Tags["oPE_AgendaSlide"]))
                    {
                        if (ExtractItemsFromSlide(slide)) return true;
                    }
                }
            }
            catch { }
            return false;
        }

        private bool TryExtractFromEfficientElements(Presentation pres)
        {
            try
            {
                // EE4P marks presentations with an agenda with this tag
                string eeTag = pres.Tags["EE4P_AGENDAWIZARD"];
                if (string.IsNullOrEmpty(eeTag)) return false;

                // Collective list of items found
                var eeItems = new List<(int Level, string Topic, int Index)>();

                // Better approach for EE: Scan divider slides
                for (int i = 1; i <= pres.Slides.Count; i++)
                {
                    var slide = pres.Slides[i];
                    string content = slide.Tags["EE4P_AGENDAWIZARD_CONTENT"];

                    if (!string.IsNullOrEmpty(content) && content.StartsWith("/"))
                    {
                        // Some EE tags contain the full path /Topic 1/Topic 1.1
                        // We take the last part as the topic
                        var parts = content.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length > 0)
                        {
                            string topic = parts.Last().Trim();
                            // Remove leading numbers like "1. ", "1.1 "
                            topic = System.Text.RegularExpressions.Regex.Replace(topic, @"^[\d\.]+\s*", "");
                            
                            // Level can be inferred from number of parts - 1
                            int level = Math.Max(0, parts.Length - 1);
                            eeItems.Add((level, topic, i));
                        }
                    }
                }

                if (eeItems.Count > 0)
                {
                    Items.Clear();
                    foreach (var item in eeItems.OrderBy(x => x.Index))
                    {
                        Items.Add(new AgendaItem { Level = item.Level, Topic = item.Topic, Duration = 15 });
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("EE4P Extraction Error: " + ex.Message);
            }
            return false;
        }

        private bool ExtractItemsFromSlide(Slide slide)
        {
            if (slide == null) return false;

            // Check if this slide itself has EE tags
            string eeContent = slide.Tags["EE4P_AGENDAWIZARD_CONTENT"];
            if (!string.IsNullOrEmpty(eeContent))
            {
                var pres = _manager.GetApplication().ActivePresentation;
                if (pres != null && TryExtractFromEfficientElements(pres)) return true;
            }

            // Group all agenda data by GUID first
            var topicData = new Dictionary<string, (string Topic, string Resp, int Level, float Top, float Left)>();
            
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
            {
                try
                {
                    string role = shape.Tags["EE4P_AGENDAWIZARD"];
                    if (string.IsNullOrEmpty(role)) role = shape.Tags["oPE_AgendaRole"];

                    if (!string.IsNullOrEmpty(role))
                    {
                        string guid = "";
                        string[] parts = shape.Name.Split('_');
                        if (parts.Length >= 3) guid = parts[2];
                        else guid = "legacy_" + shape.Id;

                        if (!topicData.ContainsKey(guid))
                        {
                            topicData[guid] = (null, "", 0, shape.Top, shape.Left);
                        }

                        var current = topicData[guid];

                        if (role == "Topic" && shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            string text = shape.TextFrame.TextRange.Text;
                            text = System.Text.RegularExpressions.Regex.Replace(text, @"^[\d\.]+[.:]\s*", "");
                            current.Topic = text;
                            current.Top = shape.Top; // Use topic position for sorting
                            current.Left = shape.Left;
                        }
                        else if (role == "Responsible" && shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            current.Resp = shape.TextFrame.TextRange.Text;
                        }
                        
                        topicData[guid] = current;
                    }
                }
                catch { }
            }

            // Filter out entries that didn't have a Topic text and sort by Top position
            var validTopics = topicData.Values
                .Where(t => !string.IsNullOrEmpty(t.Topic))
                .OrderBy(t => t.Top)
                .ToList();

            if (validTopics.Count > 0)
            {
                Items.Clear();
                float? mainX = null;

                foreach (var data in validTopics)
                {
                    int level = 0;
                    if (!mainX.HasValue) mainX = data.Left;
                    else if (data.Left > mainX + 5) level = 1;

                    Items.Add(new AgendaItem 
                    { 
                        Level = level, 
                        Topic = data.Topic, 
                        Responsible = data.Resp,
                        Duration = 15 // Default duration
                    });
                }
                return true;
            }
            return false;
        }

        private void DetectActiveLayout()
        {
            try
            {
                var app = _manager.GetApplication();
                var pres = app.ActivePresentation;
                if (pres == null) return;

                // 1. Try to find the tag in the presentation or the active slide
                string activeLayoutName = "";
                try { activeLayoutName = pres.Tags["oPE_AgendaLayout"]; } catch { }

                var activeSlide = app.ActiveWindow.View.Slide as Slide;
                if (string.IsNullOrEmpty(activeLayoutName) && activeSlide != null)
                {
                    try { activeLayoutName = activeSlide.Tags["oPE_AgendaLayout"]; } catch { }
                }

                // 2. If no tag, try to find an existing agenda slide and see if we can match its style
                if (string.IsNullOrEmpty(activeLayoutName))
                {
                    // Check active slide first, then scan for any agenda divider slide
                    activeLayoutName = TryMatchLayoutStyleFromShapes(activeSlide);

                    if (string.IsNullOrEmpty(activeLayoutName))
                    {
                        for (int i = 1; i <= pres.Slides.Count; i++)
                        {
                            var s = pres.Slides[i];
                            if (!string.IsNullOrEmpty(s.Tags["EE4P_AGENDA_DIVIDER"]) || !string.IsNullOrEmpty(s.Tags["oPE_AgendaSlide"]))
                            {
                                activeLayoutName = TryMatchLayoutStyleFromShapes(s);
                                if (!string.IsNullOrEmpty(activeLayoutName)) break;
                            }
                        }
                    }
                }

                if (!string.IsNullOrEmpty(activeLayoutName))
                {
                    foreach (System.Windows.Controls.ComboBoxItem item in CboLayout.Items)
                    {
                        if (item.Content.ToString() == activeLayoutName)
                        {
                            CboLayout.SelectedItem = item;
                            
                            // If we found an active layout, it's highly likely we are updating
                            if (BtnOk != null) BtnOk.Content = "Update Agenda";
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("DetectActiveLayout Error: " + ex.Message);
            }
        }

        private string TryMatchLayoutStyleFromShapes(Slide slide)
        {
            if (slide == null) return null;

            // Find an oPE agenda shape (e.g., Topic)
            Microsoft.Office.Interop.PowerPoint.Shape topicShape = null;
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
            {
                try
                {
                    string role = shape.Tags["oPE_AgendaRole"];
                    if (string.IsNullOrEmpty(role)) role = shape.Tags["EE4P_AGENDAWIZARD"];

                    if (role == "Topic")
                    {
                        topicShape = shape;
                        break;
                    }
                }
                catch { }
            }

            if (topicShape == null) return null;

            // Compare against all loaded layouts in the ComboBox
            foreach (System.Windows.Controls.ComboBoxItem item in CboLayout.Items)
            {
                if (item.Tag is AgendaLayoutConfig layout && layout.ItemTextShape != null)
                {
                    // Compare key properties: Position and Size
                    float tolerance = 2.0f; // Slightly more generous tolerance for manual adjustments
                    if (Math.Abs(topicShape.Left - layout.ItemTextShape.Left) < tolerance &&
                        Math.Abs(topicShape.Top - layout.ItemTextShape.Top) < tolerance &&
                        Math.Abs(topicShape.Width - layout.ItemTextShape.Width) < tolerance &&
                        Math.Abs(topicShape.Height - layout.ItemTextShape.Height) < tolerance)
                    {
                        return layout.LayoutName;
                    }
                }
            }

            return null;
        }

        private void Items_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            // Subscribe to PropertyChanged for new items
            if (e.NewItems != null)
            {
                foreach (AgendaItem item in e.NewItems)
                {
                    item.PropertyChanged += AgendaItem_PropertyChanged;
                }
            }

            // Unsubscribe for removed items
            if (e.OldItems != null)
            {
                foreach (AgendaItem item in e.OldItems)
                {
                    item.PropertyChanged -= AgendaItem_PropertyChanged;
                }
            }
        }

        private void AgendaItem_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            // Recalculate time slots when Duration changes
            if (e.PropertyName == nameof(AgendaItem.Duration))
            {
                RecalculateAll();
            }
        }

        private void LoadLayouts()
        {
            try
            {
                CboLayout.Items.Clear();
                CboLayout.Items.Add(new System.Windows.Controls.ComboBoxItem { Content = "-- Select Layout --" });

                var layouts = JsonService.Load<List<AgendaLayoutConfig>>(@"AgendaLayouts\agenda_layouts.json");
                if (layouts != null)
                {
                    foreach (var layout in layouts)
                    {
                        CboLayout.Items.Add(new System.Windows.Controls.ComboBoxItem { Content = layout.LayoutName, Tag = layout });
                    }
                }
            }
            catch { }
        }

        private void LoadSavedAgendas()
        {
            try
            {
                string fileName = "saved_agendas.json";
                
                _savedAgendas = JsonService.Load<List<SavedAgenda>>(fileName) ?? new List<SavedAgenda>();
                
                CboSavedAgendas.Items.Clear();
                CboSavedAgendas.Items.Add(new System.Windows.Controls.ComboBoxItem { Content = "New Agenda", Tag = null });
                
                foreach (var agenda in _savedAgendas.OrderByDescending(a => a.CreatedDate))
                {
                    CboSavedAgendas.Items.Add(new System.Windows.Controls.ComboBoxItem 
                    { 
                        Content = $"{agenda.Name} ({agenda.CreatedDate:yyyy-MM-dd})", 
                        Tag = agenda 
                    });
                }
                
                CboSavedAgendas.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading saved agendas: {ex.Message}");
            }
        }

        private void CboSavedAgendas_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CboSavedAgendas.SelectedItem is System.Windows.Controls.ComboBoxItem selectedItem && selectedItem.Tag is SavedAgenda agenda)
            {
                // Load saved agenda
                TxtTitle.Text = agenda.Title;
                TxtSubtitle.Text = agenda.Subtitle;
                Items.Clear();

                foreach (var item in agenda.Items)
                {
                    Items.Add(new AgendaItem
                    {
                        Level = item.Level,
                        Topic = item.Topic,
                        Duration = item.Duration,
                        Format = item.Format,
                        ShouldGenerate = item.ShouldGenerate,
                        ShowMinutes = item.ShowMinutes
                    });
                }

                RecalculateAll();
            }
        }

        private void BtnDeleteAgenda_Click(object sender, RoutedEventArgs e)
        {
            if (CboSavedAgendas.SelectedItem is System.Windows.Controls.ComboBoxItem selectedItem && selectedItem.Tag is SavedAgenda agendaToDelete)
            {
                var result = System.Windows.Forms.MessageBox.Show(
                    $"Are you sure you want to delete the saved agenda '{agendaToDelete.Name}'?\n\nThis action cannot be undone.",
                    "Delete Agenda",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Warning);

                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    try
                    {
                        // Remove from list
                        _savedAgendas.RemoveAll(a => a.Name == agendaToDelete.Name);

                        // Save updated list
                        string appDataPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "oPenEfficiency", "SavedAgendas");
                        string fileName = "saved_agendas.json";

                        if (_savedAgendas.Count > 0)
                        {
                            JsonService.Save(_savedAgendas, fileName);
                        }
                        else
                        {
                            // Delete file if no agendas left
                            string filePath = System.IO.Path.Combine(appDataPath, fileName);
                            if (System.IO.File.Exists(filePath))
                            {
                                System.IO.File.Delete(filePath);
                            }
                        }

                        // Refresh dropdown and reset to "New Agenda"
                        LoadSavedAgendas();

                        // Clear form
                        TxtTitle.Text = "Agenda";
                        TxtSubtitle.Text = "";
                        Items.Clear();
                        Items.Add(new AgendaItem { Level = 0, Topic = "Welcome & Intro", Duration = 10 });
                        Items.Add(new AgendaItem { Level = 0, Topic = "Main Strategy", Duration = 30 });

                        System.Windows.Forms.MessageBox.Show("Agenda deleted successfully!", "Success", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error deleting agenda: {ex.Message}");
                        System.Windows.Forms.MessageBox.Show($"Error deleting agenda: {ex.Message}", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please select a saved agenda to delete.\n\nThe 'New Agenda' option cannot be deleted.", "No Agenda Selected", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        private void SaveCurrentAgenda()
        {
            try
            {
                var savedAgenda = new SavedAgenda
                {
                    Name = TxtTitle.Text ?? "Untitled Agenda",
                    Title = TxtTitle.Text ?? "Agenda",
                    Subtitle = TxtSubtitle.Text ?? "",
                    CreatedDate = DateTime.Now,
                    Items = Items.Select(item => new SavedAgendaItem
                    {
                        Level = item.Level,
                        Topic = item.Topic,
                        Duration = item.Duration,
                        Format = item.Format,
                        ShouldGenerate = item.ShouldGenerate,
                        ShowMinutes = item.ShowMinutes
                    }).ToList()
                };

                // Add to saved agendas list
                _savedAgendas.Add(savedAgenda);
                
                // Keep only last 20 saved agendas
                if (_savedAgendas.Count > 20)
                {
                    _savedAgendas = _savedAgendas.OrderByDescending(a => a.CreatedDate).Take(20).ToList();
                }

                // Save to file
                string appDataPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "oPenEfficiency", "SavedAgendas");
                string fileName = "saved_agendas.json";
                
                // Ensure directory exists
                if (!System.IO.Directory.Exists(appDataPath))
                {
                    System.IO.Directory.CreateDirectory(appDataPath);
                }

                JsonService.Save(_savedAgendas, fileName);
                
                // Refresh the dropdown
                LoadSavedAgendas();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving agenda: {ex.Message}");
            }
        }

        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (sender is ScrollViewer sv)
            {
                sv.ScrollToVerticalOffset(sv.VerticalOffset - e.Delta);
                e.Handled = true;
            }
        }

        private void ScrollViewer_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // Allow dragging window when clicking on non-interactive areas in the ScrollViewer
            if (e.ChangedButton == MouseButton.Left && e.OriginalSource is FrameworkElement source)
            {
                // Check if click is on or inside an interactive element
                var interactiveTypes = new[]
                {
                    typeof(ComboBox), typeof(TextBox), typeof(CheckBox), typeof(Button),
                    typeof(ToggleButton), typeof(ListBox), typeof(ListView), typeof(ComboBoxItem)
                };

                // Walk up the visual tree to check if click originated from an interactive element
                FrameworkElement current = source;
                while (current != null)
                {
                    foreach (var interactiveType in interactiveTypes)
                    {
                        if (interactiveType.IsInstanceOfType(current))
                        {
                            return; // Don't handle - let the interactive element process the click
                        }
                    }
                    current = VisualTreeHelper.GetParent(current) as FrameworkElement;
                }

                // If clicking on labels, borders, grids, or other non-interactive elements, allow drag
                if (source is TextBlock || source is Label || source is Border || source is Grid || source is StackPanel || source is DockPanel)
                {
                    if (e.ClickCount == 2)
                    {
                        // Double click - maximize/restore could go here
                    }
                    else
                    {
                        // Single click - start drag
                        this.DragMove();
                        e.Handled = true;
                    }
                }
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private System.Windows.Point _startPoint;

        private void AgendaListView_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Don't start drag if clicking on a TextBox or ToggleButton (allow editing/toggling)
            var originalSource = e.OriginalSource as FrameworkElement;
            if (originalSource is TextBox || originalSource is ToggleButton || originalSource is Button)
            {
                return;
            }

            _startPoint = e.GetPosition(null);
        }

        private void AgendaListView_MouseMove(object sender, MouseEventArgs e)
        {
            // Don't start drag if clicking on a TextBox or ToggleButton (allow editing/toggling)
            var originalSource = Mouse.DirectlyOver as FrameworkElement;
            if (originalSource is TextBox || originalSource is ToggleButton || originalSource is Button)
            {
                return;
            }

            System.Windows.Point mousePos = e.GetPosition(null);
            System.Windows.Vector diff = _startPoint - mousePos;

            if (e.LeftButton == MouseButtonState.Pressed &&
                (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                 Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance))
            {
                var listView = sender as ListView;
                if (listView?.SelectedItem is AgendaItem selectedItem)
                {
                    DataObject dragData = new DataObject("AgendaItemFormat", selectedItem);
                    DragDrop.DoDragDrop(listView, dragData, DragDropEffects.Move);
                }
            }
        }

        private void AgendaListView_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("AgendaItemFormat"))
            {
                AgendaItem itemToMove = e.Data.GetData("AgendaItemFormat") as AgendaItem;
                var listView = sender as ListView;

                // Get drop position
                System.Windows.Point dropPoint = e.GetPosition(listView);
                var hitTestResult = VisualTreeHelper.HitTest(listView, dropPoint);

                // Find the ListViewItem under the mouse
                ListViewItem droppedOnItem = null;
                if (hitTestResult != null)
                {
                    DependencyObject hitObj = hitTestResult.VisualHit;
                    while (hitObj != null && hitObj != listView)
                    {
                        if (hitObj is ListViewItem item)
                        {
                            droppedOnItem = item;
                            break;
                        }
                        hitObj = VisualTreeHelper.GetParent(hitObj);
                    }
                }

                int targetIndex = Items.Count; // Default to bottom
                if (droppedOnItem != null)
                {
                    targetIndex = listView.ItemContainerGenerator.IndexFromContainer(droppedOnItem);
                }

                if (itemToMove != null && targetIndex >= 0)
                {
                    int oldIndex = Items.IndexOf(itemToMove);
                    if (oldIndex != targetIndex)
                    {
                        Items.RemoveAt(oldIndex);
                        if (oldIndex < targetIndex) targetIndex--; // Adjust index after removal
                        Items.Insert(targetIndex, itemToMove);
                        MovePhysicalSection(itemToMove.Topic, oldIndex > targetIndex);
                        RecalculateAll();
                    }
                }
            }
        }

        
        private void CboLayout_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateLivePreview();
        }

        private void UpdateLivePreview()
        {
            if (PreviewShapes == null || HighlightPreviewShapes == null) return;
            PreviewShapes.Clear();
            HighlightPreviewShapes.Clear();
            
            AgendaLayoutConfig layout = null;
            if (CboLayout != null && CboLayout.SelectedItem is ComboBoxItem selectedLayoutItem && selectedLayoutItem.Tag is AgendaLayoutConfig config)
            {
                layout = config;
            }

            if (layout == null || layout.ItemTextShape == null) return;

            float currentY = layout.ItemTextShape.Top;
            float itemSpacing = layout.ItemSpacing;

            // 0. Full Preview - Slide Background (if assigned)
            if (layout.SlideShape != null && layout.SlideShape.HasFill)
            {
                PreviewShapes.Add(new PreviewShape {
                    X = layout.SlideShape.Left,
                    Y = layout.SlideShape.Top,
                    W = layout.SlideShape.Width,
                    H = layout.SlideShape.Height,
                    Text = "",
                    BgColor = GetBrush(layout.SlideShape.FillColorRgb, true),
                    FgColor = Brushes.Transparent,
                    FontSize = 1
                });
            }

            // 1. Full Preview - Title
            Brush titleColor = layout.HeadlineShape != null ? GetBrush(layout.HeadlineShape.FontColorRgb, true) : Brushes.Black;
            PreviewShapes.Add(new PreviewShape {
                X = layout.HeadlineShape?.Left ?? 50, 
                Y = layout.HeadlineShape?.Top ?? 30, 
                W = layout.HeadlineShape?.Width ?? 800, 
                H = layout.HeadlineShape?.Height ?? 50,
                Text = TxtTitle?.Text ?? "Agenda",
                BgColor = layout.HeadlineShape != null ? GetBrush(layout.HeadlineShape.FillColorRgb, layout.HeadlineShape.HasFill) : Brushes.Transparent,
                FgColor = titleColor,
                FontSize = layout.HeadlineShape?.FontSize ?? 32,
                FontWeight = (layout.HeadlineShape?.FontBold ?? true) ? FontWeights.Bold : FontWeights.Normal
            });

            // 1.1 Full Preview - Subtitle
            if (!string.IsNullOrEmpty(TxtSubtitle?.Text))
            {
                Brush subColor = layout.SubHeadlineShape != null ? GetBrush(layout.SubHeadlineShape.FontColorRgb, true) : Brushes.Gray;
                PreviewShapes.Add(new PreviewShape {
                    X = layout.SubHeadlineShape?.Left ?? 50, 
                    Y = layout.SubHeadlineShape?.Top ?? 85, 
                    W = layout.SubHeadlineShape?.Width ?? 800, 
                    H = layout.SubHeadlineShape?.Height ?? 30,
                    Text = TxtSubtitle.Text,
                    BgColor = layout.SubHeadlineShape != null ? GetBrush(layout.SubHeadlineShape.FillColorRgb, layout.SubHeadlineShape.HasFill) : Brushes.Transparent,
                    FgColor = subColor,
                    FontSize = layout.SubHeadlineShape?.FontSize ?? 18,
                    FontWeight = (layout.SubHeadlineShape?.FontBold ?? false) ? FontWeights.Bold : FontWeights.Normal
                });
            }

            // 2. Populating both previews
            int level0Index = 0;
            bool highlightItemAdded = false;

            foreach (var item in Items)
            {
                if (!item.ShouldGenerate) continue;
                
                bool isMain = item.Level == 0;
                if (isMain) level0Index++;

                // NO highlight in full preview as requested
                bool isActiveInFull = false; 

                // Add to Full Preview
                float rowHeight = AddItemToPreview(PreviewShapes, item, layout, currentY, isActiveInFull);
                currentY += rowHeight + itemSpacing;

                // Add to Highlight Preview (only once, use a dummy active state)
                if (isMain && !highlightItemAdded)
                {
                    AddItemToPreview(HighlightPreviewShapes, item, layout, 10, true);
                    highlightItemAdded = true;
                }
            }
        }

        private float AddItemToPreview(ObservableCollection<PreviewShape> collection, AgendaItem item, AgendaLayoutConfig layout, float y, bool isActive)
        {
            bool isMain = item.Level == 0;
            ShapeFormatConfig textConfig = isMain ? layout.ItemTextShape : (layout.SubItemTextShape ?? layout.ItemTextShape);
            ShapeFormatConfig numConfig = isMain ? layout.ItemNumberingShape : (layout.SubItemNumberingShape ?? layout.ItemNumberingShape);
            ShapeFormatConfig durConfig = isMain ? layout.ItemDurationShape : (layout.SubItemDurationShape ?? layout.ItemDurationShape);
            ShapeFormatConfig respConfig = isMain ? layout.ItemResponsibleShape : (layout.SubItemResponsibleShape ?? layout.ItemResponsibleShape);
            ShapeFormatConfig timeConfig = isMain ? layout.ItemTimeShape : (layout.SubItemTimeShape ?? layout.ItemTimeShape);
            ShapeFormatConfig bgConfig = isMain ? layout.ItemBackgroundShape : (layout.SubItemBackgroundShape ?? layout.ItemBackgroundShape);
            ShapeFormatConfig pageConfig = isMain ? layout.ItemPageNumberShape : (layout.SubItemPageNumberShape ?? layout.ItemPageNumberShape);

            // Calculate vertical offset to preserve layout's relative positioning
            float offsetY = y - textConfig.Top;

            // Determine preview text color
            Brush textColor = GetBrush(textConfig.FontColorRgb, true);
            
            // Determine preview highlight fill color
            Brush highlightFillBrush = null;
            if (isActive && layout.HighlightFillColor)
            {
                byte hr = (byte)(layout.HighlightFillColorRgb & 0xFF);
                byte hg = (byte)((layout.HighlightFillColorRgb >> 8) & 0xFF);
                byte hb = (byte)((layout.HighlightFillColorRgb >> 16) & 0xFF);
                highlightFillBrush = new SolidColorBrush(Color.FromRgb(hr, hg, hb));
                highlightFillBrush.Opacity = (collection == HighlightPreviewShapes) ? 0.7 : 0.5;
            }

            if (isActive)
            {
                if (layout.HighlightFontColor) textColor = GetBrush(layout.HighlightFontColorRgb, true);
                else textColor = new SolidColorBrush(Color.FromRgb(0, 120, 215)); // Default active blue
            }

            // If highlighting bold is on, use BOLD. Otherwise main items are SemiBold, sub-items Normal.
            FontWeight weight = isMain ? FontWeights.SemiBold : FontWeights.Normal;
            if (isActive && layout.HighlightBold) weight = FontWeights.Bold;
            FontStyle italic = (isActive && layout.HighlightItalic) ? FontStyles.Italic : FontStyles.Normal;

            // 1. Highlight Row Background (covers full width of agenda area)
            if (isActive && layout.HighlightFillColor)
            {
                // Calculate bounding box for highlight (must cover all possible shapes)
                float left = 50; // Standard slide margin
                if (numConfig != null) left = Math.Min(left, numConfig.Left - 5);
                if (timeConfig != null) left = Math.Min(left, timeConfig.Left - 5);
                
                float right = 910; // Standard right edge
                if (durConfig != null) right = Math.Max(right, durConfig.Left + durConfig.Width + 5);
                if (respConfig != null) right = Math.Max(right, respConfig.Left + respConfig.Width + 5);

                collection.Add(new PreviewShape {
                    X = left, Y = y - 2, W = right - left, H = textConfig.Height + 4,
                    Text = "", BgColor = highlightFillBrush, FgColor = Brushes.Transparent, FontSize = 1, FontWeight = weight
                });
            }
            else if (bgConfig != null && bgConfig.HasFill)
            {
                collection.Add(new PreviewShape {
                    X = bgConfig.Left, Y = bgConfig.Top + offsetY, W = bgConfig.Width, H = bgConfig.Height,
                    Text = "", BgColor = GetBrush(bgConfig.FillColorRgb, true), FgColor = Brushes.Transparent, FontSize = 1, FontWeight = weight
                });
            }

            // 2. Numbering
            if (!item.IsBreak && numConfig != null)
            {
                Brush numColor = isActive ? textColor : GetBrush(numConfig.FontColorRgb, true);
                collection.Add(new PreviewShape {
                    X = numConfig.Left, Y = numConfig.Top + offsetY, W = numConfig.Width, H = numConfig.Height,
                    Text = item.DisplayIndex, BgColor = Brushes.Transparent, FgColor = numColor, FontSize = numConfig.FontSize > 1 ? numConfig.FontSize : 14, FontWeight = weight, FontStyle = italic
                });
            }

            // 3. Time
            if (timeConfig != null)
            {
                Brush timeColor = isActive ? textColor : GetBrush(timeConfig.FontColorRgb, true);
                collection.Add(new PreviewShape {
                    X = timeConfig.Left, Y = timeConfig.Top + offsetY, W = timeConfig.Width, H = timeConfig.Height,
                    Text = item.TimeSlot, BgColor = Brushes.Transparent, FgColor = timeColor, FontSize = timeConfig.FontSize > 1 ? timeConfig.FontSize : 14, FontWeight = weight, FontStyle = italic
                });
            }

            // 4. Topic Text
            collection.Add(new PreviewShape {
                X = textConfig.Left, Y = textConfig.Top + offsetY, W = textConfig.Width, H = textConfig.Height,
                Text = item.Topic, BgColor = Brushes.Transparent, FgColor = textColor, FontSize = textConfig.FontSize > 1 ? textConfig.FontSize : 14, FontWeight = weight, FontStyle = italic
            });

            // 5. Duration
            if (item.ShowMinutes && durConfig != null)
            {
                Brush durColor = isActive ? textColor : GetBrush(durConfig.FontColorRgb, true);
                collection.Add(new PreviewShape {
                    X = durConfig.Left, Y = durConfig.Top + offsetY, W = durConfig.Width, H = durConfig.Height,
                    Text = $"{item.Duration} min", BgColor = Brushes.Transparent, FgColor = durColor, FontSize = durConfig.FontSize > 1 ? durConfig.FontSize : 12, FontWeight = FontWeights.Normal
                });
            }

            // 6. Responsible
            if (item.ShowResponsible && respConfig != null && !string.IsNullOrEmpty(item.Responsible))
            {
                Brush respColor = isActive ? textColor : GetBrush(respConfig.FontColorRgb, true);
                collection.Add(new PreviewShape {
                    X = respConfig.Left, Y = respConfig.Top + offsetY, W = respConfig.Width, H = respConfig.Height,
                    Text = item.Responsible, BgColor = Brushes.Transparent, FgColor = respColor, FontSize = respConfig.FontSize > 1 ? respConfig.FontSize : 12, FontWeight = FontWeights.Normal
                });
            }

            // 7. Page Number
            if (pageConfig != null)
            {
                Brush pageColor = isActive ? textColor : GetBrush(pageConfig.FontColorRgb, true);
                collection.Add(new PreviewShape {
                    X = pageConfig.Left, Y = pageConfig.Top + offsetY, W = pageConfig.Width, H = pageConfig.Height,
                    Text = "42", BgColor = Brushes.Transparent, FgColor = pageColor, FontSize = pageConfig.FontSize > 1 ? pageConfig.FontSize : 10, FontWeight = FontWeights.Normal
                });
            }

            return textConfig.Height;
        }

        private Brush GetBrush(int bgr, bool hasFill)
        {
            if (!hasFill) return Brushes.Transparent;
            byte r = (byte)(bgr & 0xFF);
            byte g = (byte)((bgr >> 8) & 0xFF);
            byte b = (byte)((bgr >> 16) & 0xFF);
            return new SolidColorBrush(Color.FromRgb(r, g, b));
        }

        private void RecalculateAll()
        {
            // Parse start time
            TimeSpan currentTime;
            if (!TimeSpan.TryParse(TxtStartTime?.Text, out currentTime))
            {
                currentTime = TimeSpan.FromHours(9); // Default to 09:00
            }

            int startNumber = 1;
            if (TxtStartNumber != null && int.TryParse(TxtStartNumber.Text, out int parsedStart))
            {
                startNumber = parsedStart;
            }

            string timeFormat = "HH:mm";
            if (CboTimeFormat != null && CboTimeFormat.SelectedItem is System.Windows.Controls.ComboBoxItem cbi)
            {
                timeFormat = cbi.Content.ToString();
                if (timeFormat == "HH:mm") timeFormat = "HH:mm";
                else if (timeFormat == "H:mm") timeFormat = "H:mm";
                else if (timeFormat == "HH:mm:ss") timeFormat = "HH:mm:ss";
                else if (timeFormat == "hh:mm tt") timeFormat = "hh:mm tt";
                else if (timeFormat == "h:mm tt") timeFormat = "h:mm tt";
            }

            // Track parent numbers at each level for hierarchical numbering
            int[] levelCounters = new int[20];
            levelCounters[0] = startNumber - 1;

            int flatIndex = 1;
            foreach (var item in Items)
            {
                if (!item.IsBreak)
                {
                    // Update level counters
                    levelCounters[item.Level]++;

                    // Reset counters for all deeper levels when we go up a level
                    for (int i = item.Level + 1; i < levelCounters.Length; i++)
                    {
                        levelCounters[i] = 0;
                    }

                    // Build hierarchical number
                    if (item.Level == 0)
                    {
                        item.DisplayIndex = levelCounters[0].ToString();
                    }
                    else
                    {
                        var numberParts = new System.Collections.Generic.List<string>();
                        for (int i = 0; i <= item.Level; i++)
                        {
                            if (levelCounters[i] > 0)
                            {
                                numberParts.Add(levelCounters[i].ToString());
                            }
                        }
                        item.DisplayIndex = string.Join(".", numberParts);
                    }
                }
                else
                {
                    item.DisplayIndex = "";
                }

                // Calculate time slot
                TimeSpan startTime = currentTime;
                TimeSpan endTime = currentTime + TimeSpan.FromMinutes(item.Duration);

                item.TimeSlot = $"{DateTime.Today.Add(startTime).ToString(timeFormat)} - {DateTime.Today.Add(endTime).ToString(timeFormat)}";

                // Move time forward for next item
                currentTime = endTime;

                flatIndex++;
            }
            UpdateLivePreview();
        }

        
        private void BtnDeleteInline_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is AgendaItem item)
            {
                Items.Remove(item);
                RecalculateAll();
            }
        }

        private void BtnIndentInline_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is AgendaItem item)
            {
                int index = Items.IndexOf(item);
                if (index > 0 && Items[index - 1].Level >= item.Level)
                {
                    item.Level++;
                    RecalculateAll();
                }
            }
        }

        private void BtnOutdentInline_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is AgendaItem item)
            {
                if (item.Level > 0)
                {
                    item.Level--;
                    RecalculateAll();
                }
            }
        }

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            var newItem = new AgendaItem { Level = 0, Topic = "New Item", Duration = 5 };

            if (AgendaListView.SelectedItem is AgendaItem selectedItem)
            {
                // Insert after the selected item
                int selectedIndex = Items.IndexOf(selectedItem);
                newItem.Level = selectedItem.Level; // Same level as selected
                Items.Insert(selectedIndex + 1, newItem);
            }
            else
            {
                // Add at the end if nothing selected
                Items.Add(newItem);
            }

            RecalculateAll();
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (AgendaListView.SelectedItem is AgendaItem selectedItem)
            {
                Items.Remove(selectedItem);
                RecalculateAll();
            }
        }

        private void BtnMoveUp_Click(object sender, RoutedEventArgs e)
        {
            if (AgendaListView.SelectedItem is AgendaItem selectedItem)
            {
                int index = Items.IndexOf(selectedItem);
                if (index > 0)
                {
                    Items.RemoveAt(index);
                    Items.Insert(index - 1, selectedItem);
                    MovePhysicalSection(selectedItem.Topic, true);
                    RecalculateAll();
                }
            }
        }

        private void BtnMoveDown_Click(object sender, RoutedEventArgs e)
        {
            if (AgendaListView.SelectedItem is AgendaItem selectedItem)
            {
                int index = Items.IndexOf(selectedItem);
                if (index < Items.Count - 1)
                {
                    Items.RemoveAt(index);
                    Items.Insert(index + 1, selectedItem);
                    MovePhysicalSection(selectedItem.Topic, false);
                    RecalculateAll();
                }
            }
        }

        private void BtnIndent_Click(object sender, RoutedEventArgs e)
        {
            if (AgendaListView.SelectedItem is AgendaItem selectedItem)
            {
                int selectedIndex = Items.IndexOf(selectedItem);

                // Cannot indent the first item
                if (selectedIndex == 0)
                {
                    return;
                }

                int targetLevel = selectedItem.Level + 1;

                // Check if the previous item can be a parent (its level must be at least targetLevel - 1)
                AgendaItem previousItem = Items[selectedIndex - 1];
                if (previousItem.Level < targetLevel - 1)
                {
                    // Cannot indent - no valid parent exists at the required level
                    return;
                }

                selectedItem.Level++;
                RecalculateAll();
            }
        }

        private void BtnOutdent_Click(object sender, RoutedEventArgs e)
        {
            if (AgendaListView.SelectedItem is AgendaItem selectedItem && selectedItem.Level > 0)
            {
                selectedItem.Level--;
                RecalculateAll();
            }
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Validate agenda items
                if (Items == null || Items.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "No agenda items found.\n\nPlease add at least one agenda item using the '+' button.",
                        "Validation Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // Get selected layout
                AgendaLayoutConfig layout = null;
                if (CboLayout.SelectedItem is ComboBoxItem selectedLayoutItem && selectedLayoutItem.Tag is AgendaLayoutConfig config)
                {
                    layout = config;
                }

                // MANDATORY: Creating an agenda without a selected layout shouldn't be possible.
                if (layout == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Please select a valid Agenda Layout from the dropdown menu before creating the agenda.\n\n" +
                        "If you haven't created a layout yet, use the 'Agenda Layout Creator' first.",
                        "Layout Required",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // Validate ItemTextShape dimensions
                if (layout.ItemTextShape.Width <= 0 || layout.ItemTextShape.Height <= 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        $"The selected layout has invalid dimensions for the Item Text shape.\n\n" +
                        $"Current values: Width={layout.ItemTextShape.Width}, Height={layout.ItemTextShape.Height}\n\n" +
                        $"Please edit the layout and assign a valid shape.",
                        "Invalid Layout Configuration",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Error);
                    return;
                }

                // Convert AgendaItems to AgendaItemData
                var agendaItems = new List<AgendaItemData>();
                foreach (var item in Items)
                {
                    if (item.ShouldGenerate)
                    {
                        agendaItems.Add(new AgendaItemData
                        {
                            Level = item.Level,
                            Topic = item.Topic,
                            Duration = item.Duration,
                            TimeSlot = item.TimeSlot,
                            Responsible = item.Responsible,
                            ShowMinutes = item.ShowMinutes,
                            ShowResponsible = item.ShowResponsible,
                        IsBreak = item.IsBreak
                        });
                    }
                }

                if (agendaItems.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "No agenda items are marked to generate.\n\n" +
                        "Please check the 'Show' toggle (⚙) for at least one agenda item in the list.",
                        "Validation Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // Validate item durations
                var itemsWithInvalidDuration = agendaItems.Where(i => i.Duration <= 0).ToList();
                if (itemsWithInvalidDuration.Count > 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        $"Found {itemsWithInvalidDuration.Count} agenda item(s) with invalid duration (0 or negative minutes).\n\n" +
                        "Please set a valid duration for all items.",
                        "Invalid Duration",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // Always create a NEW slide for the main agenda
                var presentation = _manager.GetApplication().ActivePresentation;
                if (presentation == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "No PowerPoint presentation is open.\n\n" +
                        "Please open an existing presentation or create a new one before generating the agenda.",
                        "No Presentation Open",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Error);
                    return;
                }

                var activeSlide = _manager.GetCurrentSlide();
                int insertIndex = (activeSlide != null) ? activeSlide.SlideIndex + 1 : presentation.Slides.Count + 1;
                
                var currentSlide = presentation.Slides.Add(insertIndex, PpSlideLayout.ppLayoutTitle);
                currentSlide.Tags.Add("oPE_AgendaSlide", "True");
                currentSlide.Tags.Add("oPE_AgendaMain", "True"); // Identify as main agenda slide
                currentSlide.Tags.Add("EE4P_AGENDAWIZARD_MAIN", "True"); // Compatibility tag

                // Generate agenda items on the new slide
                var generationService = new AgendaGenerationService(_manager, currentSlide, layout);
                bool createSections = ChkSections.IsChecked ?? false;
                bool createDividerSlides = ChkSeparatingSlides.IsChecked ?? false;
                bool highlightCurrent = ChkHighlightCurrent.IsChecked ?? false;
                generationService.GenerateFullAgenda(agendaItems, createSections, createDividerSlides, highlightCurrent);
                
                if (ChkBackupSlide.IsChecked == true)
                {
                    try {
                        var pres = _manager.GetApplication().ActivePresentation;
                        var backupSlide = pres.Slides.Add(pres.Slides.Count + 1, PpSlideLayout.ppLayoutTitle);
                        backupSlide.Tags.Add("EE4P_BACKUP_SLIDE", "True");
                        if (backupSlide.Shapes.HasTitle == Microsoft.Office.Core.MsoTriState.msoTrue) {
                            backupSlide.Shapes.Title.TextFrame.TextRange.Text = "Backup";
                        }

                        // Also create a section for the backup slide if sections are enabled
                        if (ChkSections.IsChecked == true)
                        {
                            bool sectionExists = false;
                            for (int s = 1; s <= pres.SectionProperties.Count; s++)
                            {
                                if (string.Equals(pres.SectionProperties.Name(s), "Backup", StringComparison.OrdinalIgnoreCase))
                                {
                                    sectionExists = true;
                                    break;
                                }
                            }

                            if (!sectionExists)
                            {
                                pres.SectionProperties.AddBeforeSlide(backupSlide.SlideIndex, "Backup");
                            }
                        }
                    } catch { }
                }

                this.Close();
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                // PowerPoint-specific error handling
                string errorMessage = comEx.Message;
                string errorTitle = "PowerPoint Error";

                if (errorMessage.Contains("out of range") || errorMessage.Contains("outside the range"))
                {
                    errorMessage = "A shape position or size value is out of PowerPoint's valid range.\n\n" +
                                   "This typically happens when:\n\n" +
                                   "• The layout configuration has invalid position/size values (e.g., negative or too large)\n" +
                                   "• The slide is full and cannot accommodate more shapes (maximum ~100 shapes per slide)\n" +
                                   "• The shape positioning exceeds PowerPoint's slide boundaries (typically 720x540 points for 4:3)\n\n" +
                                   "Solution: Check your layout configuration in the Agenda Layout Creator or try creating the agenda on a new, empty slide.";
                    errorTitle = "Shape Position Out of Range";
                }
                else if (errorMessage.Contains("Shape doesn't exist") || errorMessage.Contains("Shape not found"))
                {
                    errorMessage = "A referenced shape cannot be found on the slide.\n\n" +
                                   "This may happen if:\n\n" +
                                   "• The layout references a shape that was deleted from PowerPoint\n" +
                                   "• The shape ID stored in the layout is no longer valid\n\n" +
                                   "Solution: Re-create the layout by re-assigning all shapes in the Agenda Layout Creator.";
                    errorTitle = "Shape Not Found";
                }
                else if (errorMessage.Contains("Cannot insert") || errorMessage.Contains("Cannot add"))
                {
                    errorMessage = "PowerPoint cannot insert more shapes on this slide.\n\n" +
                                   "Possible reasons:\n\n" +
                                   "• The slide has reached the maximum number of shapes\n" +
                                   "• The presentation file is corrupted or in an invalid state\n\n" +
                                   "Solution: Try using a new presentation or reduce the number of shapes in your layout.";
                    errorTitle = "Cannot Add Shapes";
                }
                else if (errorMessage.Contains("Object does not exist") || errorMessage.Contains("unknown member"))
                {
                    errorMessage = "The Agenda Wizard lost track of a PowerPoint object (Slide or Section).\n\n" +
                                   "This typically happens when:\n\n" +
                                   "• Existing agenda slides were deleted manually while the wizard was open\n" +
                                   "• PowerPoint's internal selection state is out of sync\n" +
                                   "• A slide referenced by the wizard was moved or removed\n\n" +
                                   "Solution: Close the wizard, select the slide where you want the agenda, and try again.";
                    errorTitle = "PowerPoint Object Missing";
                }
                else
                {
                    errorMessage = $"The Agenda Wizard encountered a PowerPoint error:\n\n{errorMessage}\n\n" +
                                   "Possible causes:\n" +
                                   "• PowerPoint is busy or not responding\n" +
                                   "• The presentation is in a protected or restricted state\n" +
                                   "• Conflict with another PowerPoint add-in";
                }

                System.Windows.Forms.MessageBox.Show(errorMessage, errorTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            catch (InvalidOperationException invalidOpEx)
            {
                string errorMessage = invalidOpEx.Message;
                string errorTitle = "Configuration Error";

                if (invalidOpEx.InnerException != null)
                {
                    errorMessage += $"\n\nDetails: {invalidOpEx.InnerException.Message}";
                }

                System.Windows.Forms.MessageBox.Show(errorMessage, errorTitle, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                string detailedMessage = $"Error creating agenda: {ex.Message}\n\n" +
                                         $"Error Type: {ex.GetType().Name}\n\n" +
                                         "Common causes:\n" +
                                         "• The selected layout has invalid or incomplete configuration values\n" +
                                         "• PowerPoint is not responding or is in an unsupported state\n" +
                                         "• The presentation is protected, read-only, or corrupted\n" +
                                         "• Insufficient system resources to create additional shapes\n\n" +
                                         "Try: Using a different layout, restarting PowerPoint, or creating a new presentation.";

                System.Windows.Forms.MessageBox.Show(detailedMessage, "Unexpected Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void BtnSaveAgenda_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveCurrentAgenda();
                System.Windows.Forms.MessageBox.Show("Agenda saved successfully!", "Success", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error saving agenda: {ex.Message}", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e) => this.Close();

        
        private void BtnReadFromPres_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var app = _manager.GetApplication();
                var slide = app.ActiveWindow.View.Slide as Slide;
                
                if (slide != null)
                {
                    if (ExtractItemsFromSlide(slide))
                    {
                        RecalculateAll();
                        System.Windows.Forms.MessageBox.Show($"Successfully read agenda items from the current slide.", "Import Successful", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("No compatible agenda shapes found on the current slide.", "Import Failed", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error reading presentation: " + ex.Message);
            }
        }


        private void BtnImportLayouts_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var importWindow = new oPenEfficiency.UI.ImportAgendaLayoutWindow();
                importWindow.Owner = this;
                importWindow.ShowDialog();

                // Reload layouts after import
                LoadLayouts();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Import error: {ex.Message}");
            }
        }

        private void TxtStartTime_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            RecalculateAll();
        }

        private void TxtStartNumber_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            RecalculateAll();
        }

        private void CboTimeFormat_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            RecalculateAll();
        }

        private void TxtTitle_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateLivePreview();
        }

        private void TxtSubtitle_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateLivePreview();
        }

        private void BtnStatus_Click(object sender, RoutedEventArgs e)
        {
            // Placeholder for status button functionality
        }
    }

    // Model classes for saved agendas
    public class SavedAgenda
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public string Subtitle { get; set; }
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public List<SavedAgendaItem> Items { get; set; } = new List<SavedAgendaItem>();
    }

    public class SavedAgendaItem
    {
        public int Level { get; set; }
        public string Topic { get; set; }
        public int Duration { get; set; }
        public ShapeFormatConfig Format { get; set; }
        public bool ShouldGenerate { get; set; } = true;
        public bool ShowMinutes { get; set; } = true;
        public bool ShowResponsible { get; set; } = true;
    }

    public class AgendaItem : INotifyPropertyChanged
    {
        private string _displayIndex;
        private string _topic;
        private string _responsible;
        private int _duration;
        private string _timeSlot;
        private int _level;
        private ShapeFormatConfig _format;
        private bool _shouldGenerate = true;
        private bool _showMinutes = true;
        private bool _showResponsible = true;

        public string DisplayIndex { get => _displayIndex; set { _displayIndex = value; OnPropertyChanged(); } }
        public string Topic { get => _topic; set { _topic = value; OnPropertyChanged(); } }
        public string Responsible { get => _responsible; set { _responsible = value; OnPropertyChanged(); } }
        public int Duration { get => _duration; set { _duration = value; OnPropertyChanged(); } }
        public string TimeSlot { get => _timeSlot; set { _timeSlot = value; OnPropertyChanged(); } }
        public int Level { get => _level; set { _level = value; OnPropertyChanged(); OnPropertyChanged(nameof(IndentPadding)); } }
        public ShapeFormatConfig Format { get => _format; set { _format = value; OnPropertyChanged(); } }
        public bool ShouldGenerate { get => _shouldGenerate; set { _shouldGenerate = value; OnPropertyChanged(); } }
        public bool ShowMinutes { get => _showMinutes; set { _showMinutes = value; OnPropertyChanged(); } }
        public bool ShowResponsible { get => _showResponsible; set { _showResponsible = value; OnPropertyChanged(); } }
        private bool _isBreak = false;
        public bool IsBreak { get => _isBreak; set { _isBreak = value; OnPropertyChanged(); } }

        public Thickness IndentPadding => new Thickness(Level * 20, 0, 0, 0);

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class LevelToFontWeightConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is int level)
            {
                return level == 0 ? FontWeights.Bold : FontWeights.Normal;
            }
            return FontWeights.Normal;
        }
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) => throw new NotImplementedException();
    }

    public class PreviewShape
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double W { get; set; }
        public double H { get; set; }
        public string Text { get; set; }
        public Brush BgColor { get; set; }
        public Brush FgColor { get; set; }
        public double FontSize { get; set; }
        public FontWeight FontWeight { get; set; }
        public FontStyle FontStyle { get; set; }
    }
}
