using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI
{
    public partial class SaveAgendaLayoutWindow : Window
    {
        private PowerPointManager _manager;

        // Slide level shapes
        private ShapeFormatConfig _slideConfig;
        private ShapeFormatConfig _headlineConfig;
        private ShapeFormatConfig _subHeadlineConfig;

        // Main item shapes
        private ShapeFormatConfig _itemNumberingConfig;
        private ShapeFormatConfig _itemTextConfig;
        private ShapeFormatConfig _itemBackgroundConfig;
        private ShapeFormatConfig _itemTimeConfig;
        private ShapeFormatConfig _itemDurationConfig;
        private ShapeFormatConfig _itemResponsibleConfig;

        // Sub-item shapes (optional)
        private ShapeFormatConfig _subItemNumberingConfig;
        private ShapeFormatConfig _subItemTextConfig;
        private ShapeFormatConfig _subItemBackgroundConfig;
        private ShapeFormatConfig _subItemTimeConfig;
        private ShapeFormatConfig _subItemDurationConfig;
        private ShapeFormatConfig _subItemResponsibleConfig;

        private List<AgendaLayoutConfig> _existingLayouts = new List<AgendaLayoutConfig>();

        public SaveAgendaLayoutWindow(PowerPointManager manager)
        {
            InitializeComponent();
            _manager = manager;
            LoadExistingLayouts();

            // Initialize info panel as expanded by default
            InfoContent.Visibility = Visibility.Visible;
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void LoadExistingLayouts()
        {
            _existingLayouts = JsonService.Load<List<AgendaLayoutConfig>>(@"AgendaLayouts\agenda_layouts.json") ?? new List<AgendaLayoutConfig>();
            
            // Add "Create New Layout" as the first item
            var displayList = new List<AgendaLayoutConfig>();
            displayList.Add(new AgendaLayoutConfig { LayoutName = "-- Create New Layout --" });
            displayList.AddRange(_existingLayouts);

            CboExistingLayouts.ItemsSource = displayList;
            CboExistingLayouts.SelectedIndex = 0;
        }

        private Microsoft.Office.Interop.PowerPoint.Shape GetSingleSelectedShape()
        {
            var shapes = _manager?.GetSelectedShapes();
            if (shapes == null || shapes.Count != 1)
            {
                MessageBox.Show("Please select exactly ONE shape.", "Invalid Selection", MessageBoxButton.OK, MessageBoxImage.Warning);
                return null;
            }
            return shapes[1];
        }

        private bool IsShapeAlreadyAssigned(int shapeId, string currentField)
        {
            var assignedFields = new List<string>();

            // Slide level
            if (_slideConfig?.ShapeId == shapeId && currentField != "Slide") assignedFields.Add("Slide");
            if (_headlineConfig?.ShapeId == shapeId && currentField != "Headline") assignedFields.Add("Headline");
            if (_subHeadlineConfig?.ShapeId == shapeId && currentField != "Sub-headline") assignedFields.Add("Sub-headline");

            // Main item
            if (_itemNumberingConfig?.ShapeId == shapeId && currentField != "ItemNumbering") assignedFields.Add("Item Numbering");
            if (_itemTextConfig?.ShapeId == shapeId && currentField != "ItemText") assignedFields.Add("Item Text");
            if (_itemBackgroundConfig?.ShapeId == shapeId && currentField != "ItemBackground") assignedFields.Add("Item Background");
            if (_itemTimeConfig?.ShapeId == shapeId && currentField != "ItemTime") assignedFields.Add("Item Time");
            if (_itemDurationConfig?.ShapeId == shapeId && currentField != "ItemDuration") assignedFields.Add("Item Duration");
            if (_itemResponsibleConfig?.ShapeId == shapeId && currentField != "ItemResponsible") assignedFields.Add("Item Responsible");

            // Sub-item
            if (_subItemNumberingConfig?.ShapeId == shapeId && currentField != "SubItemNumbering") assignedFields.Add("Sub-Item Numbering");
            if (_subItemTextConfig?.ShapeId == shapeId && currentField != "SubItemText") assignedFields.Add("Sub-Item Text");
            if (_subItemBackgroundConfig?.ShapeId == shapeId && currentField != "SubItemBackground") assignedFields.Add("Sub-Item Background");
            if (_subItemTimeConfig?.ShapeId == shapeId && currentField != "SubItemTime") assignedFields.Add("Sub-Item Time");
            if (_subItemDurationConfig?.ShapeId == shapeId && currentField != "SubItemDuration") assignedFields.Add("Sub-Item Duration");
            if (_subItemResponsibleConfig?.ShapeId == shapeId && currentField != "SubItemResponsible") assignedFields.Add("Sub-Item Responsible");

            if (assignedFields.Count > 0)
            {
                MessageBox.Show($"This shape is already assigned to:\n\n{string.Join("\n", assignedFields)}\n\nA shape cannot be assigned to multiple fields.", "Duplicate Assignment", MessageBoxButton.OK, MessageBoxImage.Warning);
                return true;
            }
            return false;
        }

        private void SetAssigned(ComboBox comboBox, string shapeType)
        {
            comboBox.Items.Clear();
            comboBox.Items.Add(new ComboBoxItem { Content = $"Assigned ({shapeType})" });
            comboBox.SelectedIndex = 0;
        }

        private void ResetStatus(ComboBox comboBox)
        {
            comboBox.Items.Clear();
            comboBox.Items.Add(new ComboBoxItem { Content = "Not Assigned", Tag = "0" });
            comboBox.SelectedIndex = 0;
        }

        // Slide Level Assignment Handlers
        private void BtnAssignSlide_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "Slide")) return;
                _slideConfig = ExtractFormat(shape);
                SetAssigned(CboSlideStatus, "Slide");
            }
        }

        private void BtnClearSlide_Click(object sender, RoutedEventArgs e)
        {
            _slideConfig = null;
            ResetStatus(CboSlideStatus);
        }

        private void BtnAssignHeadline_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "Headline")) return;
                _headlineConfig = ExtractFormat(shape);
                SetAssigned(CboHeadlineStatus, "Headline");
            }
        }

        private void BtnClearHeadline_Click(object sender, RoutedEventArgs e)
        {
            _headlineConfig = null;
            ResetStatus(CboHeadlineStatus);
        }

        private void BtnAssignSubHeadline_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "Sub-headline")) return;
                _subHeadlineConfig = ExtractFormat(shape);
                SetAssigned(CboSubHeadlineStatus, "Sub-headline");
            }
        }

        private void BtnClearSubHeadline_Click(object sender, RoutedEventArgs e)
        {
            _subHeadlineConfig = null;
            ResetStatus(CboSubHeadlineStatus);
        }

        // Main Item Assignment Handlers
        private void BtnAssignItemNumbering_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "ItemNumbering")) return;
                _itemNumberingConfig = ExtractFormat(shape);
                SetAssigned(CboItemNumberingStatus, "Item #");
            }
        }

        private void BtnClearItemNumbering_Click(object sender, RoutedEventArgs e)
        {
            _itemNumberingConfig = null;
            ResetStatus(CboItemNumberingStatus);
        }

        private void BtnAssignItemText_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "ItemText")) return;
                _itemTextConfig = ExtractFormat(shape);
                SetAssigned(CboItemTextStatus, "Item Text");
            }
        }

        private void BtnClearItemText_Click(object sender, RoutedEventArgs e)
        {
            _itemTextConfig = null;
            ResetStatus(CboItemTextStatus);
        }

        private void BtnAssignItemBackground_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "ItemBackground")) return;
                _itemBackgroundConfig = ExtractFormat(shape);
                SetAssigned(CboItemBackgroundStatus, "Background");
            }
        }

        private void BtnClearItemBackground_Click(object sender, RoutedEventArgs e)
        {
            _itemBackgroundConfig = null;
            ResetStatus(CboItemBackgroundStatus);
        }

        private void BtnAssignItemTime_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "ItemTime")) return;
                _itemTimeConfig = ExtractFormat(shape);
                SetAssigned(CboItemTimeStatus, "Time");
            }
        }

        private void BtnClearItemTime_Click(object sender, RoutedEventArgs e)
        {
            _itemTimeConfig = null;
            ResetStatus(CboItemTimeStatus);
        }

        private void BtnAssignItemDuration_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "ItemDuration")) return;
                _itemDurationConfig = ExtractFormat(shape);
                SetAssigned(CboItemDurationStatus, "Duration");
            }
        }

        private void BtnClearItemDuration_Click(object sender, RoutedEventArgs e)
        {
            _itemDurationConfig = null;
            ResetStatus(CboItemDurationStatus);
        }

        private void BtnAssignItemResponsible_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "ItemResponsible")) return;
                _itemResponsibleConfig = ExtractFormat(shape);
                SetAssigned(CboItemResponsibleStatus, "Responsible");
            }
        }

        private void BtnClearItemResponsible_Click(object sender, RoutedEventArgs e)
        {
            _itemResponsibleConfig = null;
            ResetStatus(CboItemResponsibleStatus);
        }

        // Sub-Item Assignment Handlers
        private void BtnAssignSubItemNumbering_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "SubItemNumbering")) return;
                _subItemNumberingConfig = ExtractFormat(shape);
                SetAssigned(CboSubItemNumberingStatus, "Sub #");
            }
        }

        private void BtnClearSubItemNumbering_Click(object sender, RoutedEventArgs e)
        {
            _subItemNumberingConfig = null;
            ResetStatus(CboSubItemNumberingStatus);
        }

        private void BtnAssignSubItemText_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "SubItemText")) return;
                _subItemTextConfig = ExtractFormat(shape);
                SetAssigned(CboSubItemTextStatus, "Sub Text");
            }
        }

        private void BtnClearSubItemText_Click(object sender, RoutedEventArgs e)
        {
            _subItemTextConfig = null;
            ResetStatus(CboSubItemTextStatus);
        }

        private void BtnAssignSubItemBackground_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "SubItemBackground")) return;
                _subItemBackgroundConfig = ExtractFormat(shape);
                SetAssigned(CboSubItemBackgroundStatus, "Sub Background");
            }
        }

        private void BtnClearSubItemBackground_Click(object sender, RoutedEventArgs e)
        {
            _subItemBackgroundConfig = null;
            ResetStatus(CboSubItemBackgroundStatus);
        }

        private void BtnAssignSubItemTime_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "SubItemTime")) return;
                _subItemTimeConfig = ExtractFormat(shape);
                SetAssigned(CboSubItemTimeStatus, "Sub Time");
            }
        }

        private void BtnClearSubItemTime_Click(object sender, RoutedEventArgs e)
        {
            _subItemTimeConfig = null;
            ResetStatus(CboSubItemTimeStatus);
        }

        private void BtnAssignSubItemDuration_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "SubItemDuration")) return;
                _subItemDurationConfig = ExtractFormat(shape);
                SetAssigned(CboSubItemDurationStatus, "Sub Duration");
            }
        }

        private void BtnClearSubItemDuration_Click(object sender, RoutedEventArgs e)
        {
            _subItemDurationConfig = null;
            ResetStatus(CboSubItemDurationStatus);
        }

        private void BtnAssignSubItemResponsible_Click(object sender, RoutedEventArgs e)
        {
            var shape = GetSingleSelectedShape();
            if (shape != null)
            {
                if (IsShapeAlreadyAssigned(shape.Id, "SubItemResponsible")) return;
                _subItemResponsibleConfig = ExtractFormat(shape);
                SetAssigned(CboSubItemResponsibleStatus, "Sub Responsible");
            }
        }

        private void BtnClearSubItemResponsible_Click(object sender, RoutedEventArgs e)
        {
            _subItemResponsibleConfig = null;
            ResetStatus(CboSubItemResponsibleStatus);
        }



        private void BtnDeleteLayout_Click(object sender, RoutedEventArgs e)
        {
            if (CboExistingLayouts.SelectedItem == null || CboExistingLayouts.SelectedIndex == 0)
            {
                MessageBox.Show("Please select an existing layout to delete.", "Invalid Selection", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Toggle confirm button visibility
            if (BtnDeleteLayout.Visibility == Visibility.Visible)
            {
                BtnDeleteLayout.Visibility = Visibility.Collapsed;
                BtnDeleteLayoutConfirm.Visibility = Visibility.Visible;
            }
            else
            {
                // Reset button state
                BtnDeleteLayout.Visibility = Visibility.Visible;
                BtnDeleteLayoutConfirm.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnConfirmDeleteLayout_Click(object sender, RoutedEventArgs e)
        {
            if (CboExistingLayouts.SelectedItem is AgendaLayoutConfig layoutToDelete)
            {
                try
                {
                    string relativePath = @"AgendaLayouts\agenda_layouts.json";
                    var layouts = JsonService.Load<List<AgendaLayoutConfig>>(relativePath) ?? new List<AgendaLayoutConfig>();

                    int index = layouts.FindIndex(l => l.LayoutName == layoutToDelete.LayoutName);
                    if (index >= 0)
                    {
                        layouts.RemoveAt(index);
                        JsonService.Save(layouts, relativePath);

                        // Delete template file if exists
                        string appDataPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "oPenEfficiency", "AgendaLayouts");
                        string templatePath = System.IO.Path.Combine(appDataPath, $"{layoutToDelete.LayoutName}.pptx");
                        if (System.IO.File.Exists(templatePath))
                        {
                            System.IO.File.Delete(templatePath);
                        }

                        // Reload layouts and clear current config
                        LoadExistingLayouts();
                        BtnClearAll_Click(sender, e);

                        MessageBox.Show($"Layout '{layoutToDelete.LayoutName}' deleted successfully!", "Layout Deleted", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error deleting layout: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnLoadLayout_Click(object sender, RoutedEventArgs e)
        {
            // Reset delete button state when loading a layout
            BtnDeleteLayout.Visibility = Visibility.Visible;
            BtnDeleteLayoutConfirm.Visibility = Visibility.Collapsed;

            if (CboExistingLayouts.SelectedIndex == 0)
            {
                // "Create New Layout" selected - clear all
                BtnClearAll_Click(sender, e);
                return;
            }

            if (CboExistingLayouts.SelectedItem is AgendaLayoutConfig config)
            {
                // Load layout name
                TxtLayoutName.Text = config.LayoutName;

                // Load slide level shapes
                _slideConfig = config.SlideShape;
                _headlineConfig = config.HeadlineShape;
                _subHeadlineConfig = config.SubHeadlineShape;

                // Load main item shapes
                _itemNumberingConfig = config.ItemNumberingShape;
                _itemTextConfig = config.ItemTextShape;
                _itemBackgroundConfig = config.ItemBackgroundShape;
                _itemTimeConfig = config.ItemTimeShape;
                _itemDurationConfig = config.ItemDurationShape;
                _itemResponsibleConfig = config.ItemResponsibleShape;

                // Load sub-item shapes
                _subItemNumberingConfig = config.SubItemNumberingShape;
                _subItemTextConfig = config.SubItemTextShape;
                _subItemBackgroundConfig = config.SubItemBackgroundShape;
                _subItemTimeConfig = config.SubItemTimeShape;
                _subItemDurationConfig = config.SubItemDurationShape;
                _subItemResponsibleConfig = config.SubItemResponsibleShape;

                // Update UI status labels
                if (_slideConfig != null) SetAssigned(CboSlideStatus, "Slide"); else ResetStatus(CboSlideStatus);
                if (_headlineConfig != null) SetAssigned(CboHeadlineStatus, "Headline"); else ResetStatus(CboHeadlineStatus);
                if (_subHeadlineConfig != null) SetAssigned(CboSubHeadlineStatus, "Sub-headline"); else ResetStatus(CboSubHeadlineStatus);
                if (_itemNumberingConfig != null) SetAssigned(CboItemNumberingStatus, "Item #"); else ResetStatus(CboItemNumberingStatus);
                if (_itemTextConfig != null) SetAssigned(CboItemTextStatus, "Item Text"); else ResetStatus(CboItemTextStatus);
                if (_itemBackgroundConfig != null) SetAssigned(CboItemBackgroundStatus, "Background"); else ResetStatus(CboItemBackgroundStatus);
                if (_itemTimeConfig != null) SetAssigned(CboItemTimeStatus, "Time"); else ResetStatus(CboItemTimeStatus);
                if (_itemDurationConfig != null) SetAssigned(CboItemDurationStatus, "Duration"); else ResetStatus(CboItemDurationStatus);
                if (_itemResponsibleConfig != null) SetAssigned(CboItemResponsibleStatus, "Responsible"); else ResetStatus(CboItemResponsibleStatus);
                if (_subItemNumberingConfig != null) SetAssigned(CboSubItemNumberingStatus, "Sub #"); else ResetStatus(CboSubItemNumberingStatus);
                if (_subItemTextConfig != null) SetAssigned(CboSubItemTextStatus, "Sub Text"); else ResetStatus(CboSubItemTextStatus);
                if (_subItemBackgroundConfig != null) SetAssigned(CboSubItemBackgroundStatus, "Sub Background"); else ResetStatus(CboSubItemBackgroundStatus);
                if (_subItemTimeConfig != null) SetAssigned(CboSubItemTimeStatus, "Sub Time"); else ResetStatus(CboSubItemTimeStatus);
                if (_subItemDurationConfig != null) SetAssigned(CboSubItemDurationStatus, "Sub Duration"); else ResetStatus(CboSubItemDurationStatus);
                if (_subItemResponsibleConfig != null) SetAssigned(CboSubItemResponsibleStatus, "Sub Responsible"); else ResetStatus(CboSubItemResponsibleStatus);

                // Load settings
                foreach (ComboBoxItem item in CboListAlignment.Items)
                {
                    if (item.Content.ToString() == config.ListAlignment)
                    {
                        CboListAlignment.SelectedItem = item;
                        break;
                    }
                }
                foreach (ComboBoxItem item in CboListOrientation.Items)
                {
                    if (item.Content.ToString() == config.ListOrientation)
                    {
                        CboListOrientation.SelectedItem = item;
                        break;
                    }
                }
                TxtItemSpacing.Text = config.ItemSpacing.ToString();
                ChkAlignMiddle.IsChecked = config.AlignItemsMiddle;
                ChkAlignInfoShapes.IsChecked = config.AlignInfoShapesWithItem;
                
                // Load Highlight settings
                ChkHighlightBold.IsChecked = config.HighlightBold;
                ChkHighlightItalic.IsChecked = config.HighlightItalic;
                ChkHighlightUnderline.IsChecked = config.HighlightUnderline;
                
                ChkHighlightFontColor.IsChecked = config.HighlightFontColor;
                SetColorPreview(RectHighlightFontColor, config.HighlightFontColorRgb, true);

                ChkHighlightFillColor.IsChecked = config.HighlightFillColor;
                SetColorPreview(RectHighlightFillColor, config.HighlightFillColorRgb, config.HighlightFillColor);
            }
        }

        private void SetColorPreview(Border rect, int rgb, bool isActive)
        {
            if (!isActive)
            {
                rect.Background = Brushes.Transparent;
                return;
            }
            byte r = (byte)(rgb & 0xFF);
            byte g = (byte)((rgb >> 8) & 0xFF);
            byte b = (byte)((rgb >> 16) & 0xFF);
            rect.Background = new SolidColorBrush(Color.FromRgb(r, g, b));
        }

        private void BtnHighlightFontColor_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.ColorDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var color = dialog.Color;
                int rgb = color.R | (color.G << 8) | (color.B << 16);
                RectHighlightFontColor.Background = new SolidColorBrush(Color.FromRgb(color.R, color.G, color.B));
                ChkHighlightFontColor.IsChecked = true;
                // Store temporarily in Tag
                RectHighlightFontColor.Tag = rgb;
            }
        }

        private void BtnHighlightFillColor_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.ColorDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var color = dialog.Color;
                int rgb = color.R | (color.G << 8) | (color.B << 16);
                RectHighlightFillColor.Background = new SolidColorBrush(Color.FromRgb(color.R, color.G, color.B));
                ChkHighlightFillColor.IsChecked = true;
                // Store temporarily in Tag
                RectHighlightFillColor.Tag = rgb;
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            // Validate minimum required shapes
            if (_itemTextConfig == null)
            {
                MessageBox.Show("You must assign at least the Agenda Item Text shape.\n\nThis is the core shape that displays the agenda item text. Without it, the agenda cannot be generated.", "Missing Required Shape", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(TxtLayoutName.Text))
            {
                MessageBox.Show("Please enter a name for the layout.\n\nThe layout name is used to identify and load this configuration later.", "Missing Layout Name", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Validate shape positions and dimensions for all assigned shapes
            var validationErrors = ValidateShapeConfigurations();
            if (validationErrors.Count > 0)
            {
                MessageBox.Show(
                    "The following layout validation errors were found:\n\n" + string.Join("\n\n", validationErrors) + "\n\n" +
                    "Please adjust the shapes on your PowerPoint slide to have valid positions and dimensions.\n\n" +
                    "Typical valid ranges:\n" +
                    "• Position (Left/Top): 0 - 720 points (for 4:3 slides) or 0 - 960 points (for 16:9 slides)\n" +
                    "• Size (Width/Height): 10 - 500 points\n" +
                    "• Item Spacing: 1 - 100 points",
                    "Layout Validation Failed",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
                return;
            }

            // Validate item spacing
            float spacing = 15f;
            if (float.TryParse(TxtItemSpacing.Text, out float parsedSpacing))
            {
                spacing = parsedSpacing;
            }
            if (spacing < 1 || spacing > 100)
            {
                MessageBox.Show(
                    $"Item spacing value '{spacing}' is outside the valid range (1-100 points).\n\n" +
                    $"Current value: {spacing} pts ({spacing * 0.0352778:F2} cm | {spacing / 72.0:F2} in)\n\n" +
                    $"Please enter a value between 1 and 100.",
                    "Invalid Item Spacing",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
                return;
            }

            string layoutName = TxtLayoutName.Text.Trim();

            // Get list alignment
            string listAlignment = "Left";
            if (CboListAlignment.SelectedItem is ComboBoxItem alignmentItem)
            {
                listAlignment = alignmentItem.Content.ToString();
            }

            // Get list orientation
            string listOrientation = "Vertical";
            if (CboListOrientation.SelectedItem is ComboBoxItem orientationItem)
            {
                listOrientation = orientationItem.Content.ToString();
            }

            // Get align middle option
            bool alignItemsMiddle = ChkAlignMiddle.IsChecked == true;

            // Get align info shapes with item option
            bool alignInfoShapesWithItem = ChkAlignInfoShapes.IsChecked == true;

            // Get highlight format options
            bool highlightBold = ChkHighlightBold.IsChecked == true;
            bool highlightItalic = ChkHighlightItalic.IsChecked == true;
            bool highlightUnderline = ChkHighlightUnderline.IsChecked == true;
            bool highlightFontColor = ChkHighlightFontColor.IsChecked == true;
            bool highlightFillColor = ChkHighlightFillColor.IsChecked == true;

            int fontColorRgb = RectHighlightFontColor.Tag is int fRgb ? fRgb : 0; // Default black
            int fillColorRgb = RectHighlightFillColor.Tag is int bRgb ? bRgb : 16777215; // Default white/transparent

            AgendaLayoutConfig newLayout = new AgendaLayoutConfig
            {
                LayoutName = layoutName,

                // Slide level shapes
                SlideShape = _slideConfig,
                HeadlineShape = _headlineConfig,
                SubHeadlineShape = _subHeadlineConfig,

                // Main item shapes
                ItemNumberingShape = _itemNumberingConfig,
                ItemTextShape = _itemTextConfig,
                ItemBackgroundShape = _itemBackgroundConfig,
                ItemTimeShape = _itemTimeConfig,
                ItemDurationShape = _itemDurationConfig,
                ItemResponsibleShape = _itemResponsibleConfig,

                // Sub-item shapes (optional)
                SubItemNumberingShape = _subItemNumberingConfig,
                SubItemTextShape = _subItemTextConfig,
                SubItemBackgroundShape = _subItemBackgroundConfig,
                SubItemTimeShape = _subItemTimeConfig,
                SubItemDurationShape = _subItemDurationConfig,
                SubItemResponsibleShape = _subItemResponsibleConfig,

                // Settings
                ListAlignment = listAlignment,
                ListOrientation = listOrientation,
                ItemSpacing = spacing,
                AlignItemsMiddle = alignItemsMiddle,
                AlignInfoShapesWithItem = alignInfoShapesWithItem,
                HighlightBold = highlightBold,
                HighlightItalic = highlightItalic,
                HighlightUnderline = highlightUnderline,
                HighlightFontColor = highlightFontColor,
                HighlightFontColorRgb = fontColorRgb,
                HighlightFillColor = highlightFillColor,
                HighlightFillColorRgb = fillColorRgb
            };

            // Save to JSON using JsonService (consistent with other windows)
            var existingLayouts = JsonService.Load<List<AgendaLayoutConfig>>(@"AgendaLayouts\agenda_layouts.json") ?? new List<AgendaLayoutConfig>();
            existingLayouts.RemoveAll(l => l.LayoutName.Equals(layoutName, StringComparison.OrdinalIgnoreCase));
            existingLayouts.Add(newLayout);
            JsonService.Save(existingLayouts, @"AgendaLayouts\agenda_layouts.json");

            // Setup AppData Directory for backup/templates
            string appDataPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "oPenEfficiency", "AgendaLayouts");
            System.IO.Directory.CreateDirectory(appDataPath);

            // Export clean slide template
            ExportCleanSlideTemplate(layoutName, newLayout, appDataPath);

            MessageBox.Show($"Agenda Layout '{layoutName}' saved successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            this.Close();
        }

        private void ExportCleanSlideTemplate(string layoutName, AgendaLayoutConfig config, string backupDir)
        {
            var app = _manager.GetApplication();
            if (app == null || app.ActivePresentation == null || app.ActiveWindow == null || app.ActiveWindow.View.Slide == null)
                return;

            Slide activeSlide = (Microsoft.Office.Interop.PowerPoint.Slide)app.ActiveWindow.View.Slide;

            // Collect all shape IDs that were assigned (markers to remove)
            var idsToDelete = new List<int>();

            // Slide level
            if (config.SlideShape != null) idsToDelete.Add(config.SlideShape.ShapeId);
            if (config.HeadlineShape != null) idsToDelete.Add(config.HeadlineShape.ShapeId);
            if (config.SubHeadlineShape != null) idsToDelete.Add(config.SubHeadlineShape.ShapeId);

            // Main item shapes
            if (config.ItemNumberingShape != null) idsToDelete.Add(config.ItemNumberingShape.ShapeId);
            if (config.ItemTextShape != null) idsToDelete.Add(config.ItemTextShape.ShapeId);
            if (config.ItemBackgroundShape != null) idsToDelete.Add(config.ItemBackgroundShape.ShapeId);
            if (config.ItemTimeShape != null) idsToDelete.Add(config.ItemTimeShape.ShapeId);
            if (config.ItemDurationShape != null) idsToDelete.Add(config.ItemDurationShape.ShapeId);
            if (config.ItemResponsibleShape != null) idsToDelete.Add(config.ItemResponsibleShape.ShapeId);

            // Sub-item shapes
            if (config.SubItemNumberingShape != null) idsToDelete.Add(config.SubItemNumberingShape.ShapeId);
            if (config.SubItemTextShape != null) idsToDelete.Add(config.SubItemTextShape.ShapeId);
            if (config.SubItemBackgroundShape != null) idsToDelete.Add(config.SubItemBackgroundShape.ShapeId);
            if (config.SubItemTimeShape != null) idsToDelete.Add(config.SubItemTimeShape.ShapeId);
            if (config.SubItemDurationShape != null) idsToDelete.Add(config.SubItemDurationShape.ShapeId);
            if (config.SubItemResponsibleShape != null) idsToDelete.Add(config.SubItemResponsibleShape.ShapeId);

            // Create invisible temporary presentation
            Presentation tempPres = app.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoFalse);
            tempPres.PageSetup.SlideWidth = app.ActivePresentation.PageSetup.SlideWidth;
            tempPres.PageSetup.SlideHeight = app.ActivePresentation.PageSetup.SlideHeight;

            activeSlide.Copy();
            Slide pastedSlide = tempPres.Slides.Paste()[1];

            // Delete marker shapes
            for (int i = pastedSlide.Shapes.Count; i >= 1; i--)
            {
                var shape = pastedSlide.Shapes[i];
                if (idsToDelete.Contains(shape.Id))
                {
                    shape.Delete();
                }
            }

            string templatePath = System.IO.Path.Combine(backupDir, $"{layoutName}.pptx");
            tempPres.SaveAs(templatePath);
            tempPres.Close();
        }

        private void BtnToggleInfo_Click(object sender, RoutedEventArgs e)
        {
            if (InfoContent.Visibility == Visibility.Collapsed)
            {
                InfoContent.Visibility = Visibility.Visible;
                BtnToggleInfo.Content = "▼ Info";
            }
            else
            {
                InfoContent.Visibility = Visibility.Collapsed;
                BtnToggleInfo.Content = "+ Info";
            }
        }

        private void BtnClearAll_Click(object sender, RoutedEventArgs e)
        {
            // Reset all config fields
            _slideConfig = null;
            _headlineConfig = null;
            _subHeadlineConfig = null;
            _itemNumberingConfig = null;
            _itemTextConfig = null;
            _itemBackgroundConfig = null;
            _itemTimeConfig = null;
            _itemDurationConfig = null;
            _itemResponsibleConfig = null;
            _subItemNumberingConfig = null;
            _subItemTextConfig = null;
            _subItemBackgroundConfig = null;
            _subItemTimeConfig = null;
            _subItemDurationConfig = null;
            _subItemResponsibleConfig = null;

            // Reset all status labels
            ResetStatus(CboSlideStatus);
            ResetStatus(CboHeadlineStatus);
            ResetStatus(CboSubHeadlineStatus);
            ResetStatus(CboItemNumberingStatus);
            ResetStatus(CboItemTextStatus);
            ResetStatus(CboItemBackgroundStatus);
            ResetStatus(CboItemTimeStatus);
            ResetStatus(CboItemDurationStatus);
            ResetStatus(CboItemResponsibleStatus);
            ResetStatus(CboSubItemNumberingStatus);
            ResetStatus(CboSubItemTextStatus);
            ResetStatus(CboSubItemBackgroundStatus);
            ResetStatus(CboSubItemTimeStatus);
            ResetStatus(CboSubItemDurationStatus);
            ResetStatus(CboSubItemResponsibleStatus);

            // Reset settings
            TxtLayoutName.Text = "";
            CboListAlignment.SelectedIndex = 0;
            CboListOrientation.SelectedIndex = 0;
            TxtItemSpacing.Text = "15";
            ChkAlignMiddle.IsChecked = false;
            ChkAlignInfoShapes.IsChecked = false;
            ChkHighlightBold.IsChecked = false;
            ChkHighlightItalic.IsChecked = false;
            ChkHighlightUnderline.IsChecked = false;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Validates all assigned shape configurations for valid positions and dimensions.
        /// Returns a list of error messages for any invalid configurations.
        /// </summary>
        private List<string> ValidateShapeConfigurations()
        {
            var errors = new List<string>();

            // Validate each assigned shape configuration
            ValidateShapeConfig(_slideConfig, "Agenda Slide", errors);
            ValidateShapeConfig(_headlineConfig, "Agenda Headline", errors);
            ValidateShapeConfig(_subHeadlineConfig, "Agenda Sub-headline", errors);
            ValidateShapeConfig(_itemNumberingConfig, "Item Numbering", errors);
            ValidateShapeConfig(_itemTextConfig, "Item Text", errors);
            ValidateShapeConfig(_itemBackgroundConfig, "Item Background", errors);
            ValidateShapeConfig(_itemTimeConfig, "Item Time", errors);
            ValidateShapeConfig(_itemDurationConfig, "Item Duration", errors);
            ValidateShapeConfig(_itemResponsibleConfig, "Item Responsible", errors);
            ValidateShapeConfig(_subItemNumberingConfig, "Sub-Item Numbering", errors);
            ValidateShapeConfig(_subItemTextConfig, "Sub-Item Text", errors);
            ValidateShapeConfig(_subItemBackgroundConfig, "Sub-Item Background", errors);
            ValidateShapeConfig(_subItemTimeConfig, "Sub-Item Time", errors);
            ValidateShapeConfig(_subItemDurationConfig, "Sub-Item Duration", errors);
            ValidateShapeConfig(_subItemResponsibleConfig, "Sub-Item Responsible", errors);

            return errors;
        }

        /// <summary>
        /// Validates a single shape configuration for valid positions and dimensions.
        /// PowerPoint slides are typically 720x540 points (4:3) or 960x540 (16:9).
        /// </summary>
        private void ValidateShapeConfig(ShapeFormatConfig config, string fieldName, List<string> errors)
        {
            if (config == null) return; // Not assigned, skip validation

            const float MinPosition = 0f;
            const float MaxPosition = 1000f; // Generous upper limit to accommodate different slide sizes
            const float MinDimension = 5f;   // Minimum reasonable size
            const float MaxDimension = 500f; // Maximum reasonable size

            if (config.Left < MinPosition || config.Left > MaxPosition)
            {
                errors.Add($"{fieldName}: Left position ({config.Left:F1} pts) is outside valid range ({MinPosition}-{MaxPosition} pts). " +
                          $"Shape may be positioned off-slide to the left or too far right.");
            }

            if (config.Top < MinPosition || config.Top > MaxPosition)
            {
                errors.Add($"{fieldName}: Top position ({config.Top:F1} pts) is outside valid range ({MinPosition}-{MaxPosition} pts). " +
                          $"Shape may be positioned off-slide above or below.");
            }

            if (config.Width < MinDimension || config.Width > MaxDimension)
            {
                errors.Add($"{fieldName}: Width ({config.Width:F1} pts) is outside valid range ({MinDimension}-{MaxDimension} pts). " +
                          $"Shape may be too small to see or too large for the slide.");
            }

            if (config.Height < MinDimension || config.Height > MaxDimension)
            {
                errors.Add($"{fieldName}: Height ({config.Height:F1} pts) is outside valid range ({MinDimension}-{MaxDimension} pts). " +
                          $"Shape may be too small to see or too large for the slide.");
            }

            // Margin validation
            float hMargins = Math.Abs(config.MarginLeft) + Math.Abs(config.MarginRight);
            float vMargins = Math.Abs(config.MarginTop) + Math.Abs(config.MarginBottom);
            if (hMargins >= config.Width)
            {
                errors.Add($"{fieldName}: The sum of left and right margins ({hMargins:F1} pts) is greater than or equal to the shape width ({config.Width:F1} pts). " +
                          "Please reduce the text margins in PowerPoint for this shape.");
            }
            if (vMargins >= config.Height)
            {
                errors.Add($"{fieldName}: The sum of top and bottom margins ({vMargins:F1} pts) is greater than or equal to the shape height ({config.Height:F1} pts). " +
                          "Please reduce the text margins in PowerPoint for this shape.");
            }

            // Check if shape extends beyond typical slide boundaries
            float slideWidth = 960f; // Assume 16:9, will work for 4:3 too
            float slideHeight = 540f;
            if (config.Left + config.Width > slideWidth)
            {
                errors.Add($"{fieldName}: Shape extends beyond right edge of slide " +
                          $"(Left {config.Left:F1} + Width {config.Width:F1} = {config.Left + config.Width:F1} pts, slide width is ~{slideWidth} pts).");
            }

            if (config.Top + config.Height > slideHeight)
            {
                errors.Add($"{fieldName}: Shape extends beyond bottom edge of slide " +
                          $"(Top {config.Top:F1} + Height {config.Height:F1} = {config.Top + config.Height:F1} pts, slide height is ~{slideHeight} pts).");
            }
        }

        private static ShapeFormatConfig ExtractFormat(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            var config = new ShapeFormatConfig
            {
                ShapeId = shape.Id,
                Top = shape.Top,
                Left = shape.Left,
                Width = shape.Width,
                Height = shape.Height,
                HasFill = shape.Fill.Visible == Microsoft.Office.Core.MsoTriState.msoTrue,
                HasLine = shape.Line.Visible == Microsoft.Office.Core.MsoTriState.msoTrue
            };

            if (config.HasFill)
                config.FillColorRgb = shape.Fill.ForeColor.RGB;

            if (config.HasLine)
            {
                config.LineColorRgb = shape.Line.ForeColor.RGB;
                config.LineWeight = shape.Line.Weight;
            }

            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                var tr = shape.TextFrame.TextRange;
                
                // Capture font info (works even if text length is 0 in modern PowerPoint)
                config.FontName = tr.Font.Name;
                config.FontSize = tr.Font.Size;
                config.FontColorRgb = tr.Font.Color.RGB;
                config.FontBold = tr.Font.Bold == Microsoft.Office.Core.MsoTriState.msoTrue;
                config.FontItalic = tr.Font.Italic == Microsoft.Office.Core.MsoTriState.msoTrue;
                config.TextAlignment = (int)tr.ParagraphFormat.Alignment;

                // Handle cases where PowerPoint returns 0 or weird values for empty/mixed ranges
                if (config.FontSize <= 0 || config.FontSize > 1000)
                {
                    config.FontSize = 11; // Sensible default
                }

                config.MarginLeft = shape.TextFrame.MarginLeft;
                config.MarginRight = shape.TextFrame.MarginRight;
                config.MarginTop = shape.TextFrame.MarginTop;
                config.MarginBottom = shape.TextFrame.MarginBottom;

                config.Orientation = (int)shape.TextFrame.Orientation;
            }

            return config;
        }
    }
}
