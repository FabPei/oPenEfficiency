using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using oPenEfficiency.Features.Tables;

namespace oPenEfficiency.UI.Dialogs
{
    /// <summary>
    /// Dialog for managing table branding styles.
    /// </summary>
    public partial class TableBrandingDialog : Window
    {
        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                this.DragMove();
        }
        public List<TableBrandingFeature.TableStyle> AvailableStyles { get; private set; }
        public TableBrandingFeature.TableStyle SelectedStyle { get; private set; }
        public string StyleToSave { get; private set; }
        public string Action { get; private set; } // "Apply" or "Save" or "Delete"

        public TableBrandingDialog()
        {
            InitializeComponent();
            AvailableStyles = TableBrandingFeature.LoadStyles();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadStylesList();
        }

        private void LoadStylesList()
        {
            StylesListBox.Items.Clear();

            if (AvailableStyles.Count == 0)
            {
                StylesListBox.Items.Add("No saved styles. Save a style from a table first.");
                StylesListBox.IsEnabled = false;
                ApplyButton.IsEnabled = false;
                DeleteButton.IsEnabled = false;
            }
            else
            {
                StylesListBox.IsEnabled = true;
                ApplyButton.IsEnabled = true;
                DeleteButton.IsEnabled = true;

                foreach (var style in AvailableStyles.Where(s => s != null).OrderBy(s => s.Name ?? string.Empty))
                {
                    var item = new ListBoxItem
                    {
                        Content = $"{style.Name} ({style.CreatedDate})",
                        Tag = style
                    };

                    // Show preview of colors
                    var stackPanel = new StackPanel { Orientation = Orientation.Horizontal };

                    if (style.HeaderFillColor.HasValue)
                    {
                        var headerColor = new Border
                        {
                            Width = 16,
                            Height = 16,
                            Background = ToBrush(style.HeaderFillColor.Value),
                            BorderBrush = System.Windows.Media.Brushes.Gray,
                            BorderThickness = new Thickness(1),
                            ToolTip = "Header Color"
                        };
                        stackPanel.Children.Add(headerColor);
                    }

                    if (style.BodyFillColor.HasValue)
                    {
                        var bodyColor = new Border
                        {
                            Width = 16,
                            Height = 16,
                            Margin = new Thickness(4, 0, 0, 0),
                            Background = ToBrush(style.BodyFillColor.Value),
                            BorderBrush = System.Windows.Media.Brushes.Gray,
                            BorderThickness = new Thickness(1),
                            ToolTip = "Body Color"
                        };
                        stackPanel.Children.Add(bodyColor);
                    }

                    if (style.AltRowFillColor.HasValue && style.HasAlternatingRows)
                    {
                        var altColor = new Border
                        {
                            Width = 16,
                            Height = 16,
                            Margin = new Thickness(4, 0, 0, 0),
                            Background = ToBrush(style.AltRowFillColor.Value),
                            BorderBrush = System.Windows.Media.Brushes.Gray,
                            BorderThickness = new Thickness(1),
                            ToolTip = "Alt Row Color"
                        };
                        stackPanel.Children.Add(altColor);
                    }

                    var textBlock = new TextBlock
                    {
                        Text = $"{style.Name}\n{style.CreatedDate}",
                        Margin = new Thickness(8, 0, 0, 0),
                        VerticalAlignment = VerticalAlignment.Center
                    };

                    var itemContainer = new StackPanel { Orientation = Orientation.Horizontal };
                    itemContainer.Children.Add(stackPanel);
                    itemContainer.Children.Add(textBlock);

                    item.Content = itemContainer;
                    StylesListBox.Items.Add(item);
                }
            }
        }

        private void StylesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (StylesListBox.SelectedItem is ListBoxItem selectedItem)
            {
                SelectedStyle = selectedItem.Tag as TableBrandingFeature.TableStyle;
            }
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedStyle != null)
            {
                Action = "Apply";
                DialogResult = true;
                Close();
            }
        }

        private void SaveNewButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(StyleNameTextBox.Text))
            {
                StyleToSave = StyleNameTextBox.Text.Trim();
                Action = "Save";
                DialogResult = true;
                Close();
            }
            else
            {
                MessageBox.Show(this, "Please enter a name for the style.", "Input Required",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                StyleNameTextBox.Focus();
            }
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedStyle != null)
            {
                var result = MessageBox.Show(this,
                    $"Are you sure you want to delete the style '{SelectedStyle.Name}'?",
                    "Confirm Delete",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    var previous = new List<TableBrandingFeature.TableStyle>(AvailableStyles);
                    AvailableStyles.RemoveAll(s => s.Name == SelectedStyle.Name);
                    if (!TableBrandingFeature.SaveStyles(AvailableStyles))
                    {
                        AvailableStyles = previous;
                        MessageBox.Show(this, "The style could not be deleted. The existing style file was left unchanged.", "Table Style Library", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                    LoadStylesList();
                    SelectedStyle = null;
                }
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Action = "Cancel";
            DialogResult = false;
            Close();
        }

        private void StylesListBox_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (SelectedStyle != null)
            {
                ApplyButton_Click(sender, e);
            }
        }

        private static System.Windows.Media.Brush ToBrush(uint officeRgb)
        {
            var color = System.Drawing.ColorTranslator.FromOle(unchecked((int)officeRgb));
            return new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(color.R, color.G, color.B));
        }
    }
}
