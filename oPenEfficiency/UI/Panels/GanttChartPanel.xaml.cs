using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using oPenEfficiency.Services;

namespace oPenEfficiency.UI.Panels
{
    public enum GanttStartAction { Open, CreateNew, UpdateSelected, ImportExcel, SyncExcel }

    /// <summary>
    /// Interaction logic for GanttChartPanel.xaml
    /// </summary>
    public partial class GanttChartPanel : UserControl
    {
        private Point _startPoint;
        private DataGridRow _draggedRow;

        public GanttChartPanel(PowerPointManager pptManager, GanttStartAction startAction = GanttStartAction.Open)
        {
            InitializeComponent();
            DataContext = new GanttPanelViewModel(pptManager);
            Loaded += (sender, args) => ApplyStartAction(startAction);
        }

        public void ApplyStartAction(GanttStartAction action)
        {
            var viewModel = DataContext as GanttPanelViewModel;
            if (viewModel == null) return;
            switch (action)
            {
                case GanttStartAction.CreateNew:
                    if (viewModel.CreateNewChartCommand.CanExecute(null)) viewModel.CreateNewChartCommand.Execute(null);
                    break;
                case GanttStartAction.UpdateSelected:
                    if (viewModel.RefreshChartCommand.CanExecute(null)) viewModel.RefreshChartCommand.Execute(null);
                    break;
                case GanttStartAction.ImportExcel:
                    if (viewModel.ImportExcelCommand.CanExecute(null)) viewModel.ImportExcelCommand.Execute(null);
                    break;
                case GanttStartAction.SyncExcel:
                    if (viewModel.SyncExcelCommand.CanExecute(null)) viewModel.SyncExcelCommand.Execute(null);
                    break;
            }
        }

        private void DataGridRow_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _startPoint = e.GetPosition(null);
            _draggedRow = sender as DataGridRow;
        }

        private void TasksDataGrid_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (DataContext is GanttPanelViewModel viewModel && viewModel.Project.IsLinkedToExcel) return;

            if (e.LeftButton == MouseButtonState.Pressed && _draggedRow != null)
            {
                Point position = e.GetPosition(null);
                if (Math.Abs(position.X - _startPoint.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(position.Y - _startPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    var task = _draggedRow.Item as oPenEfficiency.Models.GanttTask;
                    if (task != null)
                    {
                        DragDrop.DoDragDrop(_draggedRow, task, DragDropEffects.Move);
                    }
                }
            }
        }

        private void TasksDataGrid_DragOver(object sender, DragEventArgs e)
        {
            if (DataContext is GanttPanelViewModel viewModel && viewModel.Project.IsLinkedToExcel)
            {
                e.Effects = DragDropEffects.None;
            }
            else
            {
                e.Effects = DragDropEffects.Move;
            }
            e.Handled = true;
        }

        private void TasksDataGrid_Drop(object sender, DragEventArgs e)
        {
            if (!(DataContext is GanttPanelViewModel viewModel) || viewModel.Project.IsLinkedToExcel) return;

            if (e.Data.GetDataPresent(typeof(oPenEfficiency.Models.GanttTask)))
            {
                var droppedTask = e.Data.GetData(typeof(oPenEfficiency.Models.GanttTask)) as oPenEfficiency.Models.GanttTask;
                var targetRow = GetTargetRow(e.GetPosition(TasksDataGrid));

                if (droppedTask != null)
                {
                    int oldIndex = viewModel.Project.Tasks.IndexOf(droppedTask);
                    // Use Project.Tasks index (not Items/VisibleTasks index) so hidden rows don't shift the target position
                    int newIndex = targetRow != null
                        ? viewModel.Project.Tasks.IndexOf(targetRow.Item as oPenEfficiency.Models.GanttTask)
                        : viewModel.Project.Tasks.Count - 1;

                    if (oldIndex != newIndex && oldIndex != -1)
                    {
                        viewModel.ReorderTask(oldIndex, newIndex);
                    }
                }
            }
        }

        private DataGridRow GetTargetRow(Point pos)
        {
            for (int i = 0; i < TasksDataGrid.Items.Count; i++)
            {
                var row = TasksDataGrid.ItemContainerGenerator.ContainerFromIndex(i) as DataGridRow;
                if (row != null && IsMouseOverTarget(row, pos))
                {
                    return row;
                }
            }
            return null;
        }

        private bool IsMouseOverTarget(UIElement target, Point pos)
        {
            var bounds = VisualTreeHelper.GetDescendantBounds(target);
            var mousePos = TasksDataGrid.TranslatePoint(pos, target);
            return bounds.Contains(mousePos);
        }

        private void LinkExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog.ShowDialog() == true)
            {
                if (DataContext is GanttPanelViewModel viewModel)
                {
                    viewModel.LinkExcelFile(openFileDialog.FileName);
                }
            }
        }

        private void HelpBtn_Click(object sender, RoutedEventArgs e)
        {
            HelpOverlay.Visibility = Visibility.Visible;
        }

        private void CloseHelpBtn_Click(object sender, RoutedEventArgs e)
        {
            HelpOverlay.Visibility = Visibility.Collapsed;
        }

        private void HelpOverlay_MouseDown(object sender, MouseButtonEventArgs e)
        {
            HelpOverlay.Visibility = Visibility.Collapsed;
        }

        private void HelpCard_MouseDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
        }
    }

    /// <summary>
    /// Converts bool to Visibility (true=Collapsed, false=Visible).
    /// </summary>
    public class InverseBooleanToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool b && b) return Visibility.Collapsed;
            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }

    /// <summary>
    /// Converts OutlineLevel to horizontal margin for indentation.
    /// </summary>
    public class OutlineLevelToMarginConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int level)
            {
                return new Thickness(level * 15, 0, 0, 0);
            }
            return new Thickness(0);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }

    /// <summary>
    /// Converts IsSummary (bool) to Bold FontWeight if true.
    /// </summary>
    public class IsSummaryToFontWeightConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isSummary && isSummary)
            {
                return FontWeights.Bold;
            }
            return FontWeights.Normal;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }

    /// <summary>
    /// Inverts a boolean value.
    /// </summary>
    public class InverseBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool b) return !b;
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool b) return !b;
            return value;
        }
    }

    public class BooleanToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool b && b) return Visibility.Visible;
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }

    public class IsMilestoneToTypeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isMilestone && isMilestone) return "Milestone";
            return "Task";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }

    public class MilestoneShapeToUnicodeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is oPenEfficiency.Models.MilestoneShape shape)
            {
                switch (shape)
                {
                    case oPenEfficiency.Models.MilestoneShape.Diamond: return "♦";
                    case oPenEfficiency.Models.MilestoneShape.Square: return "■";
                    case oPenEfficiency.Models.MilestoneShape.Star: return "★";
                    case oPenEfficiency.Models.MilestoneShape.Triangle: return "▲";
                    case oPenEfficiency.Models.MilestoneShape.Circle: return "●";
                }
            }
            return value?.ToString();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }

    public class CollapseToggleConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => (value is bool b && b) ? "▶" : "▼";

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => Binding.DoNothing;
    }
}
