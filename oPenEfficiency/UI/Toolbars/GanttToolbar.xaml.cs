using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI.Toolbars
{
    public partial class GanttToolbar : Window
    {
        private PowerPoint.Shape _targetShape;
        private oPenEfficiency.UI.Panels.GanttPanelViewModel _viewModel;
        private Guid _taskId;
        private bool _isUpdatingUI = false;
        private bool _isMilestone = false;

        public GanttToolbar(PowerPoint.Shape targetShape, oPenEfficiency.UI.Panels.GanttPanelViewModel viewModel)
        {
            InitializeComponent();
            _targetShape = targetShape;
            _viewModel = viewModel;

            string taskIdStr = PowerPointShapeManager.GetGanttId(_targetShape);
            string typeStr = PowerPointShapeManager.GetGanttType(_targetShape);

            _isMilestone = typeStr == PowerPointShapeManager.GanttShapeType.Milestone.ToString();

            TxtStyleHeader.Text = _isMilestone ? "Milestone style:" : "Task style:";

            if (Guid.TryParse(taskIdStr, out _taskId))
            {
                PopulateShapes();
                LoadTaskData();
            }
        }

        private void PopulateShapes()
        {
            CmbShape.Items.Clear();
            if (_isMilestone)
            {
                CmbShape.Items.Add(new ComboBoxItem { Content = "Diamond", Tag = MilestoneShape.Diamond.ToString() });
                CmbShape.Items.Add(new ComboBoxItem { Content = "Square", Tag = MilestoneShape.Square.ToString() });
                CmbShape.Items.Add(new ComboBoxItem { Content = "Star", Tag = MilestoneShape.Star.ToString() });
                CmbShape.Items.Add(new ComboBoxItem { Content = "Triangle", Tag = MilestoneShape.Triangle.ToString() });
                CmbShape.Items.Add(new ComboBoxItem { Content = "Circle", Tag = MilestoneShape.Circle.ToString() });
            }
            else
            {
                CmbShape.Items.Add(new ComboBoxItem { Content = "Rectangle", Tag = TaskShape.Rectangle.ToString() });
                CmbShape.Items.Add(new ComboBoxItem { Content = "Rounded", Tag = TaskShape.RoundedRectangle.ToString() });
                CmbShape.Items.Add(new ComboBoxItem { Content = "Right Arrow", Tag = TaskShape.RightArrow.ToString() });
                CmbShape.Items.Add(new ComboBoxItem { Content = "Chevron", Tag = TaskShape.Chevron.ToString() });
                CmbShape.Items.Add(new ComboBoxItem { Content = "Pentagon", Tag = TaskShape.Pentagon.ToString() });
            }
        }

        private void LoadTaskData()
        {
            _isUpdatingUI = true;
            if (_viewModel != null)
            {
                var task = _viewModel.Project.Tasks.FirstOrDefault(t => t.Id == _taskId);
                if (task != null)
                {
                    string currentShapeStr = _isMilestone ? task.MilestoneShape.ToString() : task.Shape.ToString();
                    foreach (ComboBoxItem item in CmbShape.Items)
                    {
                        if (item.Tag.ToString() == currentShapeStr)
                        {
                            CmbShape.SelectedItem = item;
                            break;
                        }
                    }
                }
            }
            _isUpdatingUI = false;
        }

        private void CmbShape_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingUI) return;
            if (CmbShape.SelectedItem is ComboBoxItem item)
            {
                string shapeTag = item.Tag.ToString();
                ApplyStyling(task => {
                    if (_isMilestone && Enum.TryParse(shapeTag, out MilestoneShape mShape))
                    {
                        task.MilestoneShape = mShape;
                    }
                    else if (!_isMilestone && Enum.TryParse(shapeTag, out TaskShape tShape))
                    {
                        task.Shape = tShape;
                    }
                });
            }
        }

        private void BtnColor_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.ColorDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string hex = $"#{dialog.Color.R:X2}{dialog.Color.G:X2}{dialog.Color.B:X2}";
                ApplyStyling(task => task.CustomColor = hex);
            }
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            ApplyStyling(task => {
                if (_isMilestone) task.MilestoneShape = MilestoneShape.Diamond;
                else task.Shape = TaskShape.Rectangle;
                task.CustomColor = null;
            });
            LoadTaskData();
        }

        private void ApplyStyling(Action<GanttTask> stylingAction)
        {
            if (_viewModel != null)
            {
                var task = _viewModel.Project.Tasks.FirstOrDefault(t => t.Id == _taskId);
                if (task != null)
                {
                    stylingAction(task);
                    _viewModel.RefreshChartCommand.Execute(null);
                }
            }
        }

        private void DragSurface_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed) DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape) return;
            e.Handled = true;
            Close();
        }
    }
}
