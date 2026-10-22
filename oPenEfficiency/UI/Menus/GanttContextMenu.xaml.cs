using System;
using System.Windows;
using System.Windows.Controls;
using oPenEfficiency.UI.Panels;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI.Menus
{
    public partial class GanttContextMenu : Window
    {
        private GanttPanelViewModel _viewModel;
        private oPenEfficiency.Models.GanttTask _targetTask;
        private int _insertionIndex;
        private DateTime _insertionDate;
        private DateTime? _shadeStart;
        private DateTime? _shadeEnd;
        private bool _isClosing;
        private DateTime _openedTime;

        public GanttContextMenu(GanttPanelViewModel viewModel, int index, DateTime date, bool isFlexibleMode, DateTime? shadeStart = null, DateTime? shadeEnd = null, oPenEfficiency.Models.GanttTask targetTask = null, bool isTaskLabel = false)
        {
            InitializeComponent();
            _openedTime = DateTime.Now;

            _viewModel = viewModel;
            _targetTask = targetTask;
            _insertionIndex = index;
            _insertionDate = date;
            _shadeStart = shadeStart;
            _shadeEnd = shadeEnd;

            this.Loaded += (s, e) =>
            {
                this.Activate();
                this.Focus();
                this.Topmost = true;
                oPenEfficiency.ThisAddIn.LogDebugStatic("GanttContextMenu Loaded and Focus requested.");
            };

            // Only show Toggle Shade if we have a range (clicked in header)
            if (shadeStart == null || shadeEnd == null)
            {
                BtnToggleShade.Visibility = Visibility.Collapsed;
                ShadeSeparator.Visibility = Visibility.Collapsed;
            }

            // Only show Indent/Outdent when right-clicking the task name label (RowAnchor)
            if (!isTaskLabel)
            {
                IndentSeparator.Visibility = Visibility.Collapsed;
                BtnIndent.Visibility = Visibility.Collapsed;
                BtnOutdent.Visibility = Visibility.Collapsed;
            }
            else
            {
                BtnOutdent.IsEnabled = _targetTask != null && _targetTask.OutlineLevel > 0;
            }

            // Only show Highlight Row if a task is targeted
            if (_targetTask == null)
            {
                HighlightSeparator.Visibility = Visibility.Collapsed;
                BtnHighlightRow.Visibility = Visibility.Collapsed;
            }
            else
            {
                TxtHighlightRow.Text = _targetTask.IsHighlighted ? "Remove Row Highlight" : "Highlight Row";
            }

            // Only show Task shapes if a non-milestone task is targeted
            if (_targetTask == null || _targetTask.IsMilestone)
            {
                TaskShapeSeparator.Visibility = Visibility.Collapsed;
                TxtTaskShapeHeader.Visibility = Visibility.Collapsed;
                PanelTaskShapes.Visibility = Visibility.Collapsed;
            }

            // Only show Milestone shapes if a milestone is targeted
            if (_targetTask == null || !_targetTask.IsMilestone)
            {
                MilestoneSeparator.Visibility = Visibility.Collapsed;
                TxtMilestoneHeader.Visibility = Visibility.Collapsed;
                PanelMilestoneShapes.Visibility = Visibility.Collapsed;
            }

            // Only show Roadmap-specific options if Gantt Rules are disabled
            if (!isFlexibleMode)
            {
                BtnAddSubMilestone.Visibility = Visibility.Collapsed;
                BtnAddSubBar.Visibility = Visibility.Collapsed;
                RoadmapSeparator1.Visibility = Visibility.Collapsed;
                RoadmapSeparator2.Visibility = Visibility.Collapsed;
            }
        }

        private void SetTaskShape_Click(object sender, RoutedEventArgs e)
        {
            if (_isClosing || _targetTask == null || _targetTask.IsMilestone) return;
            _isClosing = true;

            if (sender is Button btn && btn.Tag != null)
            {
                string shapeName = btn.Tag.ToString();
                if (Enum.TryParse(shapeName, out oPenEfficiency.Models.TaskShape shape))
                {
                    _targetTask.Shape = shape;
                    _viewModel.RefreshChartCommand.Execute(null);
                }
            }
            this.Close();
        }

        private void SetMilestoneShape_Click(object sender, RoutedEventArgs e)
        {
            if (_isClosing || _targetTask == null || !_targetTask.IsMilestone) return;
            _isClosing = true;

            if (sender is Button btn && btn.Tag != null)
            {
                string shapeName = btn.Tag.ToString();
                if (Enum.TryParse(shapeName, out oPenEfficiency.Models.MilestoneShape shape))
                {
                    _targetTask.MilestoneShape = shape;
                    _viewModel.RefreshChartCommand.Execute(null);
                }
            }
            this.Close();
        }

        private void ToggleShade_Click(object sender, RoutedEventArgs e)
        {
            if (_isClosing) return;
            _isClosing = true;
            if (_shadeStart.HasValue && _shadeEnd.HasValue)
            {
                _viewModel.ToggleCustomShade(_shadeStart.Value, _shadeEnd.Value);
            }
            this.Close();
        }

        private void AddTaskAbove_Click(object sender, RoutedEventArgs e)
        {
            if (_isClosing) return;
            _isClosing = true;
            _viewModel.InsertTaskAt(_insertionIndex, _insertionDate, false, _targetTask, false);
            this.Close();
        }

        private void AddTaskBelow_Click(object sender, RoutedEventArgs e)
        {
            if (_isClosing) return;
            _isClosing = true;
            _viewModel.InsertTaskAt(_insertionIndex + 1, _insertionDate, false, _targetTask, true);
            this.Close();
        }

        private void AddSubMilestone_Click(object sender, RoutedEventArgs e)
        {
            if (_isClosing) return;
            _isClosing = true;
            _viewModel.InsertSubMilestoneAt(_insertionIndex, _insertionDate);
            this.Close();
        }

        private void AddSubBar_Click(object sender, RoutedEventArgs e)
        {
            if (_isClosing) return;
            _isClosing = true;
            _viewModel.InsertSubBarAt(_insertionIndex, _insertionDate);
            this.Close();
        }

        private void AddMilestone_Click(object sender, RoutedEventArgs e)
        {
            if (_isClosing) return;
            _isClosing = true;
            _viewModel.InsertTaskAt(_insertionIndex, _insertionDate, true);
            this.Close();
        }

        private void HighlightRow_Click(object sender, RoutedEventArgs e)
        {
            if (_isClosing || _targetTask == null) return;
            _isClosing = true;
            _targetTask.IsHighlighted = !_targetTask.IsHighlighted;
            _viewModel.RefreshChartCommand.Execute(null);
            this.Close();
        }

        private void AddEventLine_Click(object sender, RoutedEventArgs e)
        {
            if (_isClosing) return;
            _isClosing = true;
            _viewModel.InsertEventLineAt(_insertionIndex, _insertionDate);
            this.Close();
        }

        private void Indent_Click(object sender, RoutedEventArgs e)
        {
            if (_isClosing || _targetTask == null) return;
            _isClosing = true;
            _viewModel.SelectedTask = _targetTask;
            _viewModel.IndentCommand.Execute(null);
            _viewModel.RefreshChartCommand.Execute(null);
            this.Close();
        }

        private void Outdent_Click(object sender, RoutedEventArgs e)
        {
            if (_isClosing || _targetTask == null) return;
            _isClosing = true;
            _viewModel.SelectedTask = _targetTask;
            _viewModel.OutdentCommand.Execute(null);
            _viewModel.RefreshChartCommand.Execute(null);
            this.Close();
        }

        private void Window_Deactivated(object sender, EventArgs e)
        {
            // FOCUS LOCK: Ignore deactivation events for the first 500ms.
            if ((DateTime.Now - _openedTime).TotalMilliseconds < 500)
            {
                oPenEfficiency.ThisAddIn.LogDebugStatic("GanttContextMenu Deactivated IGNORED (Focus Lock active).");
                return;
            }

            oPenEfficiency.ThisAddIn.LogDebugStatic("GanttContextMenu Deactivated. Closing window.");
            if (_isClosing) return;
            _isClosing = true;
            try { this.Close(); }
            catch (InvalidOperationException)
            {
                // The modeless menu is already closing or closed.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttContextMenu.CloseMenu");
            }
        }
    }
}
