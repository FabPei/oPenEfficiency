using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.Win32;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.UI;
using oPenEfficiency.Utils;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.UI.Panels
{
    public class GanttPanelViewModel : INotifyPropertyChanged
    {
        private GanttProject _project;
        private GanttTask _selectedTask;
        private PowerPointManager _pptManager;
        private PersistenceManager _persistenceManager;
        private GanttUndoRedoManager _undoRedoManager;
        private DispatcherTimer _debounceTimer;
        private bool _isLivePreviewEnabled = true;
        private bool _isApplyingProjectMutation;
        private bool _hasRenderedChart;
        private string _statusMessage;

        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(); }
        }

        public GanttProject Project
        {
            get => _project;
            set
            {
                _project = value;
                OnPropertyChanged();
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public GanttTask SelectedTask
        {
            get => _selectedTask;
            set
            {
                _selectedTask = value;
                OnPropertyChanged();
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public bool IsLivePreviewEnabled
        {
            get => _isLivePreviewEnabled;
            set { _isLivePreviewEnabled = value; OnPropertyChanged(); }
        }

        public bool IsDataEditable => Project != null && !Project.IsLinkedToExcel;

        public bool HasRenderedChart
        {
            get => _hasRenderedChart;
            private set
            {
                _hasRenderedChart = value;
                OnPropertyChanged();
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public ICommand AddTaskCommand { get; }
        public RelayCommand AddMilestoneCommand { get; }
        public RelayCommand AdoptAsTaskCommand { get; }
        public RelayCommand AdoptAsMilestoneCommand { get; }
        public RelayCommand DeleteTaskCommand { get; }
        public ICommand IndentCommand { get; }
        public ICommand OutdentCommand { get; }
        public ICommand RefreshChartCommand { get; }
        public ICommand CreateNewChartCommand { get; }
        public ICommand UndoCommand { get; }
        public ICommand RedoCommand { get; }
        public ICommand ImportExcelCommand { get; }
        public ICommand SyncFromSlideCommand { get; }
        public ICommand CreateExcelTemplateCommand { get; }
        public ICommand ExportToExcelCommand { get; }
        public ICommand SyncExcelCommand { get; }
        public ICommand UnlinkExcelCommand { get; }
        public ICommand ToggleCollapseCommand { get; }
        public ObservableCollection<GanttTask> VisibleTasks { get; } = new ObservableCollection<GanttTask>();

        public GanttPanelViewModel(PowerPointManager pptManager)
        {
            _pptManager = pptManager;
            _persistenceManager = new PersistenceManager(pptManager);
            _undoRedoManager = new GanttUndoRedoManager();

            _debounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
            _debounceTimer.Tick += DebounceTimer_Tick;

            Shape selectedShape = null;
            try
            {
                var selection = _pptManager.GetSelectedShapes();
                if (selection != null && selection.Count > 0)
                {
                    selectedShape = selection[1];
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttPanelViewModel.Constructor.GetSelectedShape");
            }

            Project = _persistenceManager.LoadProject(selectedShape);
            HasRenderedChart = IsCurrentChartRendered();

            if (!HasRenderedChart)
            {
                Project = CreateStarterProject();
            }
            else if (Project.Tasks.Count == 0)
            {
                Project.Tasks.Add(new GanttTask { Id = Guid.NewGuid(), Title = "Project Kickoff", Start = DateTime.Today, End = DateTime.Today, OutlineLevel = 0 });
            }

            AddTaskCommand = new RelayCommand(AddTask, p => !Project.IsLinkedToExcel);
            AddMilestoneCommand = new RelayCommand(AddMilestone, p => !Project.IsLinkedToExcel);
            DeleteTaskCommand = new RelayCommand(DeleteTask, p => SelectedTask != null && !Project.IsLinkedToExcel);
            IndentCommand = new RelayCommand(IndentTask, p => !Project.IsLinkedToExcel && GanttProjectOperations.CanIndentTask(Project, SelectedTask));
            OutdentCommand = new RelayCommand(OutdentTask, p => SelectedTask != null && !Project.IsLinkedToExcel);
            RefreshChartCommand = new RelayCommand(p => UpdatePowerPoint(), p => HasRenderedChart);
            CreateNewChartCommand = new RelayCommand(CreateNewChart);
            UndoCommand = new RelayCommand(p => PerformUndo(), p => _undoRedoManager.CanUndo && !Project.IsLinkedToExcel);
            RedoCommand = new RelayCommand(p => PerformRedo(), p => _undoRedoManager.CanRedo && !Project.IsLinkedToExcel);
            ImportExcelCommand = new RelayCommand(ImportExcel, p => !Project.IsLinkedToExcel);
            SyncFromSlideCommand = new RelayCommand(SyncFromSlide, p => !Project.IsLinkedToExcel);
            CreateExcelTemplateCommand = new RelayCommand(CreateExcelTemplate);
            ExportToExcelCommand = new RelayCommand(ExportToExcel);
            SyncExcelCommand = new RelayCommand(SyncExcel, p => Project.IsLinkedToExcel);
            UnlinkExcelCommand = new RelayCommand(UnlinkExcel, p => Project.IsLinkedToExcel);

            AdoptAsTaskCommand = new RelayCommand(p => AdoptSelection(false), p => !Project.IsLinkedToExcel);
            AdoptAsMilestoneCommand = new RelayCommand(p => AdoptSelection(true), p => !Project.IsLinkedToExcel);
            ToggleCollapseCommand = new RelayCommand(p =>
            {
                if (p is GanttTask t && t.IsSummary)
                    t.IsCollapsed = !t.IsCollapsed;
            });

            AttachProjectEvents();
            _undoRedoManager.SaveState(Project);
        }

        private void CreateNewChart(object parameter)
        {
            _debounceTimer.Stop();
            if (HasRenderedChart) UpdatePowerPoint();
            DetachProjectEvents();

            Project = CreateStarterProject();
            SelectedTask = Project.Tasks[0];
            HasRenderedChart = false;
            _undoRedoManager = new GanttUndoRedoManager();
            AttachProjectEvents();
            _undoRedoManager.SaveState(Project);
            OnPropertyChanged(nameof(IsDataEditable));

            GetSuggestedNewChartPosition(out double topOffset, out double leftMargin);
            UpdatePowerPoint(topOffset, leftMargin);
            StatusMessage = HasRenderedChart
                ? "New chart created. Edit the starter task or add more lanes."
                : "The new chart could not be created. Check the openEfficiency error log.";
        }

        private static GanttProject CreateStarterProject()
        {
            var project = new GanttProject();
            DateTime start = GanttMathEngine.NextWorkDay(DateTime.Today);
            project.Tasks.Add(new GanttTask
            {
                Id = Guid.NewGuid(),
                Title = "New Task",
                Start = start,
                End = GanttMathEngine.AddWorkDays(start, 1),
                OutlineLevel = 0
            });
            project.RecalculateWBS();
            return project;
        }

        private void DetachProjectEvents()
        {
            if (Project == null) return;

            Project.Tasks.CollectionChanged -= Tasks_CollectionChanged;
            Project.PropertyChanged -= Project_PropertyChanged;
            foreach (var task in Project.Tasks)
            {
                task.PropertyChanged -= Task_PropertyChanged;
            }
        }

        private bool IsCurrentChartRendered()
        {
            Slide slide = _pptManager.GetCurrentSlide();
            if (slide == null) return false;

            Shape group = null;
            try
            {
                group = PowerPointShapeManager.FindChartGroup(slide, Project.ChartId);
                return group != null;
            }
            finally
            {
                if (group != null) Marshal.ReleaseComObject(group);
            }
        }

        private void GetSuggestedNewChartPosition(out double topOffset, out double leftMargin)
        {
            topOffset = 100;
            leftMargin = 50;

            Slide slide = _pptManager.GetCurrentSlide();
            if (slide == null) return;

            Shapes shapes = null;
            int chartCount = 0;
            try
            {
                shapes = slide.Shapes;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    Shape shape = null;
                    try
                    {
                        shape = shapes[i];
                        if (PowerPointShapeManager.HasGanttProjectTag(shape)) chartCount++;
                    }
                    finally
                    {
                        if (shape != null) Marshal.ReleaseComObject(shape);
                    }
                }
            }
            finally
            {
                if (shapes != null) Marshal.ReleaseComObject(shapes);
            }

            // Cascade additional charts so the new group remains visible and selectable.
            int offsetSteps = Math.Min(chartCount, 6);
            topOffset += offsetSteps * 20;
            leftMargin += offsetSteps * 20;
        }

        private void AttachProjectEvents()
        {
            Project.Tasks.CollectionChanged -= Tasks_CollectionChanged;
            Project.Tasks.CollectionChanged += Tasks_CollectionChanged;

            Project.PropertyChanged -= Project_PropertyChanged;
            Project.PropertyChanged += Project_PropertyChanged;

            foreach (var task in Project.Tasks)
            {
                task.PropertyChanged -= Task_PropertyChanged;
                task.PropertyChanged += Task_PropertyChanged;
            }

            UpdateVisibleTasks();
        }

        private void Tasks_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.NewItems != null)
            {
                foreach (GanttTask item in e.NewItems)
                {
                    item.PropertyChanged -= Task_PropertyChanged;
                    item.PropertyChanged += Task_PropertyChanged;
                }
            }
            if (_isApplyingProjectMutation) return;

            UpdateVisibleTasks();
            RequestUpdate(true);
        }

        private void Project_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(GanttProject.ViewMode) ||
                e.PropertyName == nameof(GanttProject.Theme) ||
                e.PropertyName == nameof(GanttProject.FontFamily) ||
                e.PropertyName == nameof(GanttProject.ShowNrColumn) ||
                e.PropertyName == nameof(GanttProject.ShowTaskNameColumn) ||
                e.PropertyName == nameof(GanttProject.ShowDurationColumn) ||
                e.PropertyName == nameof(GanttProject.ShowLegend) ||
                e.PropertyName == nameof(GanttProject.ShowYears) ||
                e.PropertyName == nameof(GanttProject.ShowMonths) ||
                e.PropertyName == nameof(GanttProject.ShowWeeks) ||
                e.PropertyName == nameof(GanttProject.ShowDays) ||
                e.PropertyName == nameof(GanttProject.AutoFitToSlide) ||
                e.PropertyName == nameof(GanttProject.LockElements) ||
                e.PropertyName == nameof(GanttProject.ShadeWeekends) ||
                e.PropertyName == nameof(GanttProject.FollowGanttRules) ||
                e.PropertyName == nameof(GanttProject.UseManualRange) ||
                e.PropertyName == nameof(GanttProject.ManualStart) ||
                e.PropertyName == nameof(GanttProject.ManualEnd))
            {
                RequestUpdate(true);
            }

            if (e.PropertyName == nameof(GanttProject.IsLinkedToExcel))
            {
                if (Project.IsLinkedToExcel)
                {
                    StatusMessage = "Note: Data is linked to Excel and read-only in PowerPoint.";
                }
                else
                {
                    StatusMessage = string.Empty;
                }

                // Refresh command states
                CommandManager.InvalidateRequerySuggested();
                OnPropertyChanged(nameof(IsDataEditable)); // For XAML binding if we add it
            }
        }

        private bool _isCalculating = false;
        private void Task_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (_isCalculating) return;
            if (!(sender is GanttTask task)) return;

            if (e.PropertyName == nameof(GanttTask.OutlineLevel))
            {
                _isCalculating = true;
                try { Project.RecalculateWBS(); }
                finally { _isCalculating = false; }
                UpdateVisibleTasks();
                RequestUpdate(true);
            }
            else if (e.PropertyName == nameof(GanttTask.Start) || e.PropertyName == nameof(GanttTask.End) || e.PropertyName == nameof(GanttTask.PercentComplete) || e.PropertyName == nameof(GanttTask.Title))
            {
                _isCalculating = true;
                try
                {
                    // Date validation: Start should not be after End
                    if (task.Start > task.End)
                    {
                        if (e.PropertyName == nameof(GanttTask.Start))
                        {
                            task.End = task.Start;
                        }
                        else
                        {
                            task.Start = task.End;
                        }
                    }

                    if (CheckForCircularDependencies(task))
                    {
                        StatusMessage = "Warning: Circular dependency detected!";
                        return;
                    }

                    StatusMessage = string.Empty;
                    AdjustSuccessors(task);
                    Project.RecalculateWBS();
                }
                finally
                {
                    _isCalculating = false;
                }
                RequestUpdate(true);
            }
            else if (e.PropertyName == nameof(GanttTask.Shape) || e.PropertyName == nameof(GanttTask.CustomColor) || e.PropertyName == nameof(GanttTask.MilestoneShape))
            {
                RequestUpdate(true);
            }
            else if (e.PropertyName == nameof(GanttTask.IsCollapsed))
            {
                UpdateVisibleTasks();
                RequestUpdate(true);
            }
            else if (e.PropertyName == nameof(GanttTask.PredecessorsString))
            {
                _isCalculating = true;
                try
                {
                    var newPredIds = new List<Guid>();
                    var newPredLinks = new List<Models.GanttDependencyLink>();

                    if (!string.IsNullOrWhiteSpace(task.PredecessorsString))
                    {
                        var tokens = task.PredecessorsString.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var token in tokens)
                        {
                            var (wbs, depType, lag) = ParseDependencyToken(token);
                            var predTask = Project.Tasks.FirstOrDefault(t => t.WBS == wbs);
                            if (predTask != null && predTask.Id != task.Id)
                            {
                                newPredIds.Add(predTask.Id);
                                newPredLinks.Add(new Models.GanttDependencyLink
                                {
                                    PredecessorId = predTask.Id,
                                    Type = depType,
                                    LagDays = lag
                                });
                            }
                        }
                    }

                    var oldPredIds = task.PredecessorIds;
                    var oldPredLinks = task.PredecessorLinks;
                    task.PredecessorIds = newPredIds;
                    task.PredecessorLinks = newPredLinks;

                    if (CheckForCircularDependencies(task))
                    {
                        StatusMessage = "Warning: Circular dependency detected! Reverting.";
                        task.PredecessorIds = oldPredIds;
                        task.PredecessorLinks = oldPredLinks;
                        Project.RecalculateWBS();
                        return;
                    }

                    StatusMessage = string.Empty;
                    AdjustSuccessors(task);
                    Project.RecalculateWBS();
                }
                finally
                {
                    _isCalculating = false;
                }
                RequestUpdate(true);
            }
        }

        /// <summary>
        /// Parses a single predecessor token from the Pred. column.
        /// Accepted formats: "1", "1FS", "2SS", "3FF", "2SS+2", "3FF-1"
        /// </summary>
        private static (string wbs, Models.DependencyType type, int lag) ParseDependencyToken(string token)
        {
            token = token.Trim();
            if (string.IsNullOrEmpty(token)) return (token, Models.DependencyType.FinishToStart, 0);

            // Find the first letter — everything before it is the WBS number.
            int typeStart = -1;
            for (int i = 0; i < token.Length; i++)
            {
                if (char.IsLetter(token[i])) { typeStart = i; break; }
            }

            if (typeStart < 0)
                return (token, Models.DependencyType.FinishToStart, 0);

            string wbs = token.Substring(0, typeStart);
            string rest = token.Substring(typeStart); // e.g. "FS+2", "SS", "FF-1"

            var depType = Models.DependencyType.FinishToStart;
            int lagDays = 0;

            if (rest.Length >= 2)
            {
                switch (rest.Substring(0, 2).ToUpperInvariant())
                {
                    case "SS": depType = Models.DependencyType.StartToStart; break;
                    case "FF": depType = Models.DependencyType.FinishToFinish; break;
                }
                if (rest.Length > 2)
                    int.TryParse(rest.Substring(2), out lagDays);
            }

            return (wbs, depType, lagDays);
        }

        private void ResolvePredecessorLinks()
        {
            foreach (var task in Project.Tasks)
            {
                if (string.IsNullOrWhiteSpace(task.PredecessorsString)) continue;

                var newPredIds = new List<Guid>();
                var newPredLinks = new List<Models.GanttDependencyLink>();

                var tokens = task.PredecessorsString.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var token in tokens)
                {
                    var (wbs, depType, lag) = ParseDependencyToken(token);
                    var predTask = Project.Tasks.FirstOrDefault(t => t.WBS == wbs);
                    if (predTask != null && predTask.Id != task.Id)
                    {
                        newPredIds.Add(predTask.Id);
                        newPredLinks.Add(new Models.GanttDependencyLink { PredecessorId = predTask.Id, Type = depType, LagDays = lag });
                    }
                }

                task.PredecessorIds = newPredIds;
                task.PredecessorLinks = newPredLinks;
            }
        }

        public void UpdateVisibleTasks()
        {
            VisibleTasks.Clear();
            int collapsedLevel = -1;
            foreach (var task in Project.Tasks)
            {
                if (collapsedLevel >= 0 && task.OutlineLevel > collapsedLevel)
                {
                    task.IsVisible = false;
                    continue;
                }
                collapsedLevel = -1;
                task.IsVisible = true;
                VisibleTasks.Add(task);
                if (task.IsSummary && task.IsCollapsed)
                    collapsedLevel = task.OutlineLevel;
            }
        }

        private void AdjustSuccessors(GanttTask predecessor)
        {
            if (!Project.FollowGanttRules) return;

            if (!GanttProjectOperations.RecalculateDependencySchedule(Project))
                StatusMessage = "Warning: Circular dependency detected!";
        }

        private bool CheckForCircularDependencies(GanttTask task)
        {
            return HasCycle(task.Id, new HashSet<Guid>());
        }

        private bool HasCycle(Guid taskId, HashSet<Guid> visited)
        {
            if (visited.Contains(taskId)) return true;
            visited.Add(taskId);

            var task = Project.Tasks.FirstOrDefault(t => t.Id == taskId);
            if (task == null) return false;

            foreach (var predId in task.PredecessorIds)
            {
                if (HasCycle(predId, new HashSet<Guid>(visited))) return true;
            }

            return false;
        }

        public void ToggleCustomShade(DateTime start, DateTime end)
        {
            if (Project.IsLinkedToExcel) return;

            _undoRedoManager.SaveState(Project);

            // Check if already exists
            var existing = Project.CustomShades.FirstOrDefault(s => s.Start == start && s.End == end);
            if (existing != null)
            {
                Project.CustomShades.Remove(existing);
            }
            else
            {
                Project.CustomShades.Add(new GanttCustomShade
                {
                    Start = start,
                    End = end,
                    Color = "#4A90E2", // Professional blue
                    Transparency = 0.2f
                });
            }

            RequestUpdate(true);
        }

        public void InsertEventLineAt(int index, DateTime date)
        {
            if (Project.IsLinkedToExcel) return;

            _undoRedoManager.SaveState(Project);

            var newTask = new GanttTask
            {
                Id = Guid.NewGuid(),
                Title = "New Event",
                Start = date,
                End = date,
                IsEventLine = true,
                OutlineLevel = SelectedTask?.OutlineLevel ?? 0
            };

            if (index < 0) index = 0;
            if (index > Project.Tasks.Count) index = Project.Tasks.Count;

            Project.Tasks.Insert(index, newTask);
            SelectedTask = newTask;
            Project.RecalculateWBS();
            RequestUpdate(true);
        }

        public void InsertSubMilestoneAt(int index, DateTime date)
        {
            if (Project.IsLinkedToExcel) return;
            if (Project.Tasks == null || index < 0 || index >= Project.Tasks.Count) return;

            _undoRedoManager.SaveState(Project);

            var task = Project.Tasks[index];

            var subEvent = new GanttSubEvent
            {
                Date = date,
                Title = "Sub-Milestone",
                Shape = MilestoneShape.Diamond
            };

            task.SubEvents.Add(subEvent);
            RequestUpdate(true);
        }

        public void InsertSubBarAt(int index, DateTime date)
        {
            if (Project.IsLinkedToExcel) return;
            if (Project.Tasks == null || index < 0 || index >= Project.Tasks.Count) return;

            _undoRedoManager.SaveState(Project);

            var task = Project.Tasks[index];

            var subBar = new GanttSubBar
            {
                Start = date,
                End = date.AddDays(5), // Default duration
                Title = "Sub-Task",
                Shape = TaskShape.Rectangle
            };

            task.SubBars.Add(subBar);
            RequestUpdate(true);
        }

        public void InsertTaskAt(
            int index,
            DateTime date,
            bool isMilestone,
            GanttTask referenceTask = null,
            bool insertBelowReference = false)
        {
            if (Project.IsLinkedToExcel) return;

            int referenceIndex = referenceTask != null
                ? Project.Tasks.IndexOf(referenceTask)
                : -1;

            if (referenceIndex >= 0)
            {
                index = referenceIndex;

                if (insertBelowReference)
                {
                    // A same-level sibling belongs after the complete subtree. Inserting it
                    // directly before deeper rows would make those rows its children and turn
                    // the new task into an unintended summary task.
                    index++;
                    while (index < Project.Tasks.Count &&
                           Project.Tasks[index].OutlineLevel > referenceTask.OutlineLevel)
                    {
                        index++;
                    }
                }
            }

            DateTime taskStart = Project.FollowGanttRules
                ? GanttMathEngine.NextWorkDay(date.Date)
                : date.Date;
            var newTask = new GanttTask
            {
                Id = Guid.NewGuid(),
                Title = isMilestone ? "New Milestone" : "New Task",
                Start = taskStart,
                End = isMilestone
                    ? taskStart
                    : Project.FollowGanttRules
                        ? GanttMathEngine.AddWorkDays(taskStart, 1)
                        : taskStart.AddDays(1),
                OutlineLevel = referenceTask?.OutlineLevel ?? SelectedTask?.OutlineLevel ?? 0,
                Shape = referenceTask?.Shape ?? TaskShape.Rectangle,
                MilestoneShape = referenceTask?.MilestoneShape ?? MilestoneShape.Diamond,
                CustomColor = referenceTask?.CustomColor
            };

            if (index < 0) index = 0;
            if (index > Project.Tasks.Count) index = Project.Tasks.Count;

            ExecuteProjectMutation(() =>
            {
                Project.Tasks.Insert(index, newTask);
                SelectedTask = newTask;
            });
        }

        private void AddTask(object parameter)
        {
            DateTime start = Project.FollowGanttRules
                ? GanttMathEngine.NextWorkDay(DateTime.Today)
                : DateTime.Today;
            var newTask = new GanttTask
            {
                Id = Guid.NewGuid(),
                Title = "New Task",
                Start = start,
                End = Project.FollowGanttRules
                    ? GanttMathEngine.AddWorkDays(start, 1)
                    : start.AddDays(1),
                OutlineLevel = SelectedTask?.OutlineLevel ?? 0
            };

            int index = SelectedTask != null
                ? GanttProjectOperations.GetIndexAfterSubtree(Project, Project.Tasks.IndexOf(SelectedTask))
                : Project.Tasks.Count;
            ExecuteProjectMutation(() =>
            {
                Project.Tasks.Insert(index, newTask);
                SelectedTask = newTask;
            });
        }

        private void AddMilestone(object parameter)
        {
            DateTime date = Project.FollowGanttRules
                ? GanttMathEngine.NextWorkDay(DateTime.Today)
                : DateTime.Today;
            var newTask = new GanttTask
            {
                Id = Guid.NewGuid(),
                Title = "New Milestone",
                Start = date,
                End = date, // Start == End makes it a milestone
                OutlineLevel = SelectedTask?.OutlineLevel ?? 0
            };

            int index = SelectedTask != null
                ? GanttProjectOperations.GetIndexAfterSubtree(Project, Project.Tasks.IndexOf(SelectedTask))
                : Project.Tasks.Count;
            ExecuteProjectMutation(() =>
            {
                Project.Tasks.Insert(index, newTask);
                SelectedTask = newTask;
            });
        }

        private void DeleteTask(object parameter)
        {
            if (SelectedTask == null) return;
            GanttTask taskToDelete = SelectedTask;
            ExecuteProjectMutation(() =>
            {
                GanttProjectOperations.DeleteTask(Project, taskToDelete);
                SelectedTask = null;
            });
        }

        private void IndentTask(object parameter)
        {
            if (SelectedTask == null) return;
            ExecuteProjectMutation(() => GanttProjectOperations.IndentTask(Project, SelectedTask));
        }

        public void OutdentTask(object parameter)
        {
            if (SelectedTask == null || SelectedTask.OutlineLevel <= 0) return;
            ExecuteProjectMutation(() => SelectedTask.OutlineLevel--);
        }

        public void ReorderTask(int oldIndex, int newIndex)
        {
            if (oldIndex < 0 || oldIndex >= Project.Tasks.Count || newIndex < 0 || newIndex >= Project.Tasks.Count) return;

            var task = Project.Tasks[oldIndex];
            ExecuteProjectMutation(() =>
            {
                if (GanttProjectOperations.MoveTaskSubtree(Project, oldIndex, newIndex))
                {
                    SelectedTask = task;
                }
            });
        }

        private void ExecuteProjectMutation(Action mutation)
        {
            if (mutation == null) return;

            _isApplyingProjectMutation = true;
            _isCalculating = true;
            try
            {
                _undoRedoManager.ExecuteTransaction(Project, () =>
                {
                    mutation();
                    Project.RecalculateWBS();
                });
            }
            finally
            {
                _isCalculating = false;
                _isApplyingProjectMutation = false;
            }

            UpdateVisibleTasks();
            RequestUpdate(false);
        }

        private void PerformUndo()
        {
            var oldState = _undoRedoManager.Undo(Project);
            if (oldState != null)
            {
                DetachProjectEvents();
                Project = oldState;
                AttachProjectEvents();
                RequestUpdate(false);
            }
        }

        private void PerformRedo()
        {
            var oldState = _undoRedoManager.Redo(Project);
            if (oldState != null)
            {
                DetachProjectEvents();
                Project = oldState;
                AttachProjectEvents();
                RequestUpdate(false);
            }
        }

        private void ImportExcel(object parameter)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog.ShowDialog() == true)
            {
                GanttExcelImportResult result = GanttExcelManager.ImportFromExcel(openFileDialog.FileName, Project.Tasks);
                if (result.Tasks.Any())
                {
                    ApplyExcelTasks(result.Tasks);
                    StatusMessage = BuildExcelStatus($"Imported {result.Tasks.Count} tasks from Excel.", result);
                }
                else
                {
                    StatusMessage = BuildExcelStatus("No valid Excel tasks were imported.", result);
                }
            }
        }

        private void CreateExcelTemplate(object parameter)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xlsx";
            saveFileDialog.FileName = "Gantt_Template.xlsx";
            if (saveFileDialog.ShowDialog() == true)
            {
                GanttExcelManager.CreateTemplate(saveFileDialog.FileName);
                StatusMessage = "Excel template created.";
            }
        }

        private void ExportToExcel(object parameter)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xlsx";
            saveFileDialog.FileName = "Gantt_Export.xlsx";
            if (saveFileDialog.ShowDialog() == true)
            {
                GanttExcelManager.ExportToExcel(Project, saveFileDialog.FileName);
                StatusMessage = "Project exported to Excel.";
            }
        }

        private void SyncExcel(object parameter)
        {
            if (string.IsNullOrEmpty(Project.LinkedExcelFilePath)) return;

            GanttExcelImportResult result = GanttExcelManager.ImportFromExcel(Project.LinkedExcelFilePath, Project.Tasks);
            if (result.Tasks.Any())
            {
                ApplyExcelTasks(result.Tasks);
                StatusMessage = BuildExcelStatus("Synced from Excel.", result);
            }
            else
            {
                StatusMessage = BuildExcelStatus("Excel sync contained no valid tasks; the chart was not changed.", result);
            }
        }

        public void LinkExcelFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) return;

            ExecuteProjectMutation(() => Project.LinkedExcelFilePath = filePath);
            SyncExcel(null);
            _debounceTimer.Stop();
            if (HasRenderedChart) UpdatePowerPoint();
        }

        private void UnlinkExcel(object parameter)
        {
            var confirmation = System.Windows.MessageBox.Show(
                "Remove the Excel source connection from this Gantt chart?\n\nThe current tasks and chart content will be kept.",
                "Unlink Gantt Excel Source",
                System.Windows.MessageBoxButton.OKCancel,
                System.Windows.MessageBoxImage.Warning);
            if (confirmation != System.Windows.MessageBoxResult.OK) return;

            ExecuteProjectMutation(() => Project.LinkedExcelFilePath = null);
            _debounceTimer.Stop();
            if (HasRenderedChart) UpdatePowerPoint();
            StatusMessage = "Excel link removed.";
        }

        private void ApplyExcelTasks(IList<GanttTask> importedTasks)
        {
            ExecuteProjectMutation(() =>
            {
                Project.Tasks.Clear();
                foreach (GanttTask task in importedTasks) Project.Tasks.Add(task);
                Project.RecalculateWBS();
                ResolvePredecessorLinks();
            });
        }

        private static string BuildExcelStatus(string message, GanttExcelImportResult result)
        {
            if (result == null || result.Warnings.Count == 0) return message;
            string firstWarning = result.Warnings[0];
            string suffix = result.Warnings.Count == 1
                ? firstWarning
                : firstWarning + $" (+{result.Warnings.Count - 1} more)";
            return message + " Warning: " + suffix;
        }

        private void SyncFromSlide(object parameter)
        {
            var slide = _pptManager.GetCurrentSlide();
            if (slide == null) return;

            float slideWidth = GetSlideWidthOrDefault(
                slide,
                720,
                "GanttPanelViewModel.SyncFromSlide.GetSlideWidth");

            double totalWidth = slideWidth - 100;
            double leftMargin = 50;

            double effectiveLeft = leftMargin;
            Shape group = PowerPointShapeManager.FindChartGroup(slide, Project.ChartId);
            if (group != null) effectiveLeft = group.Left;

            // The chart itself knows which scale and which timeline day-zero it was rendered with;
            // a resized bar must be read in that frame (REG-GANTT-RESIZE-ADOPT-001). Legacy charts
            // carry no tags and fall back to the recomputation and the live ProjectStart.
            double recordedPointsPerDay = PowerPointShapeManager.GetGanttScale(group);
            DateTime recordedProjectStart = PowerPointShapeManager.GetGanttProjectStart(group);

            double tableWidth = 0;
            if (Project.ShowNrColumn) tableWidth += 30; // COL_NO_WIDTH
            if (Project.ShowTaskNameColumn) tableWidth += 170; // COL_NAME_WIDTH
            if (Project.ShowDurationColumn) tableWidth += 50; // COL_DURATION_WIDTH

            double exactDays = (Project.ProjectEnd - Project.ProjectStart).TotalDays;
            double totalDays = Math.Max(1, Math.Ceiling(exactDays));

            double pointsPerDay = 20;
            if (Project.AutoFitToSlide)
            {
                double availableGanttWidth = totalWidth - tableWidth;
                if (availableGanttWidth < 50) availableGanttWidth = 50;
                pointsPerDay = availableGanttWidth / totalDays;
            }
            else
            {
                if (Project.ShowDays) pointsPerDay = 20;
                else if (Project.ShowWeeks) pointsPerDay = 5;
                else if (Project.ShowMonths) pointsPerDay = 1.5;
                else if (Project.ShowYears) pointsPerDay = 0.4;
            }

            pointsPerDay = GanttSyncAdoptionPolicy.ResolvePointsPerDay(recordedPointsPerDay, pointsPerDay);

            DateTime adoptionOrigin = recordedProjectStart == default ? Project.ProjectStart : recordedProjectStart;

            double ganttLeftMargin = effectiveLeft + tableWidth;

            bool anyChanges = false;
            _isCalculating = true;
            _isApplyingProjectMutation = true;
            var shapesToScan = new List<Shape>();
            _undoRedoManager.SaveState(Project);

            try
            {
                if (group != null && group.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                {
                    foreach (Shape s in group.GroupItems) shapesToScan.Add(s);
                }
                else
                {
                    foreach (Shape s in slide.Shapes) shapesToScan.Add(s);
                }

                // --- PASS 1: Task names and vertical reordering via RowAnchors ---
                var anchorTops = new Dictionary<Guid, double>();
                foreach (Shape shape in shapesToScan)
                {
                    try
                    {
                        string idString = PowerPointShapeManager.GetGanttId(shape);
                        string typeString = PowerPointShapeManager.GetGanttType(shape);

                        if (!string.IsNullOrEmpty(idString) && typeString == PowerPointShapeManager.GanttShapeType.RowAnchor.ToString())
                        {
                            if (Guid.TryParse(idString, out Guid taskId))
                            {
                                anchorTops[taskId] = shape.Top;

                                var task = Project.Tasks.FirstOrDefault(t => t.Id == taskId);
                                if (task != null && shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                                {
                                    string slideTitle = shape.TextFrame.TextRange.Text;
                                    slideTitle = slideTitle?.TrimEnd('\r', '\n');

                                    if (slideTitle != null && !string.Equals(task.Title, slideTitle, StringComparison.Ordinal))
                                    {
                                        task.Title = slideTitle;
                                        anyChanges = true;
                                    }
                                }
                            }
                        }
                    }
                    catch (COMException)
                    {
                        // A deleted or regrouped child is skipped during slide synchronization.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "GanttPanelViewModel.SyncFromSlide.ReadAnchor");
                    }
                }

                if (anchorTops.Count > 0)
                {
                    // In rule-based mode, sort siblings but keep every hierarchy subtree intact.
                    // Flexible timelines intentionally allow independent lane ordering.
                    var sortedTasks = Project.FollowGanttRules
                        ? FlattenTasksByAnchor(Project.Tasks.Where(task => task.ParentId == null), anchorTops).ToList()
                        : Project.Tasks.OrderBy(task => anchorTops.ContainsKey(task.Id) ? anchorTops[task.Id] : double.MaxValue).ToList();

                    // Check if order actually changed
                    bool orderChanged = false;
                    for (int i = 0; i < Project.Tasks.Count; i++)
                    {
                        if (Project.Tasks[i].Id != sortedTasks[i].Id)
                        {
                            orderChanged = true;
                            break;
                        }
                    }

                    if (orderChanged)
                    {
                        Project.Tasks.Clear();
                        foreach (var t in sortedTasks)
                        {
                            Project.Tasks.Add(t);
                        }
                        anyChanges = true;
                    }
                }

                // --- PASS 2: Horizontal Date Adjustments via Bars ---
                foreach (Shape shape in shapesToScan)
                {
                    try
                    {
                        string idString = PowerPointShapeManager.GetGanttId(shape);
                        string typeString = PowerPointShapeManager.GetGanttType(shape);

                        if (!string.IsNullOrEmpty(idString) &&
                            (typeString == PowerPointShapeManager.GanttShapeType.TaskBar.ToString() ||
                             typeString == PowerPointShapeManager.GanttShapeType.Milestone.ToString() ||
                             typeString == PowerPointShapeManager.GanttShapeType.SummaryBar.ToString() ||
                             typeString == PowerPointShapeManager.GanttShapeType.EventLine.ToString()))
                        {
                            if (Guid.TryParse(idString, out Guid taskId))
                            {
                                var task = Project.Tasks.FirstOrDefault(t => t.Id == taskId);
                                if (task != null)
                                {
                                    double shapeLeft = shape.Left;
                                    double shapeWidth = shape.Width;

                                    GanttSyncAdoptionPolicy.AdoptDates(
                                        shapeLeft,
                                        shapeWidth,
                                        ganttLeftMargin,
                                        pointsPerDay,
                                        adoptionOrigin,
                                        Project.FollowGanttRules,
                                        typeString == PowerPointShapeManager.GanttShapeType.Milestone.ToString(),
                                        out DateTime newStart,
                                        out DateTime newEnd);

                                    if (task.Start != newStart || task.End != newEnd)
                                    {
                                        task.Start = newStart;
                                        task.End = newEnd;
                                        anyChanges = true;
                                        AdjustSuccessors(task);
                                    }
                                }
                            }
                        }
                    }
                    catch (COMException)
                    {
                        // A task bar can disappear while the user edits the slide.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "GanttPanelViewModel.SyncFromSlide.ReadTaskBar");
                    }
                }

                if (anyChanges)
                {
                    Project.RecalculateWBS();
                    StatusMessage = "Data synced from slide.";
                }
                else
                {
                    StatusMessage = "No manual changes detected.";
                }
            }
            finally
            {
                foreach (Shape shape in shapesToScan)
                {
                    if (shape != null) Marshal.ReleaseComObject(shape);
                }
                if (group != null) Marshal.ReleaseComObject(group);
                _isCalculating = false;
                _isApplyingProjectMutation = false;
                if (anyChanges) RequestUpdate(true);
            }
        }

        private static IEnumerable<GanttTask> FlattenTasksByAnchor(
            IEnumerable<GanttTask> siblings,
            IDictionary<Guid, double> anchorTops)
        {
            foreach (GanttTask task in siblings.OrderBy(item =>
                anchorTops.ContainsKey(item.Id) ? anchorTops[item.Id] : double.MaxValue))
            {
                yield return task;
                foreach (GanttTask child in FlattenTasksByAnchor(task.Children, anchorTops))
                {
                    yield return child;
                }
            }
        }

        /// <summary>
        /// Applies a task-bar resize performed by the non-persistent slide overlay.
        /// Coordinates are converted back to whole timeline days before the chart is rendered again.
        /// </summary>
        public void ApplyTaskBarResize(Guid taskId, float shapeLeft, float shapeWidth)
        {
            if (Project.IsLinkedToExcel) return;

            var task = Project.Tasks.FirstOrDefault(item => item.Id == taskId);
            if (task == null || task.IsMilestone || task.IsSummary || task.IsEventLine) return;

            var slide = _pptManager.GetCurrentSlide();
            if (slide == null) return;

            try
            {
                float slideWidth = GetSlideWidthOrDefault(
                    slide,
                    720,
                    "GanttPanelViewModel.ApplyTaskBarResize.GetSlideWidth");

                double effectiveLeft = 50;
                Shape group = PowerPointShapeManager.FindChartGroup(slide, Project.ChartId);
                double recordedPointsPerDay = PowerPointShapeManager.GetGanttScale(group);
                DateTime recordedProjectStart = PowerPointShapeManager.GetGanttProjectStart(group);
                try
                {
                    if (group != null) effectiveLeft = group.Left;
                }
                finally
                {
                    if (group != null) Marshal.ReleaseComObject(group);
                }

                double tableWidth = 0;
                if (Project.ShowNrColumn) tableWidth += 30;
                if (Project.ShowTaskNameColumn) tableWidth += 170;
                if (Project.ShowDurationColumn) tableWidth += 50;

                double totalDays = Math.Max(1, Math.Ceiling((Project.ProjectEnd - Project.ProjectStart).TotalDays));
                double pointsPerDay;
                if (Project.AutoFitToSlide)
                {
                    double availableGanttWidth = (slideWidth - 100) - tableWidth;
                    pointsPerDay = Math.Max(50, availableGanttWidth) / totalDays;
                }
                else if (Project.ShowDays) pointsPerDay = 20;
                else if (Project.ShowWeeks) pointsPerDay = 5;
                else if (Project.ShowMonths) pointsPerDay = 1.5;
                else pointsPerDay = 0.4;

                pointsPerDay = GanttSyncAdoptionPolicy.ResolvePointsPerDay(recordedPointsPerDay, pointsPerDay);

                if (pointsPerDay <= 0) return;

                DateTime adoptionOrigin = recordedProjectStart == default ? Project.ProjectStart : recordedProjectStart;

                double ganttLeftMargin = effectiveLeft + tableWidth;
                int startOffsetDays = (int)GanttSyncAdoptionPolicy.RoundDays((shapeLeft - ganttLeftMargin) / pointsPerDay);
                int durationDays = Math.Max(1, (int)GanttSyncAdoptionPolicy.RoundDays(shapeWidth / pointsPerDay));
                DateTime newStart = adoptionOrigin.AddDays(startOffsetDays).Date;
                DateTime newEnd = newStart.AddDays(durationDays).Date;
                if (Project.FollowGanttRules)
                {
                    newStart = GanttMathEngine.NextWorkDay(newStart);
                    newEnd = GanttMathEngine.NextWorkDay(newEnd);
                    if (newEnd <= newStart) newEnd = GanttMathEngine.AddWorkDays(newStart, 1);
                }

                if (task.Start == newStart && task.End == newEnd) return;

                _undoRedoManager.SaveState(Project);
                _isCalculating = true;
                try
                {
                    task.Start = newStart;
                    task.End = newEnd;
                    AdjustSuccessors(task);
                    Project.RecalculateWBS();
                    StatusMessage = $"Resized '{task.Title}' to {task.Start:d} – {task.End:d}.";
                }
                finally
                {
                    _isCalculating = false;
                }

                UpdatePowerPoint();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttPanelViewModel.ApplyTaskBarResize");
            }
        }

        private void RequestUpdate(bool saveState)
        {
            if (saveState)
            {
                _undoRedoManager.SaveState(Project);
            }

            if (IsLivePreviewEnabled && HasRenderedChart)
            {
                _debounceTimer.Stop();
                _debounceTimer.Start();
            }
        }

        private void DebounceTimer_Tick(object sender, EventArgs e)
        {
            _debounceTimer.Stop();
            UpdatePowerPoint();
        }

        public void UpdatePowerPoint(double topOffset = 100, double leftMargin = 50)
        {
            try
            {
                var slide = _pptManager.GetCurrentSlide();
                if (slide != null)
                {
                    if (Project.FollowGanttRules)
                    {
                        Project.CalculateCriticalPath();
                    }
                    else
                    {
                        foreach (GanttTask task in Project.Tasks) task.TotalSlack = 1;
                    }

                    // Get slide width to support Auto-fit
                    float slideWidth = GetSlideWidthOrDefault(
                        slide,
                        720,
                        "GanttPanelViewModel.UpdatePowerPoint.GetSlideWidth");

                    double availableWidth = Math.Max(100, slideWidth - leftMargin - 50);

                    Shape groupShape = GanttRenderer.Render(slide, Project, topOffset, leftMargin, availableWidth);

                    if (groupShape != null)
                    {
                        _persistenceManager.SaveProject(Project, groupShape);
                        Marshal.ReleaseComObject(groupShape);
                        HasRenderedChart = true;
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttPanelViewModel.UpdatePowerPoint");
            }
        }

        private void AdoptSelection(bool asMilestone)
        {
            if (Project.IsLinkedToExcel) return;

            var slide = _pptManager.GetCurrentSlide();
            if (slide == null) return;

            var selection = _pptManager.GetSelectedShapes();
            if (selection == null || selection.Count == 0)
            {
                StatusMessage = "Please select a shape first.";
                return;
            }

            _undoRedoManager.SaveState(Project);
            _isCalculating = true;
            Shape group = null;

            try
            {
                // Find existing Gantt context
                double effectiveLeft = 50; // Default
                group = PowerPointShapeManager.FindChartGroup(slide, Project.ChartId);
                double recordedPointsPerDay = PowerPointShapeManager.GetGanttScale(group);
                DateTime recordedProjectStart = PowerPointShapeManager.GetGanttProjectStart(group);
                if (group != null) effectiveLeft = group.Left;

                double tableWidth = 0;
                if (Project.ShowNrColumn) tableWidth += 30;
                if (Project.ShowTaskNameColumn) tableWidth += 170;
                if (Project.ShowDurationColumn) tableWidth += 50;

                double ganttLeftMargin = effectiveLeft + tableWidth;

                // Get slide width for Auto-fit calculation
                float slideWidth = GetSlideWidthOrDefault(
                    slide,
                    720,
                    "GanttPanelViewModel.CalculateLayout.GetSlideWidth");
                double totalWidth = slideWidth - 100;

                double pointsPerDay = 20;
                double exactDays = (Project.ProjectEnd - Project.ProjectStart).TotalDays;
                double totalDays = Math.Max(1, Math.Ceiling(exactDays));

                if (Project.AutoFitToSlide)
                {
                    double availableGanttWidth = totalWidth - tableWidth;
                    pointsPerDay = availableGanttWidth / totalDays;
                }
                else
                {
                    if (Project.ShowDays) pointsPerDay = 20;
                    else if (Project.ShowWeeks) pointsPerDay = 5;
                    else if (Project.ShowMonths) pointsPerDay = 1.5;
                    else if (Project.ShowYears) pointsPerDay = 0.4;
                }

                pointsPerDay = GanttSyncAdoptionPolicy.ResolvePointsPerDay(recordedPointsPerDay, pointsPerDay);

                DateTime adoptionOrigin = recordedProjectStart == default ? Project.ProjectStart : recordedProjectStart;

                bool addedAny = false;
                foreach (Shape shape in selection)
                {
                    // Skip if already a Gantt shape
                    if (PowerPointShapeManager.IsGanttShape(shape)) continue;

                    // Calculate Dates in the chart's own rendered scale (no rule snapping here:
                    // imported shapes adopt raw dates, matching the historical behavior).
                    GanttSyncAdoptionPolicy.AdoptDates(
                        shape.Left,
                        shape.Width,
                        ganttLeftMargin,
                        pointsPerDay,
                        adoptionOrigin,
                        false,
                        asMilestone,
                        out DateTime start,
                        out DateTime end);

                    // Create Task
                    var newTask = new GanttTask
                    {
                        Id = Guid.NewGuid(),
                        Title = shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue ? shape.TextFrame.TextRange.Text : (asMilestone ? "New Milestone" : "New Task"),
                        Start = start,
                        End = end,
                        OutlineLevel = 0
                    };

                    // Determine insertion index based on vertical position
                    // We compare the Top of the manual shape with the Top of existing row anchors
                    int insertIndex = Project.Tasks.Count;

                    // We need to know the current Top of the Gantt chart to map pixels to rows
                    double timelineTop = 100 + GanttRenderer.GetHeaderHeight(Project); // Base offset + Header
                    if (group != null) timelineTop = group.Top + GanttRenderer.GetHeaderHeight(Project);

                    insertIndex = GanttRenderer.GetTaskIndexAtVerticalOffset(Project, shape.Top - timelineTop);
                    if (insertIndex < 0) insertIndex = Project.Tasks.Count;

                    if (insertIndex > Project.Tasks.Count) insertIndex = Project.Tasks.Count;

                    Project.Tasks.Insert(insertIndex, newTask);
                    addedAny = true;
                }

                if (addedAny)
                {
                    Project.RecalculateWBS();
                    StatusMessage = "Selected shapes adopted into Gantt chart.";
                    RequestUpdate(true);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttPanelViewModel.AdoptSelection");
                StatusMessage = "Error adopting selection: " + ex.Message;
            }
            finally
            {
                if (group != null) Marshal.ReleaseComObject(group);
                _isCalculating = false;
            }
        }

        private static float GetSlideWidthOrDefault(
            Slide slide,
            float fallback,
            string context)
        {
            try
            {
                Presentation presentation = (Presentation)slide.Parent;
                return presentation.PageSetup.SlideWidth;
            }
            catch (COMException)
            {
                return fallback;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, context);
                return fallback;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
