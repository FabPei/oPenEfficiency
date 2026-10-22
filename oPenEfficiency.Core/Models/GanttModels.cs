using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Xml.Serialization;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Models
{
    public enum MilestoneShape
    {
        Diamond,
        Square,
        Star,
        Triangle,
        Circle
    }

    public enum TaskShape
    {
        Rectangle,
        RoundedRectangle,
        RightArrow,
        Chevron,
        Pentagon
    }

    public enum DependencyType
    {
        FinishToStart,   // successor starts after predecessor ends (most common)
        StartToStart,    // successor starts when predecessor starts
        FinishToFinish,  // successor finishes when predecessor finishes
    }

    /// <summary>
    /// Stores type and lag metadata for a single predecessor relationship.
    /// </summary>
    public class GanttDependencyLink
    {
        public Guid PredecessorId { get; set; }
        public DependencyType Type { get; set; } = DependencyType.FinishToStart;
        public int LagDays { get; set; } = 0;
    }

    /// <summary>
    /// Represents an extra visual event (milestone) within a task lane.
    /// used for 'Flexible Mode' roadmapping.
    /// </summary>
    public class GanttSubEvent : INotifyPropertyChanged
    {
        private DateTime _date;
        private string _title;
        private MilestoneShape _shape = MilestoneShape.Diamond;
        private string _color;

        public DateTime Date
        {
            get => _date;
            set { _date = value; OnPropertyChanged(); }
        }

        public string Title
        {
            get => _title;
            set { _title = value; OnPropertyChanged(); }
        }

        public MilestoneShape Shape
        {
            get => _shape;
            set { _shape = value; OnPropertyChanged(); }
        }

        public string Color
        {
            get => _color;
            set { _color = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    /// <summary>
    /// Represents an extra task bar (duration) within a task lane.
    /// Used for 'Flexible Mode' roadmapping.
    /// </summary>
    public class GanttSubBar : INotifyPropertyChanged
    {
        private DateTime _start;
        private DateTime _end;
        private string _title;
        private string _color;
        private TaskShape _shape = TaskShape.Rectangle;
        private double _percentComplete;

        public DateTime Start
        {
            get => _start;
            set { _start = value; OnPropertyChanged(); }
        }

        public DateTime End
        {
            get => _end;
            set { _end = value; OnPropertyChanged(); }
        }

        public string Title
        {
            get => _title;
            set { _title = value; OnPropertyChanged(); }
        }

        public string Color
        {
            get => _color;
            set { _color = value; OnPropertyChanged(); }
        }

        public TaskShape Shape
        {
            get => _shape;
            set { _shape = value; OnPropertyChanged(); }
        }

        public double PercentComplete
        {
            get => _percentComplete;
            set { _percentComplete = Math.Max(0, Math.Min(1, value)); OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    /// <summary>
    /// Represents a single task or summary in a professional Gantt project.
    /// </summary>
    public class GanttTask : INotifyPropertyChanged
    {
        private Guid _id;
        private string _title;
        private DateTime _start;
        private DateTime _end;
        private int _outlineLevel;
        private double _percentComplete;
        private string _responsible;
        private MilestoneShape _milestoneShape = MilestoneShape.Diamond;
        private TaskShape _shape = TaskShape.Rectangle;
        private string _customColor;      // Hex fill color
        private string _customTextColor;  // Hex text color for label
        private string _customBarText;     // User-typed text inside the task bar shape
        private float?  _customBarTextSize; // Font size of that text
        private bool?   _labelBold;
        private bool?   _labelItalic;
        private bool?   _labelUnderline;
        private float?  _labelFontSize;
        private string  _labelFontName;
        private Guid? _parentId;
        private List<Guid> _predecessorIds = new List<Guid>();
        private List<GanttDependencyLink> _predecessorLinks = new List<GanttDependencyLink>();
        private ObservableCollection<GanttTask> _children = new ObservableCollection<GanttTask>();
        private ObservableCollection<GanttSubEvent> _subEvents = new ObservableCollection<GanttSubEvent>();
        private ObservableCollection<GanttSubBar> _subBars = new ObservableCollection<GanttSubBar>();

        public GanttTask()
        {
            _children.CollectionChanged += OnChildrenCollectionChanged;
        }

        private void OnChildrenCollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            OnPropertyChanged(nameof(IsSummary));
        }

        public ObservableCollection<GanttSubEvent> SubEvents
        {
            get => _subEvents;
            set { _subEvents = value; OnPropertyChanged(); }
        }

        public ObservableCollection<GanttSubBar> SubBars
        {
            get => _subBars;
            set { _subBars = value; OnPropertyChanged(); }
        }

        private bool _isEventLine;
        public bool IsEventLine
        {
            get => _isEventLine;
            set { _isEventLine = value; OnPropertyChanged(); }
        }

        private bool _isHighlighted;
        public bool IsHighlighted
        {
            get => _isHighlighted;
            set { if (_isHighlighted == value) return; _isHighlighted = value; OnPropertyChanged(); }
        }

        private bool _isCollapsed;
        public bool IsCollapsed
        {
            get => _isCollapsed;
            set { if (_isCollapsed == value) return; _isCollapsed = value; OnPropertyChanged(); }
        }

        [System.Xml.Serialization.XmlIgnore]
        public bool IsVisible { get; set; } = true;

        public MilestoneShape MilestoneShape
        {
            get => _milestoneShape;
            set { _milestoneShape = value; OnPropertyChanged(); }
        }

        public TaskShape Shape
        {
            get => _shape;
            set { _shape = value; OnPropertyChanged(); }
        }

        public string CustomColor
        {
            get => _customColor;
            set { if (_customColor == value) return; _customColor = value; OnPropertyChanged(); }
        }

        public string CustomTextColor
        {
            get => _customTextColor;
            set { if (_customTextColor == value) return; _customTextColor = value; OnPropertyChanged(); }
        }

        public string CustomBarText
        {
            get => _customBarText;
            set { if (_customBarText == value) return; _customBarText = value; OnPropertyChanged(); }
        }

        public float? CustomBarTextSize
        {
            get => _customBarTextSize;
            set { if (_customBarTextSize == value) return; _customBarTextSize = value; OnPropertyChanged(); }
        }

        public bool? LabelBold
        {
            get => _labelBold;
            set { if (_labelBold == value) return; _labelBold = value; OnPropertyChanged(); }
        }

        public bool? LabelItalic
        {
            get => _labelItalic;
            set { if (_labelItalic == value) return; _labelItalic = value; OnPropertyChanged(); }
        }

        public bool? LabelUnderline
        {
            get => _labelUnderline;
            set { if (_labelUnderline == value) return; _labelUnderline = value; OnPropertyChanged(); }
        }

        public float? LabelFontSize
        {
            get => _labelFontSize;
            set { if (_labelFontSize == value) return; _labelFontSize = value; OnPropertyChanged(); }
        }

        public string LabelFontName
        {
            get => _labelFontName;
            set { if (_labelFontName == value) return; _labelFontName = value; OnPropertyChanged(); }
        }

        public Guid Id
        {
            get => _id;
            set { _id = value; OnPropertyChanged(); }
        }

        public string Title
        {
            get => _title;
            set { _title = value; OnPropertyChanged(); }
        }

        public DateTime Start
        {
            get => _start;
            set
            {
                if (_start != value)
                {
                    _start = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsMilestone));
                    OnPropertyChanged(nameof(Duration));
                    OnPropertyChanged(nameof(DurationDays));
                }
            }
        }

        public DateTime End
        {
            get => _end;
            set
            {
                if (_end != value)
                {
                    _end = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsMilestone));
                    OnPropertyChanged(nameof(Duration));
                    OnPropertyChanged(nameof(DurationDays));
                }
            }
        }

        public int OutlineLevel
        {
            get => _outlineLevel;
            set { _outlineLevel = value; OnPropertyChanged(); }
        }

        public double PercentComplete
        {
            get => _percentComplete;
            set { _percentComplete = Math.Max(0, Math.Min(1, value)); OnPropertyChanged(); }
        }

        public string Responsible
        {
            get => _responsible;
            set { _responsible = value; OnPropertyChanged(); }
        }

        public Guid? ParentId
        {
            get => _parentId;
            set { _parentId = value; OnPropertyChanged(); }
        }

        private string _wbs;
        public string WBS
        {
            get => _wbs;
            set { _wbs = value; OnPropertyChanged(); }
        }

        private string _predecessorsString;
        public string PredecessorsString
        {
            get => _predecessorsString;
            set { _predecessorsString = value; OnPropertyChanged(); }
        }

        public List<Guid> PredecessorIds
        {
            get => _predecessorIds;
            set { _predecessorIds = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Stores type and lag for each predecessor. Indexed by the same predecessor IDs as
        /// PredecessorIds. When present, AdjustSuccessors uses these instead of the FS default.
        /// </summary>
        public List<GanttDependencyLink> PredecessorLinks
        {
            get => _predecessorLinks;
            set { _predecessorLinks = value ?? new List<GanttDependencyLink>(); OnPropertyChanged(); }
        }

        private double _totalSlack;
        public double TotalSlack
        {
            get => _totalSlack;
            set { _totalSlack = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsCritical)); }
        }

        [XmlIgnore]
        public bool IsCritical => TotalSlack <= 0.01; // Using epsilon for double comparison

        /// <summary>
        /// Gets the calculated duration in days.
        /// </summary>
        [XmlIgnore]
        public int DurationDays => GanttMathEngine.CountWorkDays(Start, End);

        public ObservableCollection<GanttTask> Children
        {
            get => _children;
            set
            {
                if (_children != null) _children.CollectionChanged -= OnChildrenCollectionChanged;
                _children = value;
                if (_children != null) _children.CollectionChanged += OnChildrenCollectionChanged;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsSummary));
            }
        }

        /// <summary>
        /// A task is a milestone if its duration is 0 (Start equals End).
        /// </summary>
        [XmlIgnore]
        public bool IsMilestone => Start == End;

        /// <summary>
        /// A task is a summary task if it has children.
        /// </summary>
        [XmlIgnore]
        public bool IsSummary => Children != null && Children.Any();

        /// <summary>
        /// Gets the duration of the task.
        /// </summary>
        [XmlIgnore]
        public TimeSpan Duration => End - Start;

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Updates the start and end dates based on the min/max of its children.
        /// </summary>
        public void UpdateFromChildren()
        {
            if (!IsSummary) return;

            // Only consider children that are either milestones or have a valid range
            var validChildren = Children.ToList();
            if (!validChildren.Any()) return;

            var newStart = validChildren.Min(c => c.Start);
            var newEnd = validChildren.Max(c => c.End);

            // Set values directly to avoid redundant calculations if needed,
            // but properties will trigger necessary notifications.
            Start = newStart;
            End = newEnd;
        }
    }

    /// <summary>
    /// Defines the zoom level/granularity of the Gantt chart timeline.
    /// </summary>
    public enum GanttViewMode
    {
        Days,
        Weeks,
        Months,
        Years
    }

    /// <summary>
    /// Defines the color palette for the Gantt chart.
    /// </summary>
    public enum GanttTheme
    {
        StandardOffice,
        HighContrast,
        Pastel
    }

    /// <summary>
    /// Represents a custom background highlight for a specific date range.
    /// </summary>
    public class GanttCustomShade : INotifyPropertyChanged
    {
        private DateTime _start;
        private DateTime _end;
        private string _color = "#E0E0E0"; // Default light grey
        private float _transparency = 0.5f;

        public DateTime Start
        {
            get => _start;
            set { _start = value; OnPropertyChanged(); }
        }

        public DateTime End
        {
            get => _end;
            set { _end = value; OnPropertyChanged(); }
        }

        public string Color
        {
            get => _color;
            set { _color = value; OnPropertyChanged(); }
        }

        public float Transparency
        {
            get => _transparency;
            set { _transparency = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    /// <summary>
    /// Represents a professional Gantt project containing a collection of tasks.
    /// </summary>
    public class GanttProject : INotifyPropertyChanged
    {
        private Guid _chartId = Guid.NewGuid();
        private ObservableCollection<GanttTask> _tasks;
        private ObservableCollection<GanttCustomShade> _customShades = new ObservableCollection<GanttCustomShade>();
        private DateTime _projectStart;
        private DateTime _projectEnd;

        public ObservableCollection<GanttCustomShade> CustomShades
        {
            get => _customShades;
            set { _customShades = value; OnPropertyChanged(); }
        }
        private GanttViewMode _viewMode = GanttViewMode.Days;
        private GanttTheme _theme = GanttTheme.StandardOffice;
        private string _fontFamily = "Segoe UI";
        private bool _showNrColumn = true;
        private bool _showTaskNameColumn = true;
        private bool _showDurationColumn = true;
        private bool _showLegend = true;

        private bool _showYears = true;
        private bool _showMonths = true;
        private bool _showWeeks = false;
        private bool _showDays = true;
        private bool _autoFitToSlide = true;
        private bool _lockElements = false;
        private bool _shadeWeekends = true;
        private bool _followGanttRules = true;
        private string _linkedExcelFilePath;
        private bool _isRecalculating = false;

        /// <summary>
        /// Stable identity of this chart instance. It distinguishes multiple Gantt/timeline
        /// charts on the same slide while remaining intact when the project is serialized.
        /// </summary>
        public Guid ChartId
        {
            get => _chartId;
            set
            {
                _chartId = value == Guid.Empty ? Guid.NewGuid() : value;
                OnPropertyChanged();
            }
        }

        public bool FollowGanttRules
        {
            get => _followGanttRules;
            set
            {
                if (_followGanttRules == value) return;
                _followGanttRules = value;
                OnPropertyChanged();
                if (value) RecalculateWBS();
            }
        }

        public string LinkedExcelFilePath
        {
            get => _linkedExcelFilePath;
            set { _linkedExcelFilePath = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsLinkedToExcel)); }
        }

        [XmlIgnore]
        public bool IsLinkedToExcel => !string.IsNullOrEmpty(LinkedExcelFilePath);

        public bool ShadeWeekends
        {
            get => _shadeWeekends;
            set { _shadeWeekends = value; OnPropertyChanged(); }
        }

        public bool LockElements
        {
            get => _lockElements;
            set { _lockElements = value; OnPropertyChanged(); }
        }

        public bool AutoFitToSlide
        {
            get => _autoFitToSlide;
            set { _autoFitToSlide = value; OnPropertyChanged(); }
        }

        public bool ShowYears
        {
            get => _showYears;
            set { _showYears = value; OnPropertyChanged(); }
        }

        public bool ShowMonths
        {
            get => _showMonths;
            set { _showMonths = value; OnPropertyChanged(); }
        }

        public bool ShowWeeks
        {
            get => _showWeeks;
            set { _showWeeks = value; OnPropertyChanged(); }
        }

        public bool ShowDays
        {
            get => _showDays;
            set { _showDays = value; OnPropertyChanged(); }
        }

        private bool _useManualRange = false;
        private DateTime _manualStart = DateTime.Today;
        private DateTime _manualEnd = DateTime.Today.AddMonths(3);

        public bool UseManualRange
        {
            get => _useManualRange;
            set { _useManualRange = value; OnPropertyChanged(); UpdateProjectBoundaries(); }
        }

        public DateTime ManualStart
        {
            get => _manualStart;
            set
            {
                _manualStart = value.Date;
                if (_manualEnd <= _manualStart)
                {
                    _manualEnd = _manualStart.AddDays(1);
                    OnPropertyChanged(nameof(ManualEnd));
                }
                OnPropertyChanged();
                if (UseManualRange) UpdateProjectBoundaries();
            }
        }

        public DateTime ManualEnd
        {
            get => _manualEnd;
            set
            {
                _manualEnd = value.Date <= _manualStart ? _manualStart.AddDays(1) : value.Date;
                OnPropertyChanged();
                if (UseManualRange) UpdateProjectBoundaries();
            }
        }

        public bool ShowLegend
        {
            get => _showLegend;
            set { _showLegend = value; OnPropertyChanged(); }
        }

        public bool ShowNrColumn
        {
            get => _showNrColumn;
            set { _showNrColumn = value; OnPropertyChanged(); }
        }

        public bool ShowTaskNameColumn
        {
            get => _showTaskNameColumn;
            set { _showTaskNameColumn = value; OnPropertyChanged(); }
        }

        public bool ShowDurationColumn
        {
            get => _showDurationColumn;
            set { _showDurationColumn = value; OnPropertyChanged(); }
        }

        public GanttViewMode ViewMode
        {
            get => _viewMode;
            set { _viewMode = value; OnPropertyChanged(); }
        }

        public GanttTheme Theme
        {
            get => _theme;
            set { _theme = value; OnPropertyChanged(); }
        }

        public string FontFamily
        {
            get => _fontFamily;
            set { _fontFamily = value; OnPropertyChanged(); }
        }

        public ObservableCollection<GanttTask> Tasks
        {
            get => _tasks;
            set
            {
                if (_tasks != null)
                    _tasks.CollectionChanged -= Tasks_CollectionChanged;

                _tasks = value;

                if (_tasks != null)
                    _tasks.CollectionChanged += Tasks_CollectionChanged;

                OnPropertyChanged();
                UpdateProjectBoundaries();
            }
        }

        [XmlIgnore]
        public DateTime ProjectStart
        {
            get => _projectStart;
            private set { _projectStart = value; OnPropertyChanged(); }
        }

        [XmlIgnore]
        public DateTime ProjectEnd
        {
            get => _projectEnd;
            private set { _projectEnd = value; OnPropertyChanged(); }
        }

        public GanttProject()
        {
            Tasks = new ObservableCollection<GanttTask>();
            CustomShades = new ObservableCollection<GanttCustomShade>();

            // Keep the parameterless constructor free of sample data. XmlSerializer invokes
            // it before populating collections, so seeded tasks here would be appended to
            // every loaded project.
            ShowWeeks = true;
            ShadeWeekends = true;
        }

        public static GanttProject CreateDefault()
        {
            var project = new GanttProject();

            // Initialize with default tasks: 5 Tasks + 1 Milestone over two weeks
            DateTime start = DateTime.Today;
            while (start.DayOfWeek != DayOfWeek.Monday) start = start.AddDays(1); // Start on next Monday

            Random rnd = new Random();
            DateTime currentTaskStart = start;

            for (int i = 1; i <= 5; i++)
            {
                int duration = rnd.Next(1, 4); // 1 to 3 days
                var task = new GanttTask
                {
                    Id = Guid.NewGuid(),
                    Title = $"Task {i}",
                    Start = currentTaskStart,
                    End = GanttMathEngine.AddWorkDays(currentTaskStart, duration),
                    OutlineLevel = 0,
                    PercentComplete = 0
                };
                project.Tasks.Add(task);

                // Sequence them slightly overlapping or sequential
                currentTaskStart = GanttMathEngine.AddWorkDays(currentTaskStart, rnd.Next(1, 3));
            }

            // Add Milestone
            DateTime milestoneDate = GanttMathEngine.NextWorkDay(project.Tasks.Max(t => t.End));
            project.Tasks.Add(new GanttTask
            {
                Id = Guid.NewGuid(),
                Title = "Milestone 1",
                Start = milestoneDate,
                End = milestoneDate, // Duration 0 = Milestone
                OutlineLevel = 0,
                PercentComplete = 0
            });

            // Add a default Shade for the first week to show off the feature
            project.CustomShades.Add(new GanttCustomShade
            {
                Start = start,
                End = start.AddDays(7),
                Color = "#4A90E2", // Professional blue
                Transparency = 0.2f
            });

            project.RecalculateWBS();
            return project;
        }

        private void Tasks_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            UpdateProjectBoundaries();
        }

        /// <summary>
        /// Calculates the overall project start and end dates based on all tasks.
        /// </summary>
        public void UpdateProjectBoundaries()
        {
            if (UseManualRange)
            {
                ProjectStart = ManualStart;
                ProjectEnd = ManualEnd <= ManualStart ? ManualStart.AddDays(1) : ManualEnd;
                return;
            }

            if (Tasks == null || !Tasks.Any())
            {
                ProjectStart = DateTime.Today;
                ProjectEnd = DateTime.Today.AddDays(1);
                return;
            }

            // We look at all tasks in the flat collection (if applicable) or recursively.
            // Assuming Tasks is the flat list of all tasks for boundary calculation.
            ProjectStart = Tasks.Min(t => t.Start);
            ProjectEnd = Tasks.Max(t => t.End);
        }

        /// <summary>
        /// Ensures all summary tasks are correctly calculated from their children (bottom-up).
        /// </summary>
        public void RecalculateWBS()
        {
            if (_isRecalculating || Tasks == null) return;
            _isRecalculating = true;
            try
            {
                // Rebuild Children/ParentId from the flat OutlineLevel values so IsSummary,
                // UpdateFromChildren, and WBS numbering all reflect the actual visual hierarchy.
                RebuildHierarchyFromOutlineLevel();

                var rootTasks = Tasks.Where(t => t.ParentId == null).ToList();
                for (int i = 0; i < rootTasks.Count; i++)
                {
                    RecalculateTaskHierarchy(rootTasks[i], (i + 1).ToString());
                }

                // Second pass to build PredecessorsString (includes type/lag when non-default)
                foreach (var task in Tasks)
                {
                    if (task.PredecessorIds != null && task.PredecessorIds.Any())
                    {
                        var parts = task.PredecessorIds.Select(id =>
                        {
                            var predTask = Tasks.FirstOrDefault(t => t.Id == id);
                            if (predTask?.WBS == null) return null;
                            var link = task.PredecessorLinks?.FirstOrDefault(l => l.PredecessorId == id);
                            if (link == null || (link.Type == DependencyType.FinishToStart && link.LagDays == 0))
                                return predTask.WBS;
                            var typeStr = link.Type == DependencyType.StartToStart ? "SS"
                                        : link.Type == DependencyType.FinishToFinish ? "FF" : "FS";
                            var lagStr = link.LagDays > 0 ? "+" + link.LagDays
                                       : link.LagDays < 0 ? link.LagDays.ToString() : "";
                            return predTask.WBS + typeStr + lagStr;
                        }).Where(s => s != null).ToList();
                        task.PredecessorsString = string.Join(", ", parts);
                    }
                    else
                    {
                        task.PredecessorsString = string.Empty;
                    }
                }

                UpdateProjectBoundaries();
            }
            finally
            {
                _isRecalculating = false;
            }
        }

        /// <summary>
        /// Derives Children/ParentId relationships from the flat OutlineLevel on each task.
        /// Must be called before RecalculateTaskHierarchy so IsSummary and UpdateFromChildren work correctly.
        /// </summary>
        private void RebuildHierarchyFromOutlineLevel()
        {
            // Clear all existing parent-child links
            foreach (var task in Tasks)
            {
                task.Children.Clear();
                task.ParentId = null;
            }

            // Walk tasks in order; the parent at each level is the most recent task one level up.
            var levelMap = new Dictionary<int, GanttTask>();

            foreach (var task in Tasks)
            {
                if (task.OutlineLevel > 0 && levelMap.TryGetValue(task.OutlineLevel - 1, out var parent))
                {
                    task.ParentId = parent.Id;
                    parent.Children.Add(task);
                }

                levelMap[task.OutlineLevel] = task;

                // A new task at level N resets any deeper-level context that no longer applies.
                foreach (var k in levelMap.Keys.Where(k => k > task.OutlineLevel).ToList())
                    levelMap.Remove(k);
            }
        }

        private void RecalculateTaskHierarchy(GanttTask task, string wbsPrefix)
        {
            task.WBS = wbsPrefix;

            if (task.Children != null && task.Children.Any())
            {
                for (int i = 0; i < task.Children.Count; i++)
                {
                    RecalculateTaskHierarchy(task.Children[i], $"{wbsPrefix}.{i + 1}");
                }

                // After children are updated, update the parent summary
                if (FollowGanttRules)
                {
                    task.UpdateFromChildren();
                }
            }
        }

        /// <summary>
        /// Calculates the total slack for all tasks and identifies the critical path.
        /// </summary>
        public void CalculateCriticalPath()
        {
            if (Tasks == null || !Tasks.Any()) return;

            var networkTasks = Tasks.Where(task => !task.IsSummary).ToList();
            var taskById = networkTasks.ToDictionary(task => task.Id, task => task);
            var successors = networkTasks.ToDictionary(task => task.Id, task => new List<GanttTask>());
            var indegree = networkTasks.ToDictionary(task => task.Id, task => 0);

            foreach (var successor in networkTasks)
            {
                foreach (Guid predecessorId in successor.PredecessorIds.Distinct())
                {
                    if (!taskById.ContainsKey(predecessorId)) continue;
                    successors[predecessorId].Add(successor);
                    indegree[successor.Id]++;
                }
            }

            var queue = new Queue<GanttTask>(networkTasks.Where(task => indegree[task.Id] == 0));
            var order = new List<GanttTask>();
            while (queue.Count > 0)
            {
                GanttTask task = queue.Dequeue();
                order.Add(task);
                foreach (GanttTask successor in successors[task.Id])
                {
                    indegree[successor.Id]--;
                    if (indegree[successor.Id] == 0) queue.Enqueue(successor);
                }
            }

            if (order.Count != networkTasks.Count)
            {
                foreach (GanttTask task in Tasks) task.TotalSlack = 0;
                return;
            }

            var duration = networkTasks.ToDictionary(
                task => task.Id,
                task => Math.Max(0, GanttMathEngine.CountWorkDays(task.Start, task.End)));
            var earliestStart = networkTasks.ToDictionary(task => task.Id, task => 0d);

            foreach (GanttTask task in order)
            {
                double constrainedStart = 0;
                foreach (Guid predecessorId in task.PredecessorIds.Distinct())
                {
                    if (!taskById.TryGetValue(predecessorId, out GanttTask predecessor)) continue;
                    GanttDependencyLink link = task.PredecessorLinks?.FirstOrDefault(item => item.PredecessorId == predecessorId);
                    DependencyType type = link?.Type ?? DependencyType.FinishToStart;
                    int lag = link?.LagDays ?? 0;
                    double predecessorFinish = earliestStart[predecessorId] + duration[predecessorId];

                    double constraint = type == DependencyType.StartToStart
                        ? earliestStart[predecessorId] + lag
                        : type == DependencyType.FinishToFinish
                            ? predecessorFinish + lag - duration[task.Id]
                            : predecessorFinish + lag;
                    constrainedStart = Math.Max(constrainedStart, constraint);
                }
                earliestStart[task.Id] = constrainedStart;
            }

            double projectFinish = networkTasks.Count == 0
                ? 0
                : networkTasks.Max(task => earliestStart[task.Id] + duration[task.Id]);
            var latestStart = networkTasks.ToDictionary(
                task => task.Id,
                task => projectFinish - duration[task.Id]);

            for (int i = order.Count - 1; i >= 0; i--)
            {
                GanttTask task = order[i];
                foreach (GanttTask successor in successors[task.Id])
                {
                    GanttDependencyLink link = successor.PredecessorLinks?.FirstOrDefault(item => item.PredecessorId == task.Id);
                    DependencyType type = link?.Type ?? DependencyType.FinishToStart;
                    int lag = link?.LagDays ?? 0;
                    double successorFinish = latestStart[successor.Id] + duration[successor.Id];

                    double candidate = type == DependencyType.StartToStart
                        ? latestStart[successor.Id] - lag
                        : type == DependencyType.FinishToFinish
                            ? successorFinish - lag - duration[task.Id]
                            : latestStart[successor.Id] - lag - duration[task.Id];
                    latestStart[task.Id] = Math.Min(latestStart[task.Id], candidate);
                }
            }

            foreach (GanttTask task in networkTasks)
            {
                task.TotalSlack = Math.Max(0, latestStart[task.Id] - earliestStart[task.Id]);
            }

            foreach (GanttTask summary in Tasks.Where(task => task.IsSummary).OrderByDescending(task => task.OutlineLevel))
            {
                summary.TotalSlack = summary.Children.Count == 0
                    ? 0
                    : summary.Children.Min(child => child.TotalSlack);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
