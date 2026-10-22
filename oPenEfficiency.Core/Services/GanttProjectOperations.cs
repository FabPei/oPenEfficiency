using System;
using System.Collections.Generic;
using System.Linq;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Host-neutral structural operations for the flat Gantt task list.
    /// Keeping these operations separate makes hierarchy changes atomic and testable.
    /// </summary>
    public static class GanttProjectOperations
    {
        public static bool DeleteTask(GanttProject project, GanttTask task)
        {
            if (project?.Tasks == null || task == null) return false;

            int taskIndex = project.Tasks.IndexOf(task);
            if (taskIndex < 0) return false;

            int subtreeEnd = GetSubtreeEndIndex(project, taskIndex);

            // Deleting a summary keeps its descendants and promotes the complete subtree
            // by one level. This avoids silently deleting user data or orphaning WBS levels.
            for (int i = taskIndex + 1; i <= subtreeEnd; i++)
            {
                project.Tasks[i].OutlineLevel = Math.Max(0, project.Tasks[i].OutlineLevel - 1);
            }

            project.Tasks.RemoveAt(taskIndex);
            foreach (GanttTask remainingTask in project.Tasks)
            {
                remainingTask.PredecessorIds?.RemoveAll(id => id == task.Id);
                remainingTask.PredecessorLinks?.RemoveAll(link => link.PredecessorId == task.Id);
            }
            project.RecalculateWBS();
            return true;
        }

        public static int GetIndexAfterSubtree(GanttProject project, int rootIndex)
        {
            if (project?.Tasks == null || rootIndex < 0 || rootIndex >= project.Tasks.Count)
                return project?.Tasks?.Count ?? 0;

            return GetSubtreeEndIndex(project, rootIndex) + 1;
        }

        public static bool CanIndentTask(GanttProject project, GanttTask task)
        {
            if (project?.Tasks == null || task == null) return false;

            int index = project.Tasks.IndexOf(task);
            if (index <= 0) return false;

            // A row may only become one level deeper than the row immediately above it.
            return task.OutlineLevel < project.Tasks[index - 1].OutlineLevel + 1;
        }

        public static bool IndentTask(GanttProject project, GanttTask task)
        {
            if (!CanIndentTask(project, task)) return false;

            task.OutlineLevel++;
            project.RecalculateWBS();
            return true;
        }

        /// <summary>
        /// Applies all predecessor constraints together in topological order. Summary rows are
        /// derived from their children and therefore act only as fixed predecessor boundaries.
        /// </summary>
        public static bool RecalculateDependencySchedule(GanttProject project)
        {
            if (project?.Tasks == null || !project.FollowGanttRules) return true;

            project.RecalculateWBS();

            var schedulableTasks = project.Tasks.Where(task => !task.IsSummary).ToList();
            var taskById = project.Tasks.ToDictionary(task => task.Id);
            var schedulableIds = new HashSet<Guid>(schedulableTasks.Select(task => task.Id));
            var inDegree = schedulableTasks.ToDictionary(task => task.Id, task => 0);
            var successors = schedulableTasks.ToDictionary(task => task.Id, task => new List<GanttTask>());

            foreach (GanttTask task in schedulableTasks)
            {
                foreach (Guid predecessorId in (task.PredecessorIds ?? new List<Guid>()).Distinct())
                {
                    if (!schedulableIds.Contains(predecessorId)) continue;
                    inDegree[task.Id]++;
                    successors[predecessorId].Add(task);
                }
            }

            var ready = new Queue<GanttTask>(schedulableTasks.Where(task => inDegree[task.Id] == 0));
            var ordered = new List<GanttTask>();
            while (ready.Count > 0)
            {
                GanttTask task = ready.Dequeue();
                ordered.Add(task);
                foreach (GanttTask successor in successors[task.Id])
                {
                    inDegree[successor.Id]--;
                    if (inDegree[successor.Id] == 0) ready.Enqueue(successor);
                }
            }

            if (ordered.Count != schedulableTasks.Count) return false;

            foreach (GanttTask task in ordered)
            {
                int duration = task.IsMilestone
                    ? 0
                    : Math.Max(1, GanttMathEngine.CountWorkDays(task.Start, task.End));
                DateTime? requiredStart = null;

                foreach (Guid predecessorId in (task.PredecessorIds ?? new List<Guid>()).Distinct())
                {
                    if (!taskById.TryGetValue(predecessorId, out GanttTask predecessor)) continue;

                    GanttDependencyLink link = task.PredecessorLinks?
                        .FirstOrDefault(candidate => candidate.PredecessorId == predecessorId);
                    DependencyType type = link?.Type ?? DependencyType.FinishToStart;
                    int lag = link?.LagDays ?? 0;
                    DateTime constraintStart;

                    switch (type)
                    {
                        case DependencyType.StartToStart:
                            constraintStart = GanttMathEngine.AddWorkDays(
                                GanttMathEngine.NextWorkDay(predecessor.Start), lag);
                            break;
                        case DependencyType.FinishToFinish:
                            DateTime requiredEnd = GanttMathEngine.AddWorkDays(
                                GanttMathEngine.NextWorkDay(predecessor.End), lag);
                            constraintStart = duration == 0
                                ? requiredEnd
                                : GanttMathEngine.AddWorkDays(requiredEnd, -duration);
                            break;
                        default:
                            constraintStart = GanttMathEngine.AddWorkDays(
                                GanttMathEngine.NextWorkDay(predecessor.End), lag);
                            break;
                    }

                    if (!requiredStart.HasValue || constraintStart > requiredStart.Value)
                        requiredStart = constraintStart;
                }

                if (!requiredStart.HasValue) continue;

                task.Start = requiredStart.Value;
                task.End = duration == 0
                    ? requiredStart.Value
                    : GanttMathEngine.AddWorkDays(requiredStart.Value, duration);
            }

            project.RecalculateWBS();
            return true;
        }

        /// <summary>
        /// Moves a task and all contiguous descendants represented by greater outline levels.
        /// Moving downward places the subtree after the target, matching the existing row-drag behavior;
        /// moving upward places it before the target.
        /// </summary>
        public static bool MoveTaskSubtree(GanttProject project, int oldIndex, int targetIndex)
        {
            if (project?.Tasks == null ||
                oldIndex < 0 || oldIndex >= project.Tasks.Count ||
                targetIndex < 0 || targetIndex >= project.Tasks.Count ||
                oldIndex == targetIndex)
            {
                return false;
            }

            int subtreeEnd = GetSubtreeEndIndex(project, oldIndex);

            // Dropping onto a row inside the dragged subtree is a no-op.
            if (targetIndex >= oldIndex && targetIndex <= subtreeEnd) return false;

            var subtree = new List<GanttTask>();
            for (int i = oldIndex; i <= subtreeEnd; i++)
            {
                subtree.Add(project.Tasks[i]);
            }

            for (int i = subtreeEnd; i >= oldIndex; i--)
            {
                project.Tasks.RemoveAt(i);
            }

            int insertionIndex = targetIndex < oldIndex
                ? targetIndex
                : targetIndex - subtree.Count + 1;
            insertionIndex = Math.Max(0, Math.Min(insertionIndex, project.Tasks.Count));

            for (int i = 0; i < subtree.Count; i++)
            {
                project.Tasks.Insert(insertionIndex + i, subtree[i]);
            }

            project.RecalculateWBS();
            return true;
        }

        private static int GetSubtreeEndIndex(GanttProject project, int rootIndex)
        {
            int rootLevel = project.Tasks[rootIndex].OutlineLevel;
            int subtreeEnd = rootIndex;
            while (subtreeEnd + 1 < project.Tasks.Count &&
                   project.Tasks[subtreeEnd + 1].OutlineLevel > rootLevel)
            {
                subtreeEnd++;
            }

            return subtreeEnd;
        }
    }
}
