using System;
using System.Linq;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace oPenEfficiency.Gantt.Tests
{
    [TestClass]
    public class Program
    {
        [TestMethod] public void DeleteChildRebuildsSummaryTest() => DeleteChildRebuildsSummary();
        [TestMethod] public void DeleteSummaryPromotesSubtreeTest() => DeletingSummaryPromotesItsSubtree();
        [TestMethod] public void InsertAfterSummarySkipsSubtreeTest() => InsertAfterSummarySkipsItsSubtree();
        [TestMethod] public void IndentRequiresParentTest() => IndentRequiresAValidParentAbove();
        [TestMethod] public void MoveSummaryDownKeepsSubtreeTest() => MovingSummaryKeepsSubtreeTogether();
        [TestMethod] public void MoveSummaryUpKeepsSubtreeTest() => MovingSummaryUpKeepsSubtreeTogether();
        [TestMethod] public void ReorderUndoIsAtomicTest() => ReorderUndoIsAtomic();
        [TestMethod] public void ChartIdentitySurvivesUndoTest() => ChartIdentitySurvivesUndoSerialization();
        [TestMethod] public void ProjectsHaveDistinctIdentitiesTest() => IndependentProjectsHaveDistinctChartIdentities();
        [TestMethod] public void WorkingDayMathUsesExclusiveEndTest() => WorkingDayMathUsesExclusiveEndDates();
        [TestMethod] public void CriticalPathUsesDependenciesTest() => CriticalPathUsesDependencyNetwork();
        [TestMethod] public void ManualRangeIsValidTest() => ManualRangeIsAlwaysValid();
        [TestMethod] public void GanttRulesCanBeReEnabledTest() => GanttRulesCanBeReEnabled();
        [TestMethod] public void MultiplePredecessorsAreCombinedTest() => MultiplePredecessorsAreEvaluatedTogether();
        [TestMethod] public void LegacyGanttFixtureLoadsTest() => LegacyGanttProjectFixtureLoadsThroughPersistenceFormat();
        [TestMethod] public void LegacyGanttFixtureRoundTripsTest() => LegacyGanttProjectFixtureRoundTripsStableIdentity();

        private static void DeleteChildRebuildsSummary()
        {
            DateTime start = new DateTime(2026, 1, 5);
            var project = new GanttProject();
            project.Tasks.Add(NewTask("Parent", 0, start));
            project.Tasks.Add(NewTask("Only child", 1, start));
            project.RecalculateWBS();
            GanttTask parent = project.Tasks[0];
            GanttTask child = project.Tasks[1];

            Assert(parent.IsSummary, "Precondition failed: parent should be a summary.");
            Assert(GanttProjectOperations.DeleteTask(project, child), "Child task was not deleted.");
            Assert(!parent.IsSummary, "Deleting the last child must clear the parent's summary state.");
            Assert(parent.Children.Count == 0, "Deleting a child must remove stale child references.");
        }

        private static void DeletingSummaryPromotesItsSubtree()
        {
            GanttProject project = CreateHierarchy();
            GanttTask summary = project.Tasks[0];
            GanttTask child = project.Tasks[1];
            GanttTask grandchild = project.Tasks[2];
            GanttTask peer = project.Tasks[3];
            peer.PredecessorIds.Add(summary.Id);
            peer.PredecessorLinks.Add(new GanttDependencyLink
            {
                PredecessorId = summary.Id,
                Type = DependencyType.FinishToStart
            });

            Assert(GanttProjectOperations.DeleteTask(project, summary), "Summary task was not deleted.");
            Assert(project.Tasks.Contains(child) && project.Tasks.Contains(grandchild),
                "Deleting a summary must retain all descendants.");
            Assert(child.OutlineLevel == 0 && grandchild.OutlineLevel == 1,
                "Retained descendants must be promoted exactly one level.");
            Assert(grandchild.ParentId == child.Id,
                "Promotion must preserve the descendant hierarchy.");
            Assert(!peer.PredecessorIds.Contains(summary.Id) &&
                   !peer.PredecessorLinks.Any(link => link.PredecessorId == summary.Id),
                "Deleting a task must remove dangling dependency references.");
        }

        private static void InsertAfterSummarySkipsItsSubtree()
        {
            GanttProject project = CreateHierarchy();
            int insertionIndex = GanttProjectOperations.GetIndexAfterSubtree(project, 0);
            Assert(insertionIndex == 3,
                "A new sibling below a summary must be inserted after all descendants.");
        }

        private static void IndentRequiresAValidParentAbove()
        {
            GanttProject project = CreateHierarchy();
            GanttTask first = project.Tasks[0];
            GanttTask peer = project.Tasks[3];

            Assert(!GanttProjectOperations.CanIndentTask(project, first),
                "The first row cannot be indented because no parent exists above it.");
            Assert(GanttProjectOperations.IndentTask(project, peer),
                "A root-level row below another row should be indentable.");
            Assert(peer.OutlineLevel == 1 && peer.ParentId == first.Id,
                "Indent must create a valid one-level parent relationship.");
        }

        private static void MovingSummaryKeepsSubtreeTogether()
        {
            GanttProject project = CreateHierarchy();
            GanttTask parent = project.Tasks[0];
            GanttTask child = project.Tasks[1];
            GanttTask grandchild = project.Tasks[2];

            Assert(GanttProjectOperations.MoveTaskSubtree(project, 0, 4), "Summary subtree was not moved.");
            Assert(project.Tasks[2] == parent && project.Tasks[3] == child && project.Tasks[4] == grandchild,
                "A moved summary must remain directly followed by all descendants.");
            Assert(child.ParentId == parent.Id && grandchild.ParentId == child.Id,
                "Moving a summary must preserve its hierarchy.");
        }

        private static void ReorderUndoIsAtomic()
        {
            GanttProject project = CreateHierarchy();
            Guid[] originalOrder = project.Tasks.Select(task => task.Id).ToArray();
            var undo = new GanttUndoRedoManager();
            undo.SaveState(project);

            undo.ExecuteTransaction(project, () =>
            {
                GanttProjectOperations.MoveTaskSubtree(project, 0, 4);
                // Simulate collection-change callbacks that previously recorded intermediate states.
                undo.SaveState(project);
            });

            GanttProject restored = undo.Undo(project);
            Assert(restored.Tasks.Select(task => task.Id).SequenceEqual(originalOrder),
                "One Undo must restore the complete pre-reorder task list.");
        }

        private static void MovingSummaryUpKeepsSubtreeTogether()
        {
            GanttProject project = CreateHierarchy();
            GanttTask peerA = project.Tasks[3];
            GanttTask peerB = project.Tasks[4];
            peerB.OutlineLevel = 1;
            project.RecalculateWBS();

            Assert(GanttProjectOperations.MoveTaskSubtree(project, 3, 0), "Summary subtree was not moved upward.");
            Assert(project.Tasks[0] == peerA && project.Tasks[1] == peerB,
                "Moving upward must keep a summary directly followed by its descendants.");
            Assert(peerB.ParentId == peerA.Id, "Moving upward must preserve the child-parent relationship.");
        }

        private static void ChartIdentitySurvivesUndoSerialization()
        {
            GanttProject project = CreateHierarchy();
            Guid chartId = project.ChartId;
            var undo = new GanttUndoRedoManager();
            undo.SaveState(project);
            undo.ExecuteTransaction(project, () => project.Tasks[3].Title = "Changed");

            GanttProject restored = undo.Undo(project);
            Assert(restored.ChartId == chartId, "Chart identity must survive project serialization.");
        }

        private static void IndependentProjectsHaveDistinctChartIdentities()
        {
            var first = new GanttProject();
            var second = new GanttProject();
            Assert(first.ChartId != Guid.Empty && second.ChartId != Guid.Empty,
                "Every chart must receive a non-empty identity.");
            Assert(first.ChartId != second.ChartId,
                "Independent charts must not share an identity.");
        }

        private static void WorkingDayMathUsesExclusiveEndDates()
        {
            DateTime friday = new DateTime(2026, 1, 9);
            DateTime monday = new DateTime(2026, 1, 12);
            Assert(GanttMathEngine.AddWorkDays(friday, 1) == monday,
                "One working day after Friday must end on Monday.");
            Assert(GanttMathEngine.CountWorkDays(friday, monday) == 1,
                "Friday-to-Monday must be displayed as one working day.");
            Assert(GanttMathEngine.AddWorkDays(monday, -1) == friday,
                "Subtracting one working day from Monday must yield Friday.");
        }

        private static void CriticalPathUsesDependencyNetwork()
        {
            DateTime start = new DateTime(2026, 1, 5);
            var project = new GanttProject();
            GanttTask first = NewTask("Critical A", 0, start);
            first.End = GanttMathEngine.AddWorkDays(start, 2);
            GanttTask second = NewTask("Critical B", 0, first.End);
            second.End = GanttMathEngine.AddWorkDays(second.Start, 3);
            second.PredecessorIds.Add(first.Id);
            second.PredecessorLinks.Add(new GanttDependencyLink
            {
                PredecessorId = first.Id,
                Type = DependencyType.FinishToStart
            });
            GanttTask parallel = NewTask("Parallel", 0, start);
            parallel.End = GanttMathEngine.AddWorkDays(start, 2);
            project.Tasks.Add(first);
            project.Tasks.Add(second);
            project.Tasks.Add(parallel);
            project.RecalculateWBS();

            project.CalculateCriticalPath();

            Assert(first.IsCritical && second.IsCritical, "Dependency chain should be critical.");
            Assert(!parallel.IsCritical && Math.Abs(parallel.TotalSlack - 3) < 0.01,
                "Shorter parallel task should have three working days of slack.");
        }

        private static void ManualRangeIsAlwaysValid()
        {
            var project = new GanttProject { UseManualRange = true };
            project.ManualStart = new DateTime(2026, 2, 10);
            project.ManualEnd = new DateTime(2026, 2, 1);
            Assert(project.ManualEnd == project.ManualStart.AddDays(1),
                "Manual end must be normalized to at least one day after start.");
        }

        private static void GanttRulesCanBeReEnabled()
        {
            var project = new GanttProject();
            project.FollowGanttRules = false;
            project.FollowGanttRules = true;
            Assert(project.FollowGanttRules, "Rule-based mode must be reversible.");
        }

        private static void MultiplePredecessorsAreEvaluatedTogether()
        {
            DateTime monday = new DateTime(2026, 1, 5);
            var project = new GanttProject();
            GanttTask early = NewTask("Early predecessor", 0, monday);
            early.End = GanttMathEngine.AddWorkDays(monday, 1);
            GanttTask late = NewTask("Late predecessor", 0, monday);
            late.End = GanttMathEngine.AddWorkDays(monday, 4);
            GanttTask successor = NewTask("Joint successor", 0, monday);
            successor.End = GanttMathEngine.AddWorkDays(monday, 2);
            successor.PredecessorIds.Add(early.Id);
            successor.PredecessorIds.Add(late.Id);
            successor.PredecessorLinks.Add(new GanttDependencyLink
            {
                PredecessorId = early.Id,
                Type = DependencyType.FinishToStart
            });
            successor.PredecessorLinks.Add(new GanttDependencyLink
            {
                PredecessorId = late.Id,
                Type = DependencyType.FinishToStart
            });
            project.Tasks.Add(early);
            project.Tasks.Add(late);
            project.Tasks.Add(successor);

            Assert(GanttProjectOperations.RecalculateDependencySchedule(project),
                "A valid dependency network should be schedulable.");
            Assert(successor.Start == late.End,
                "The latest predecessor constraint must determine the successor start.");
            Assert(GanttMathEngine.CountWorkDays(successor.Start, successor.End) == 2,
                "Joint scheduling must preserve the successor duration.");
        }

        private static GanttProject LoadLegacyFixture()
        {
            string fixturePath = System.IO.Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "Fixtures", "legacy-gantt-project.xml");
            var serializer = new System.Xml.Serialization.XmlSerializer(typeof(GanttProject));
            using (var reader = new System.IO.StreamReader(fixturePath))
            {
                return (GanttProject)serializer.Deserialize(reader);
            }
        }

        private static void LegacyGanttProjectFixtureLoadsThroughPersistenceFormat()
        {
            GanttProject project = LoadLegacyFixture();

            Assert(project.ChartId == Guid.Parse("11111111-1111-1111-1111-111111111111"),
                "Fixture ChartId must load exactly as persisted.");
            Assert(project.Tasks.Count == 6, "Fixture must contain six rows.");

            GanttTask summary = project.Tasks[0];
            GanttTask child = project.Tasks[1];
            GanttTask grandchild = project.Tasks[2];
            GanttTask peer = project.Tasks[3];
            GanttTask milestone = project.Tasks[4];
            GanttTask eventLine = project.Tasks[5];

            Assert(child.ParentId == summary.Id && grandchild.ParentId == child.Id,
                "Hierarchy must load from ParentId.");
            Assert(child.OutlineLevel == 1 && grandchild.OutlineLevel == 2,
                "Outline levels must load.");
            Assert(milestone.IsMilestone && milestone.MilestoneShape == MilestoneShape.Star,
                "Zero-duration row must load as a Star milestone.");
            Assert(eventLine.IsEventLine, "Event-line row must load.");
            Assert(peer.PredecessorIds.Count == 1 &&
                   peer.PredecessorLinks.Count == 1 &&
                   peer.PredecessorLinks[0].PredecessorId == summary.Id &&
                   peer.PredecessorLinks[0].Type == DependencyType.FinishToStart,
                "Dependency link must load with type metadata.");
            Assert(child.SubEvents.Count == 1 && child.SubEvents[0].Date == new DateTime(2026, 1, 6),
                "Sub-event must load.");
            Assert(peer.SubBars.Count == 1 && peer.SubBars[0].Title == "Wave 1",
                "Sub-bar must load.");
            Assert(project.CustomShades.Count == 1 &&
                   project.CustomShades[0].Start == new DateTime(2026, 1, 8) &&
                   project.CustomShades[0].Color == "#D9E2F3",
                "Custom shade must load.");
            Assert(project.ShowWeeks, "Non-default timescale setting must load.");
            Assert(project.ShadeWeekends && project.FollowGanttRules,
                "Weekend shading and rule mode must load.");
            Assert(!string.IsNullOrEmpty(summary.CustomColor) && summary.CustomColor == "#2E75B6",
                "Custom bar color must load.");
            Assert(!string.IsNullOrEmpty(peer.CustomBarText) && peer.CustomBarText == "rollout",
                "Custom bar text must load.");
        }

        private static void LegacyGanttProjectFixtureRoundTripsStableIdentity()
        {
            GanttProject project = LoadLegacyFixture();
            Guid chartId = project.ChartId;
            Guid summaryId = project.Tasks[0].Id;

            var serializer = new System.Xml.Serialization.XmlSerializer(typeof(GanttProject));
            using (var writer = new System.IO.StringWriter())
            {
                serializer.Serialize(writer, project);
                using (var reader = new System.IO.StringReader(writer.ToString()))
                {
                    var restored = (GanttProject)serializer.Deserialize(reader);
                    Assert(restored.ChartId == chartId,
                        "ChartId must survive a persistence round-trip.");
                    Assert(restored.Tasks[0].Id == summaryId && restored.Tasks.Count == 6,
                        "Task identities must survive a persistence round-trip.");
                    Assert(restored.Tasks[3].PredecessorLinks.Count == 1 &&
                           restored.Tasks[3].PredecessorLinks[0].LagDays == 0,
                        "Dependency metadata must survive a persistence round-trip.");
                }
            }
        }

        private static GanttProject CreateHierarchy()
        {
            DateTime start = new DateTime(2026, 1, 5);
            var project = new GanttProject();
            project.Tasks.Add(NewTask("Parent", 0, start));
            project.Tasks.Add(NewTask("Child", 1, start));
            project.Tasks.Add(NewTask("Grandchild", 2, start));
            project.Tasks.Add(NewTask("Peer A", 0, start));
            project.Tasks.Add(NewTask("Peer B", 0, start));
            project.RecalculateWBS();
            return project;
        }

        private static GanttTask NewTask(string title, int outlineLevel, DateTime start)
        {
            return new GanttTask
            {
                Id = Guid.NewGuid(),
                Title = title,
                OutlineLevel = outlineLevel,
                Start = start,
                End = start.AddDays(1)
            };
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }
    }
}
