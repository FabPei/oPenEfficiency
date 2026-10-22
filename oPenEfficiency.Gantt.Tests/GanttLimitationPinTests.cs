using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Gantt.Tests
{
    /// <summary>
    /// Characterization pins for the accepted FEAT-010 / TL-S0 baseline limitations recorded in
    /// docs/testing/feature-acceptance/regression-ledger.json (LIM-GANTT-*). Each test asserts the
    /// pinned current behavior so that a deliberate fix-forward (TL-S3 / TL-S7 / FEAT-005) has to
    /// update the ledger scenario together with the code. COM-bound surfaces that cannot run
    /// host-free are pinned through their source contracts, matching the established pattern in
    /// oPenEfficiency.Command.Tests/IllustrationLibraryCompletionTests.cs.
    /// </summary>
    [TestClass]
    public sealed class GanttLimitationPinTests
    {
        // LIM-GANTT-UNDO-001: undo is model-state snapshot based - one whole-model snapshot per
        // transaction, restored by deserialization, with no per-operation inverse commands.
        [TestMethod]
        public void UndoGranularityIsOneWholeModelSnapshotPerTransaction()
        {
            DateTime start = new DateTime(2026, 1, 5);
            var project = NewProject(start);
            Guid firstTaskId = project.Tasks[0].Id;
            Guid secondTaskId = project.Tasks[1].Id;
            Guid thirdTaskId = project.Tasks[2].Id;

            var undo = new GanttUndoRedoManager();
            undo.SaveState(project);
            undo.ExecuteTransaction(project, () =>
            {
                project.Tasks[0].Title = "Renamed";
                project.Tasks[0].PercentComplete = 0.5;
                GanttProjectOperations.MoveTaskSubtree(project, 0, 2);
            });
            undo.ExecuteTransaction(project, () =>
                project.Tasks.Single(task => task.Id == firstTaskId).Start = start.AddDays(3));

            GanttProject afterSecondUndo = undo.Undo(project);
            Assert.AreNotSame(project, afterSecondUndo,
                "Undo must restore a deserialized model snapshot, not mutate the live model in place.");
            GanttTask renamed = afterSecondUndo.Tasks.Single(task => task.Id == firstTaskId);
            Assert.AreEqual(start, renamed.Start,
                "The second transaction must be undone as one snapshot.");
            Assert.AreEqual("Renamed", renamed.Title,
                "The first transaction must stay applied while only the second snapshot is undone.");
            Assert.AreEqual(0.5, renamed.PercentComplete);

            GanttProject restored = undo.Undo(afterSecondUndo);
            CollectionAssert.AreEqual(
                new[] { firstTaskId, secondTaskId, thirdTaskId },
                restored.Tasks.Select(task => task.Id).ToArray(),
                "One Undo must restore the complete pre-transaction task list in a single step.");
            Assert.AreEqual("Parent", restored.Tasks.Single(task => task.Id == firstTaskId).Title);
            Assert.AreEqual(0.0, restored.Tasks.Single(task => task.Id == firstTaskId).PercentComplete,
                "One Undo must restore the complete pre-transaction model state in a single step.");
        }

        [TestMethod]
        public void UndoRedoManagerExposesNoOperationLevelInverse()
        {
            string[] methods = typeof(GanttUndoRedoManager)
                .GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
                .Where(method => !method.IsSpecialName)
                .Select(method => method.Name)
                .ToArray();

            CollectionAssert.AreEquivalent(
                new[] { "SaveState", "ExecuteTransaction", "Undo", "Redo" },
                methods,
                "The undo model is whole-project snapshots; an operation-level inverse API would " +
                "change the pinned LIM-GANTT-UNDO-001 granularity and must update the ledger scenario.");
        }

        // LIM-GANTT-WEEKEND-001: the weekend definition is fixed to Monday-Friday.
        [TestMethod]
        public void WeekendDefinitionIsFixedToMondayFriday()
        {
            var monday = new DateTime(2026, 10, 5);
            for (int offset = 0; offset < 7; offset++)
            {
                DateTime day = monday.AddDays(offset);
                bool expectedWorkingDay = offset < 5;
                Assert.AreEqual(expectedWorkingDay, GanttMathEngine.IsWorkDay(day),
                    $"{day:yyyy-MM-dd} ({day.DayOfWeek}) must be a working day exactly when it is Monday-Friday.");
            }

            Assert.AreEqual(new DateTime(2026, 10, 5), GanttMathEngine.NextWorkDay(new DateTime(2026, 10, 3)).Date,
                "NextWorkDay from Saturday must land on Monday.");
            Assert.AreEqual(new DateTime(2026, 10, 2), GanttMathEngine.PreviousWorkDay(new DateTime(2026, 10, 4)).Date,
                "PreviousWorkDay from Sunday must land on Friday.");
            Assert.AreEqual(new DateTime(2026, 10, 5), GanttMathEngine.AddWorkDays(new DateTime(2026, 10, 2), 1).Date,
                "Adding one work day to Friday must skip the fixed weekend.");
            Assert.AreEqual(1, GanttMathEngine.CountWorkDays(new DateTime(2026, 10, 2), new DateTime(2026, 10, 5)),
                "A Friday-to-Monday task spans one working day in the fixed Mon-Fri model.");
        }

        [TestMethod]
        public void WeekendPolicyIsNotConfigurable()
        {
            foreach (MethodInfo method in typeof(GanttMathEngine).GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly))
            {
                foreach (ParameterInfo parameter in method.GetParameters())
                {
                    Assert.IsFalse(
                        Regex.IsMatch(parameter.Name ?? string.Empty, "calendar|weekend|holiday|fiscal", RegexOptions.IgnoreCase),
                        $"{method.Name} takes a '{parameter.Name}' parameter, so the Mon-Fri weekend model is no " +
                        "longer fixed; LIM-GANTT-WEEKEND-001 must be updated.");
                }
            }
        }

        // LIM-GANTT-EXCEL-SYNC-001: the Excel side stays import-based - a remembered file path and a
        // full clear-and-replace re-import, without per-item source lineage or merge/diff semantics.
        [TestMethod]
        public void ExcelInteropIsTemplateExportImportOnly()
        {
            string source = ReadRepositoryFile("oPenEfficiency", "Services", "GanttExcelManager.cs");
            string[] methods = Regex.Matches(source, @"public static [\w<>\[\],\s]+?(\w+)\s*\(")
                .Cast<Match>()
                .Select(match => match.Groups[1].Value)
                .ToArray();

            CollectionAssert.AreEquivalent(
                new[] { "CreateTemplate", "ExportToExcel", "ImportFromExcel" },
                methods,
                "The complete Excel capability is template/export/import; a refresh, relink or sync " +
                "operation would change the pinned LIM-GANTT-EXCEL-SYNC-001 import semantics.");
        }

        [TestMethod]
        public void ExcelSyncIsWholeListReimportWithoutSourceLineage()
        {
            string[] lineageNames = typeof(GanttTask)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Select(property => property.Name)
                .Where(name => Regex.IsMatch(name, "lineage|source|refresh|relink|origin", RegexOptions.IgnoreCase))
                .ToArray();

            Assert.AreEqual(0, lineageNames.Length,
                "GanttTask gained source-lineage state (" + string.Join(", ", lineageNames) + "); " +
                "LIM-GANTT-EXCEL-SYNC-001 must be updated.");

            string viewModel = ReadRepositoryFile("oPenEfficiency", "UI", "Panels", "GanttPanelViewModel.cs");
            Assert.IsTrue(
                viewModel.IndexOf("ImportFromExcel(Project.LinkedExcelFilePath, Project.Tasks)", StringComparison.Ordinal) >= 0,
                "The pinned sync path re-imports the whole task list from the remembered file path.");
            Assert.IsTrue(
                viewModel.IndexOf("Project.Tasks.Clear();", StringComparison.Ordinal) >= 0,
                "The pinned sync path replaces the task list instead of merging imported rows into it.");
        }

        // LIM-GANTT-FULL-REDRAW-001: every render is a full redraw - the chart is cleared and fully
        // rebuilt, and no incremental update path exists.
        [TestMethod]
        public void RendererIsFullRedrawWithNoIncrementalUpdatePath()
        {
            string source = ReadRepositoryFile("oPenEfficiency", "Services", "GanttRenderer.cs");
            int clear = source.IndexOf("PowerPointShapeManager.ClearGanttShapes(slide, project.ChartId);", StringComparison.Ordinal);
            int draw = source.IndexOf("DrawHeaders(", StringComparison.Ordinal);

            Assert.IsTrue(clear >= 0, "GanttRenderer.Render no longer clears the existing chart before drawing.");
            Assert.IsTrue(draw > clear,
                "GanttRenderer.Render must clear all chart shapes before the first draw call - the " +
                "pinned LIM-GANTT-FULL-REDRAW-001 behavior is a full redraw.");

            string[] publicMethods = Regex.Matches(source, @"public static [\w<>\[\],\s]+?(\w+)\s*\(")
                .Cast<Match>()
                .Select(match => match.Groups[1].Value)
                .ToArray();
            string[] incremental = publicMethods
                .Where(name => Regex.IsMatch(name, "update|incremental|patch|diff|refresh", RegexOptions.IgnoreCase))
                .ToArray();

            Assert.AreEqual(0, incremental.Length,
                "GanttRenderer exposes incremental update methods (" + string.Join(", ", incremental) + "); " +
                "LIM-GANTT-FULL-REDRAW-001 must be updated.");
        }

        private static GanttProject NewProject(DateTime start)
        {
            var project = new GanttProject();
            project.Tasks.Add(NewTask("Parent", 0, start));
            project.Tasks.Add(NewTask("Child", 1, start));
            project.Tasks.Add(NewTask("Peer", 0, start));
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

        private static string ReadRepositoryFile(params string[] segments)
        {
            string root = FindRepositoryRoot();
            string path = Path.Combine(new[] { root }.Concat(segments).ToArray());
            return File.ReadAllText(path);
        }

        private static string FindRepositoryRoot()
        {
            var current = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (current != null)
            {
                if (File.Exists(Path.Combine(current.FullName, "oPenEfficiency.slnx"))) return current.FullName;
                current = current.Parent;
            }

            throw new DirectoryNotFoundException("Could not locate the repository root.");
        }
    }
}
