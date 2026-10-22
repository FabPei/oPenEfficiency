using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Gantt.Tests
{
    /// <summary>
    /// Permanent regression coverage for REG-GANTT-RESIZE-ADOPT-001: a widened task bar must adopt
    /// its new end date through slide synchronization. The 2026-09-30 real-host session measured
    /// the failure with a chart rendered at 450/11 pt per day on a 960pt slide while SyncFromSlide
    /// recomputed 610/11 pt per day and rounded the 1.5-day width back to one day. The fix records
    /// the render scale on the chart group and interprets bar geometry in that scale.
    /// </summary>
    [TestClass]
    public sealed class GanttSyncAdoptionTests
    {
        private const double RenderedPointsPerDay = 450.0 / 11;  // fixture render: (700 - 250 table) / 11 days
        private const double RecomputedPointsPerDay = 610.0 / 11; // old sync guess on a 960pt slide: (960 - 100 - 250) / 11

        [TestMethod]
        public void WidenedBarAdoptsTheNewEndInTheRenderedScale()
        {
            double scale = GanttSyncAdoptionPolicy.ResolvePointsPerDay(RenderedPointsPerDay, RecomputedPointsPerDay);
            Assert.AreEqual(RenderedPointsPerDay, scale, 0.0001,
                "The rendered scale must win over the recomputation.");

            var projectStart = new DateTime(2026, 1, 5); // Monday
            double margin = 290; // group left 40 + table 250
            double barLeft = margin + 2 * scale; // task starts on day 2 = 2026-01-07
            double widenedWidth = 1.5 * scale; // the measured 40.9 -> 61.4pt widening

            GanttSyncAdoptionPolicy.AdoptDates(
                barLeft, widenedWidth, margin, scale, projectStart,
                followGanttRules: false, isMilestone: false,
                out DateTime newStart, out DateTime newEnd);

            Assert.AreEqual(new DateTime(2026, 1, 7), newStart);
            Assert.AreEqual(new DateTime(2026, 1, 9), newEnd,
                "A 1.5-day wide bar must adopt a two-day task (end 2026-01-09), not keep the one-day end.");
        }

        [TestMethod]
        public void RecomputedScaleReproducesTheMeasuredResizeFailure()
        {
            // Characterization of the root cause: in the recomputed scale the very same 61.4pt bar
            // reads as ~1.1 days and rounds back to one day - the measured 2026-09-30 symptom
            // ("duration 1->1 d" although the shape width really changed). This is why the recorded
            // render scale must be preferred.
            var projectStart = new DateTime(2026, 1, 5);
            double margin = 290;
            double barLeft = margin + 2 * RenderedPointsPerDay;
            double widenedWidth = 1.5 * RenderedPointsPerDay;

            GanttSyncAdoptionPolicy.AdoptDates(
                barLeft, widenedWidth, margin, RecomputedPointsPerDay, projectStart,
                followGanttRules: false, isMilestone: false,
                out DateTime newStart, out DateTime newEnd);

            Assert.AreEqual(1.0, (newEnd - newStart).TotalDays,
                "In the wrong scale the widened bar keeps its one-day duration - the measured failure.");
            Assert.AreNotEqual(new DateTime(2026, 1, 9), newEnd,
                "The wrong scale must not accidentally produce the fixed two-day result.");
        }

        [TestMethod]
        public void DraggedBarStillAdoptsTheStartInEitherScale()
        {
            var projectStart = new DateTime(2026, 1, 5);
            double margin = 290;

            foreach (double scale in new[] { RenderedPointsPerDay, RecomputedPointsPerDay })
            {
                double draggedLeft = margin + 2 * scale + 30; // the measured +30pt drag
                GanttSyncAdoptionPolicy.AdoptDates(
                    draggedLeft, scale, margin, scale, projectStart,
                    followGanttRules: false, isMilestone: false,
                    out DateTime newStart, out DateTime newEnd);

                Assert.AreEqual(new DateTime(2026, 1, 8), newStart,
                    "A 30pt drag must still adopt a one-day start shift in either scale.");
                Assert.AreEqual(new DateTime(2026, 1, 9), newEnd);
            }
        }

        [TestMethod]
        public void MilestoneAdoptsFromItsCenterWithZeroDuration()
        {
            var projectStart = new DateTime(2026, 1, 5);
            double margin = 290;
            double scale = RenderedPointsPerDay;
            double milestoneLeft = margin + 3 * scale - scale / 2; // center on day 3

            GanttSyncAdoptionPolicy.AdoptDates(
                milestoneLeft, scale, margin, scale, projectStart,
                followGanttRules: false, isMilestone: true,
                out DateTime newStart, out DateTime newEnd);

            Assert.AreEqual(new DateTime(2026, 1, 8), newStart);
            Assert.AreEqual(newStart, newEnd, "A milestone adopts zero duration from its center.");
        }

        [TestMethod]
        public void RuleModeSnapsAdoptedDatesToWorkDays()
        {
            var projectStart = new DateTime(2026, 1, 5); // Monday
            double margin = 290;
            double scale = RenderedPointsPerDay;

            // A four-day bar starting Thursday ends on Saturday and must snap back to Monday.
            double barLeft = margin + 3 * scale;
            GanttSyncAdoptionPolicy.AdoptDates(
                barLeft, 4 * scale, margin, scale, projectStart,
                followGanttRules: true, isMilestone: false,
                out DateTime newStart, out DateTime newEnd);

            Assert.AreEqual(new DateTime(2026, 1, 8), newStart);
            Assert.AreEqual(new DateTime(2026, 1, 12), newEnd,
                "In rule mode the adopted end must snap to the next work day.");

            // A zero-day width collapses onto the start day.
            GanttSyncAdoptionPolicy.AdoptDates(
                barLeft, 0, margin, scale, projectStart,
                followGanttRules: true, isMilestone: false,
                out DateTime collapsedStart, out DateTime collapsedEnd);

            Assert.AreEqual(new DateTime(2026, 1, 8), collapsedStart);
            Assert.AreEqual(collapsedStart, collapsedEnd);
        }

        [TestMethod]
        public void FloatQuantizedHalfDayWidthStillAdoptsTwoDays()
        {
            // The real host stores shape.Width as float32, so a bar grown to exactly 1.5 days
            // reads back as e.g. 1.49967 days (61.35pt at the measured 450/11 scale). The day
            // rounding must tolerate that ripple, or the resize still adopts one day.
            var projectStart = new DateTime(2026, 1, 5);
            double margin = 290;
            double scale = RenderedPointsPerDay;
            double barLeft = margin + 3 * scale; // the measured day-3 bar position

            double[] widenedWidths =
            {
                (double)(float)((float)40.9f * 1.5f),   // the measured 61.35pt float32 width
                (double)(float)((float)(450.0 / 11) * 1.5f),
                1.5 * (450.0 / 11)                      // the exact 1.5-day width
            };

            foreach (double widenedWidth in widenedWidths)
            {
                GanttSyncAdoptionPolicy.AdoptDates(
                    barLeft, widenedWidth, margin, scale, projectStart,
                    followGanttRules: false, isMilestone: false,
                    out DateTime newStart, out DateTime newEnd);

                Assert.AreEqual(new DateTime(2026, 1, 8), newStart);
                Assert.AreEqual(new DateTime(2026, 1, 10), newEnd,
                    $"A float32-quantized 1.5-day bar ({widenedWidth}pt) must adopt a two-day task.");
            }
        }

        [TestMethod]
        public void GenuineSubHalfDayWidthsStillRoundDown()
        {
            // The tolerance must not swallow genuine sub-half-day widths: 1.48 days stays one day.
            var projectStart = new DateTime(2026, 1, 5);
            double margin = 290;
            double scale = RenderedPointsPerDay;
            double barLeft = margin + 3 * scale;

            GanttSyncAdoptionPolicy.AdoptDates(
                barLeft, 1.48 * scale, margin, scale, projectStart,
                followGanttRules: false, isMilestone: false,
                out DateTime newStart, out DateTime newEnd);

            Assert.AreEqual(new DateTime(2026, 1, 8), newStart);
            Assert.AreEqual(new DateTime(2026, 1, 9), newEnd,
                "A genuine 1.48-day width must still round down to one day.");
        }

        [TestMethod]
        public void AdoptionUsesTheRenderTimeOriginNotTheDriftedBoundary()
        {
            // ProjectStart is a live boundary (the earliest task start) and drifts while tasks move.
            // Bar geometry is anchored to the render-time origin, so the widened bar maps to its
            // two-day task only against that origin; against the drifted one it lands elsewhere
            // (the 2026-10-02 host trace showed exactly this shift mid-sync).
            double margin = 290;
            double scale = RenderedPointsPerDay;
            var renderTimeOrigin = new DateTime(2026, 1, 5);
            double barLeft = margin + 2 * scale;
            double widenedWidth = 1.5 * scale;

            GanttSyncAdoptionPolicy.AdoptDates(
                barLeft, widenedWidth, margin, scale, renderTimeOrigin,
                followGanttRules: true, isMilestone: false,
                out DateTime anchoredStart, out DateTime anchoredEnd);

            Assert.AreEqual(new DateTime(2026, 1, 7), anchoredStart);
            Assert.AreEqual(new DateTime(2026, 1, 9), anchoredEnd,
                "Against the render-time origin the widened bar is a two-day task.");

            var driftedOrigin = new DateTime(2026, 1, 7);
            GanttSyncAdoptionPolicy.AdoptDates(
                barLeft, widenedWidth, margin, scale, driftedOrigin,
                followGanttRules: true, isMilestone: false,
                out DateTime driftedStart, out DateTime driftedEnd);

            Assert.AreNotEqual(anchoredEnd, driftedEnd,
                "Against a drifted live boundary the same bar maps to different dates - the reason the sync must use the recorded origin.");
        }

        [TestMethod]
        public void DependencyReschedulePreservesTheAdoptedDuration()
        {
            // SyncFromSlide runs the global dependency reschedule after adopting a bar
            // (AdjustSuccessors -> RecalculateDependencySchedule). The adopted two-day duration
            // must survive that re-placement.
            var project = new GanttProject { FollowGanttRules = true };
            var t1 = new GanttTask
            {
                Id = Guid.NewGuid(),
                Title = "T1",
                Start = new DateTime(2026, 1, 7),
                End = new DateTime(2026, 1, 8)
            };
            var t2 = new GanttTask
            {
                Id = Guid.NewGuid(),
                Title = "T2",
                Start = new DateTime(2026, 1, 7),
                End = new DateTime(2026, 1, 9) // the adopted two-day resize result
            };
            t2.PredecessorIds.Add(t1.Id);
            t2.PredecessorLinks.Add(new GanttDependencyLink
            {
                PredecessorId = t1.Id,
                Type = DependencyType.FinishToStart,
                LagDays = 0
            });
            project.Tasks.Add(t1);
            project.Tasks.Add(t2);
            project.RecalculateWBS();

            GanttProjectOperations.RecalculateDependencySchedule(project);

            Assert.AreEqual(2, GanttMathEngine.CountWorkDays(t2.Start, t2.End),
                "The dependency reschedule must keep the adopted two-day working duration " +
                "(it re-places constrained tasks and schedules in working days, so the calendar span may grow past a weekend).");
            Assert.IsTrue((t2.End - t2.Start).TotalDays > 1,
                "A one-day duration after the reschedule would mean the adopted resize was lost (the measured failure).");
        }

        [TestMethod]
        public void ScaleResolutionPrefersTheRecordedRenderScale()
        {
            Assert.AreEqual(40.9, GanttSyncAdoptionPolicy.ResolvePointsPerDay(40.9, 55.5), 0.0001);
            Assert.AreEqual(55.5, GanttSyncAdoptionPolicy.ResolvePointsPerDay(0, 55.5), 0.0001,
                "Legacy charts without a recorded scale fall back to the recomputation.");
            Assert.AreEqual(55.5, GanttSyncAdoptionPolicy.ResolvePointsPerDay(-1, 55.5), 0.0001);
        }

        [TestMethod]
        public void SyncAndResizePathsReadTheRecordedRenderScale()
        {
            string root = FindRepositoryRoot();

            string renderer = File.ReadAllText(Path.Combine(root, "oPenEfficiency", "Services", "GanttRenderer.cs"));
            int setChartId = renderer.IndexOf("PowerPointShapeManager.SetChartId(group, chartId, true);", StringComparison.Ordinal);
            int setScale = renderer.IndexOf("PowerPointShapeManager.SetGanttScale(group, pointsPerDay);", StringComparison.Ordinal);
            Assert.IsTrue(setChartId >= 0, "GanttRenderer no longer tags the chart group with its chart id.");
            Assert.IsTrue(setScale > setChartId,
                "The render must record the points-per-day scale on the chart group it creates.");

            string viewModel = File.ReadAllText(Path.Combine(root, "oPenEfficiency", "UI", "Panels", "GanttPanelViewModel.cs"));
            Assert.IsTrue(
                viewModel.IndexOf("GanttSyncAdoptionPolicy.ResolvePointsPerDay(recordedPointsPerDay, pointsPerDay)", StringComparison.Ordinal) >= 0,
                "The scale resolution is no longer applied to the slide-sync/resize paths.");
            Assert.IsTrue(
                viewModel.IndexOf("GanttSyncAdoptionPolicy.AdoptDates(", StringComparison.Ordinal) >= 0,
                "SyncFromSlide no longer delegates its date adoption to the tested policy.");
            Assert.IsTrue(
                viewModel.IndexOf("GanttSyncAdoptionPolicy.RoundDays(", StringComparison.Ordinal) >= 0,
                "The interactive resize path no longer shares the float-tolerant day rounding.");
            Assert.IsTrue(
                renderer.IndexOf("PowerPointShapeManager.SetGanttProjectStart(group, projectStart);", StringComparison.Ordinal) > setChartId,
                "The render must record the timeline day-zero on the chart group it creates.");
            Assert.IsTrue(
                viewModel.IndexOf("PowerPointShapeManager.GetGanttProjectStart(", StringComparison.Ordinal) >= 0,
                "The sync/resize paths no longer map bar geometry against the recorded render-time origin.");
            Assert.IsFalse(
                viewModel.Contains("Math.Round(durationDays)"),
                "The untested inline adoption math returned to GanttPanelViewModel; keep it in GanttSyncAdoptionPolicy.");

            string shapeManager = File.ReadAllText(Path.Combine(root, "oPenEfficiency", "Services", "PowerPointShapeManager.cs"));
            Assert.IsTrue(shapeManager.Contains("\"oPE_GanttPointsPerDay\""),
                "The recorded-scale tag constant disappeared from PowerPointShapeManager.");
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
