using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    public sealed class AgendaMutationSnapshot
    {
        public AgendaMutationSnapshot(
            IEnumerable<int> originalSlideIds,
            string originalLayoutTag,
            IEnumerable<string> originalSectionNames = null,
            IDictionary<int, int> originalShapeCounts = null)
        {
            if (originalSlideIds == null) throw new ArgumentNullException(nameof(originalSlideIds));
            OriginalSlideOrder = new List<int>(originalSlideIds);
            OriginalSlideIds = new HashSet<int>(OriginalSlideOrder);
            OriginalLayoutTag = originalLayoutTag;
            OriginalSectionNames = originalSectionNames == null
                ? new List<string>()
                : new List<string>(originalSectionNames);
            OriginalShapeCounts = originalShapeCounts == null
                ? new Dictionary<int, int>()
                : new Dictionary<int, int>(originalShapeCounts);
        }

        public List<int> OriginalSlideOrder { get; }
        public HashSet<int> OriginalSlideIds { get; }
        public string OriginalLayoutTag { get; }
        public List<string> OriginalSectionNames { get; }
        public Dictionary<int, int> OriginalShapeCounts { get; }

        public bool MatchesPresentationStructure(
            IEnumerable<int> slideOrder,
            IEnumerable<string> sectionNames,
            IDictionary<int, int> shapeCounts)
        {
            if (slideOrder == null) throw new ArgumentNullException(nameof(slideOrder));
            if (sectionNames == null) throw new ArgumentNullException(nameof(sectionNames));
            if (shapeCounts == null) throw new ArgumentNullException(nameof(shapeCounts));

            if (!OriginalSlideOrder.SequenceEqual(slideOrder)) return false;
            if (!OriginalSectionNames.SequenceEqual(sectionNames, StringComparer.Ordinal)) return false;
            foreach (var pair in OriginalShapeCounts)
            {
                if (!shapeCounts.TryGetValue(pair.Key, out int shapeCount) ||
                    shapeCount != pair.Value)
                    return false;
            }

            return true;
        }
    }

    public interface IAgendaMutationHost
    {
        AgendaMutationSnapshot Capture();
        bool TryStartUndoEntry();
        bool TryUndo(AgendaMutationSnapshot snapshot);
        void RunFallbackCleanup(AgendaMutationSnapshot snapshot);
    }

    public sealed class AgendaMutationTransaction
    {
        private readonly IAgendaMutationHost _host;
        private AgendaMutationSnapshot _snapshot;
        private bool _undoEntryStarted;
        private bool _mutationAttempted;

        public AgendaMutationTransaction(IAgendaMutationHost host)
        {
            _host = host ?? throw new ArgumentNullException(nameof(host));
        }

        public void Begin(bool requireUndoEntry)
        {
            if (_snapshot != null) throw new InvalidOperationException("The agenda transaction has already started.");

            _snapshot = _host.Capture();
            _undoEntryStarted = _host.TryStartUndoEntry();
            if (requireUndoEntry && !_undoEntryStarted)
            {
                throw new InvalidOperationException(
                    "Agenda Creator could not start a safe PowerPoint Undo transaction. The existing agenda was not changed.");
            }
        }

        public void MarkMutationAttempted()
        {
            if (_snapshot == null) throw new InvalidOperationException("The agenda transaction has not started.");
            _mutationAttempted = true;
        }

        public void RollBack()
        {
            if (_snapshot == null || !_mutationAttempted) return;
            if (_undoEntryStarted && _host.TryUndo(_snapshot)) return;
            _host.RunFallbackCleanup(_snapshot);
        }
    }

    public sealed class PowerPointAgendaMutationHost : IAgendaMutationHost
    {
        private const string AgendaLayoutTag = "oPE_AgendaLayout";
        private readonly Presentation _presentation;

        public PowerPointAgendaMutationHost(Presentation presentation)
        {
            _presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
        }

        public AgendaMutationSnapshot Capture()
        {
            var slideIds = new List<int>();
            var shapeCounts = new Dictionary<int, int>();
            for (int slideIndex = 1; slideIndex <= _presentation.Slides.Count; slideIndex++)
            {
                Slide slide = _presentation.Slides[slideIndex];
                slideIds.Add(slide.SlideID);
                shapeCounts[slide.SlideID] = slide.Shapes.Count;
            }

            var sectionNames = new List<string>();
            for (int sectionIndex = 1; sectionIndex <= _presentation.SectionProperties.Count; sectionIndex++)
                sectionNames.Add(_presentation.SectionProperties.Name(sectionIndex));

            string layoutTag = null;
            if (!ExceptionPolicy.TryComProbe(
                () => _presentation.Tags[AgendaLayoutTag],
                out layoutTag))
            {
                layoutTag = null;
            }

            return new AgendaMutationSnapshot(slideIds, layoutTag, sectionNames, shapeCounts);
        }

        public bool TryStartUndoEntry()
        {
            try
            {
                _presentation.Application.StartNewUndoEntry();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AgendaMutationTransaction.StartUndoEntry");
                return false;
            }
        }

        public bool TryUndo(AgendaMutationSnapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));

            try
            {
                var application = _presentation.Application;
                if (application == null ||
                    application.CommandBars == null ||
                    !application.CommandBars.GetEnabledMso("Undo"))
                {
                    return false;
                }

                application.CommandBars.ExecuteMso("Undo");
                return HasOriginalPresentationStructure(snapshot);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AgendaMutationTransaction.NativeUndo");
                return false;
            }
        }

        public void RunFallbackCleanup(AgendaMutationSnapshot snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));

            ExceptionPolicy.RunBestEffort(
                () => DeleteNewSections(snapshot.OriginalSlideIds),
                "AgendaMutationTransaction.Fallback.Sections",
                ExceptionPolicy.IsExpectedComFailure,
                ExceptionLogger.Log);

            for (int slideIndex = _presentation.Slides.Count; slideIndex >= 1; slideIndex--)
            {
                int cleanupIndex = slideIndex;
                ExceptionPolicy.RunBestEffort(
                    () =>
                    {
                        if (!snapshot.OriginalSlideIds.Contains(_presentation.Slides[cleanupIndex].SlideID))
                            _presentation.Slides[cleanupIndex].Delete();
                    },
                    "AgendaMutationTransaction.Fallback.Slides",
                    ExceptionPolicy.IsExpectedComFailure,
                    ExceptionLogger.Log);
            }

            ExceptionPolicy.RunBestEffort(
                () =>
                {
                    if (string.IsNullOrEmpty(snapshot.OriginalLayoutTag))
                        _presentation.Tags.Delete(AgendaLayoutTag);
                    else
                        _presentation.Tags.Add(AgendaLayoutTag, snapshot.OriginalLayoutTag);
                },
                "AgendaMutationTransaction.Fallback.LayoutTag",
                ExceptionPolicy.IsExpectedComFailure,
                ExceptionLogger.Log);
        }

        private bool HasOriginalPresentationStructure(AgendaMutationSnapshot snapshot)
        {
            var slideOrder = new List<int>();
            var shapeCounts = new Dictionary<int, int>();
            for (int slideIndex = 1; slideIndex <= _presentation.Slides.Count; slideIndex++)
            {
                Slide slide = _presentation.Slides[slideIndex];
                slideOrder.Add(slide.SlideID);
                shapeCounts[slide.SlideID] = slide.Shapes.Count;
            }

            var sectionNames = new List<string>();
            for (int sectionIndex = 1; sectionIndex <= _presentation.SectionProperties.Count; sectionIndex++)
                sectionNames.Add(_presentation.SectionProperties.Name(sectionIndex));

            return snapshot.MatchesPresentationStructure(slideOrder, sectionNames, shapeCounts);
        }

        private void DeleteNewSections(HashSet<int> originalSlideIds)
        {
            dynamic sections = _presentation.SectionProperties;
            for (int sectionIndex = _presentation.SectionProperties.Count; sectionIndex >= 1; sectionIndex--)
            {
                int firstSlideIndex = sections.FirstSlideIndex(sectionIndex);
                if (firstSlideIndex >= 1 &&
                    firstSlideIndex <= _presentation.Slides.Count &&
                    !originalSlideIds.Contains(_presentation.Slides[firstSlideIndex].SlideID))
                {
                    _presentation.SectionProperties.Delete(sectionIndex, false);
                }
            }
        }
    }
}
