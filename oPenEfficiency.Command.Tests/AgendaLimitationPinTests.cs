using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace oPenEfficiency.Command.Tests
{
    /// <summary>
    /// Characterization pins for the accepted FEAT-009 / AG-S0 baseline limitations recorded in
    /// docs/testing/feature-acceptance/regression-ledger.json (LIM-AGENDA-*). The agenda generation
    /// and wizard flows are PowerPoint-COM bound and cannot run host-free, so these pins assert the
    /// source contracts of the pinned behavior, matching the established pattern in
    /// IllustrationLibraryCompletionTests.cs. A deliberate fix-forward (AG-S1+) must update the
    /// ledger scenario together with the code.
    /// </summary>
    [TestClass]
    public sealed class AgendaLimitationPinTests
    {
        // LIM-AGENDA-PAGENUM-001: page numbers are frozen at generation time - they are written once
        // from the divider slide's SlideIndex and never recomputed on save or sync.
        [TestMethod]
        public void PageNumbersAreFrozenAtGenerationTime()
        {
            string generation = ReadRepositoryFile("oPenEfficiency", "Services", "AgendaGenerationService.cs");
            Assert.IsTrue(
                generation.IndexOf("item.PageNumber = dividerSlide.SlideIndex.ToString();", StringComparison.Ordinal) >= 0,
                "Agenda generation no longer writes the page number from the divider slide's SlideIndex at generation time.");

            int pageNumbersWrites = EnumerateAddInSources()
                .Sum(source => Regex.Matches(source, @"\.PageNumber\s*=").Count);
            Assert.AreEqual(1, pageNumbersWrites,
                "PageNumber is written exactly once, at generation time; a save-time refresh would add a " +
                "second write and change the pinned LIM-AGENDA-PAGENUM-001 freeze semantics.");

            string[] refreshMethods = EnumerateAddInSourceFiles()
                .SelectMany(path => Regex.Matches(File.ReadAllText(path), @"(?:public|private|internal|protected)[^\n(]*?\b(\w*PageNumber\w*)\s*\(")
                    .Cast<Match>()
                    .Select(match => match.Groups[1].Value))
                .Distinct()
                .ToArray();
            Assert.AreEqual(0, refreshMethods.Length,
                "Agenda code gained page-number methods (" + string.Join(", ", refreshMethods) + "); " +
                "LIM-AGENDA-PAGENUM-001 must be updated.");
        }

        // LIM-AGENDA-SINGLE-001: a deck carries exactly one agenda definition, held in the single
        // fixed oPE_AgendaMain tag without any per-agenda name or identity slot.
        [TestMethod]
        public void DeckAgendaIdentityIsASingleFixedTag()
        {
            string[] agendaTagLiterals = EnumerateAddInSources()
                .SelectMany(source => Regex.Matches(source, "\"(oPE_AgendaMain[A-Za-z0-9_]*)\"")
                    .Cast<Match>()
                    .Select(match => match.Groups[1].Value))
                .Distinct()
                .OrderBy(literal => literal, StringComparer.Ordinal)
                .ToArray();

            CollectionAssert.AreEqual(
                new[] { "oPE_AgendaMain" },
                agendaTagLiterals,
                "The deck-side agenda identity must stay one fixed tag; a second or parameterized " +
                "oPE_AgendaMain* literal would mean multiple named agendas are modelled and change " +
                "the pinned LIM-AGENDA-SINGLE-001 behavior.");

            string wizard = ReadRepositoryFile("oPenEfficiency", "UI", "Wizards", "AgendaWizard.xaml.cs");
            Assert.IsTrue(
                wizard.IndexOf("currentSlide.Tags.Add(\"oPE_AgendaMain\", \"True\");", StringComparison.Ordinal) >= 0,
                "The agenda slide is no longer written as one fixed tag with the constant value \"True\" - " +
                "the tag carries no agenda-name payload.");
            Assert.IsTrue(
                wizard.IndexOf("slide.Tags[\"oPE_AgendaMain\"]", StringComparison.Ordinal) >= 0,
                "The existing agenda slide is no longer looked up through the single fixed oPE_AgendaMain tag.");
        }

        // LIM-AGENDA-SYNC-001: synchronization is capture-clear-regenerate-restore - manual edits are
        // overwritten without a reviewable diff stage before applying.
        [TestMethod]
        public void AgendaSyncIsCaptureClearRegenerateRestoreWithoutDiffReview()
        {
            string source = ReadRepositoryFile("oPenEfficiency", "Services", "AgendaGenerationService.cs");
            string[] steps =
            {
                "existingContent = GetContentSlideMap(pres);",
                "ClearExistingAgendaSections(pres, oldTopics);",
                "ClearExistingDividerSlides(pres, _targetSlideId);",
                "RestoreContentSlidePositions(pres, newDividerIds, existingContent, items);"
            };

            int previous = -1;
            foreach (string step in steps)
            {
                int index = source.IndexOf(step, StringComparison.Ordinal);
                Assert.IsTrue(index >= 0, "The pinned sync step is missing from GenerateFullAgenda: " + step);
                Assert.IsTrue(index > previous,
                    "The pinned sync order is capture -> clear -> regenerate -> restore; the step is out of order: " + step);
                previous = index;
            }

            Assert.IsFalse(Regex.IsMatch(source, @"\bdiff\b", RegexOptions.IgnoreCase),
                "AgendaGenerationService gained a diff stage; the pinned LIM-AGENDA-SYNC-001 behavior " +
                "regenerates without a diff-first review and must be updated.");
        }

        private static IEnumerable<string> EnumerateAddInSources()
        {
            return EnumerateAddInSourceFiles().Select(File.ReadAllText);
        }

        private static IEnumerable<string> EnumerateAddInSourceFiles()
        {
            string root = FindRepositoryRoot();
            foreach (string directory in new[] { "oPenEfficiency", "oPenEfficiency.Core" })
            {
                string path = Path.Combine(root, directory);
                foreach (string file in Directory.EnumerateFiles(path, "*.cs", SearchOption.AllDirectories))
                {
                    if (file.Contains("\\bin\\") || file.Contains("\\obj\\")) continue;
                    yield return file;
                }
            }
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
