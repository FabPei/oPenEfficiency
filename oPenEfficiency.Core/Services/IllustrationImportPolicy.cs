using System;
using System.IO;
using oPenEfficiency.Models;

namespace oPenEfficiency.Services
{
    /// <summary>Result of screening one file the user picked for import.</summary>
    public sealed class IllustrationImportDecision
    {
        public bool Accepted { get; set; }

        /// <summary>Why the file was rejected - shown to the user, so it stays plain language.</summary>
        public string Reason { get; set; }

        /// <summary>Non-blocking remark, e.g. a missing viewBox.</summary>
        public string Note { get; set; }

        /// <summary>Readable name derived from the file name.</summary>
        public string Name { get; set; }

        public string ViewBox { get; set; }

        public int PaletteSize { get; set; }

        public int PartCount { get; set; }

        public bool Decomposable { get; set; }
    }

    /// <summary>
    /// Host-neutral import rules for user-supplied SVG artwork. The add-in never downloads anything
    /// itself: the user picks files they obtained from a source listed in the Illustration Library
    /// window, and this policy decides whether the file may enter the user library. The license
    /// question stays with the user - the tool only refuses files it cannot render safely.
    /// </summary>
    public static class IllustrationImportPolicy
    {
        /// <summary>Upper bound for one SVG file; bigger artwork is refused so the library stays lean.</summary>
        public const int MaxCharacters = 512 * 1024;

        public static IllustrationImportDecision Evaluate(string fileName, string svgText)
        {
            var decision = new IllustrationImportDecision();

            if (string.IsNullOrWhiteSpace(fileName))
            {
                decision.Reason = "no file name was supplied";
                return decision;
            }

            string extension = Path.GetExtension(fileName);
            if (!string.Equals(extension, ".svg", StringComparison.OrdinalIgnoreCase))
            {
                decision.Reason = "only .svg files can be imported";
                return decision;
            }

            if (string.IsNullOrWhiteSpace(svgText))
            {
                decision.Reason = "the file is empty";
                return decision;
            }

            if (svgText.Length > MaxCharacters)
            {
                decision.Reason = "the file is larger than " + (MaxCharacters / 1024) + " KB";
                return decision;
            }

            if (svgText.IndexOf("<svg", StringComparison.OrdinalIgnoreCase) < 0)
            {
                decision.Reason = "the file does not contain an SVG element";
                return decision;
            }

            var plan = IllustrationSvgPlanner.Describe(svgText);
            if (!plan.IsSafe)
            {
                decision.Reason = "the file contains content the add-in does not insert (" + plan.UnsafeReason + ")";
                return decision;
            }

            if (plan.Palette.Count == 0)
            {
                decision.Reason = "no fill or stroke colours were found, so the artwork could not be recoloured";
                return decision;
            }

            if (string.IsNullOrWhiteSpace(plan.ViewBox))
            {
                decision.Note = "the file has no viewBox; PowerPoint will size it from its width and height attributes";
            }

            decision.Accepted = true;
            decision.Name = IllustrationSvgPlanner.DisplayName(Path.GetFileNameWithoutExtension(fileName));
            decision.ViewBox = plan.ViewBox;
            decision.PaletteSize = plan.Palette.Count;
            decision.PartCount = plan.Parts.Count;
            decision.Decomposable = plan.IsDecomposable;
            return decision;
        }

        /// <summary>
        /// License trail written next to every imported file so the user can retrace where the artwork
        /// came from and which terms applied at import time.
        /// </summary>
        public static string DescribeTrail(IllustrationSource source, string fileName, DateTime importedUtc)
        {
            string sourceName = source == null ? "unknown source" : source.Name;
            string license = source == null ? "unknown license" : source.LicenseName;
            string url = source == null ? string.Empty : source.LicenseUrl;

            return string.Format(
                "{0} | imported {1:yyyy-MM-dd} | {2} | {3} | source license: {4}{5}",
                fileName,
                importedUtc,
                sourceName,
                license,
                url,
                source != null && source.AttributionRequired ? " (credit required in the presentation)" : string.Empty);
        }
    }
}
