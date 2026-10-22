using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using oPenEfficiency.Models;

namespace oPenEfficiency.Services
{
    /// <summary>The choices one insertion applies to an illustration.</summary>
    public sealed class IllustrationInsertOptions
    {
        /// <summary>Hex-to-hex replacement map built from the palette panel.</summary>
        public Dictionary<string, string> ColorMap { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Part ids the user switched off.</summary>
        public HashSet<string> HiddenPartIds { get; } = new HashSet<string>(StringComparer.Ordinal);

        /// <summary>Drops <c>*-None</c> placeholder groups so empty variants never reach the slide.</summary>
        public bool DropPlaceholders { get; set; } = true;

        /// <summary>True when the inserted shapes should be ungrouped into their parts.</summary>
        public bool Decompose { get; set; }

        /// <summary>Target width on the slide, in points.</summary>
        public float DesiredWidth { get; set; } = 240f;
    }

    /// <summary>
    /// Host-neutral preparation of an illustration for insertion: it removes the parts the user switched
    /// off, drops placeholder groups, and applies the palette replacements. Nothing here touches COM, so
    /// the same text the insert path writes to its temp file stays unit-testable.
    /// </summary>
    public static class IllustrationInsertionPlanner
    {
        /// <summary>
        /// Returns the SVG text to insert, or null when the input cannot be prepared. The caller keeps the
        /// original file untouched - only the returned text is written to the temp file.
        /// </summary>
        public static string CreateInsertSvg(string svgText, IllustrationInsertOptions options)
        {
            if (string.IsNullOrWhiteSpace(svgText)) return null;

            var effective = options ?? new IllustrationInsertOptions();

            XDocument document;
            try
            {
                document = XDocument.Parse(svgText);
            }
            catch (Exception)
            {
                return null;
            }

            bool removesParts = effective.DropPlaceholders || effective.HiddenPartIds.Count > 0;
            if (removesParts)
            {
                var removable = document
                    .Descendants()
                    .Where(element => element.Name.LocalName == "g")
                    .Where(element =>
                    {
                        var id = element.Attribute("id");
                        if (id == null || string.IsNullOrWhiteSpace(id.Value)) return false;

                        string partId = id.Value.Trim();
                        if (effective.HiddenPartIds.Contains(partId)) return true;

                        return effective.DropPlaceholders && IllustrationSvgPlanner.IsPlaceholderPart(partId);
                    })
                    .ToList();

                foreach (var element in removable)
                {
                    element.Remove();
                }
            }

            string prepared = removesParts ? document.ToString(SaveOptions.DisableFormatting) : svgText;
            return effective.ColorMap.Count == 0
                ? prepared
                : IllustrationSvgPlanner.Recolor(prepared, effective.ColorMap);
        }

        /// <summary>True when the choices change nothing, so callers can reuse the original file as is.</summary>
        public static bool IsNoOp(IllustrationInsertOptions options)
        {
            return options == null
                || (options.ColorMap.Count == 0 && options.HiddenPartIds.Count == 0 && !options.DropPlaceholders);
        }

        /// <summary>
        /// True when the prepared copy is byte-identical to the source, so the caller can insert the original file
        /// directly instead of writing a copy into the temp folder.
        /// </summary>
        public static bool ChangesNothing(string svgText, string prepared)
        {
            return !string.IsNullOrEmpty(svgText)
                && !string.IsNullOrEmpty(prepared)
                && string.Equals(svgText, prepared, StringComparison.Ordinal);
        }

        /// <summary>Part ids of an artwork that are selectable for hiding, placeholders excluded.</summary>
        public static IReadOnlyList<IllustrationPart> SelectableParts(IllustrationSvgPlan plan)
        {
            if (plan == null || plan.Parts == null) return new List<IllustrationPart>();

            return plan.Parts
                .Where(part => !part.IsPlaceholder)
                .Where(part => !string.IsNullOrWhiteSpace(part.Id))
                .ToList();
        }

        /// <summary>Reads the viewBox so the slide keeps the artwork's own aspect ratio.</summary>
        public static bool TryReadViewBox(string svgText, out float width, out float height)
        {
            width = 0f;
            height = 0f;
            if (string.IsNullOrWhiteSpace(svgText)) return false;

            var plan = IllustrationSvgPlanner.Describe(svgText);
            if (string.IsNullOrWhiteSpace(plan.ViewBox)) return false;

            var segments = plan.ViewBox
                .Replace(",", " ")
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (segments.Length < 4) return false;

            float parsedWidth;
            float parsedHeight;
            if (!float.TryParse(segments[2], NumberStyles.Float, CultureInfo.InvariantCulture, out parsedWidth)) return false;
            if (!float.TryParse(segments[3], NumberStyles.Float, CultureInfo.InvariantCulture, out parsedHeight)) return false;
            if (parsedWidth <= 0f || parsedHeight <= 0f) return false;

            width = parsedWidth;
            height = parsedHeight;
            return true;
        }
    }
}
