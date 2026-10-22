using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using oPenEfficiency.Models;

namespace oPenEfficiency.Services
{
    /// <summary>Host-neutral description of one illustration SVG: safety state, named parts and palette.</summary>
    public sealed class IllustrationSvgPlan
    {
        public string ViewBox { get; set; }

        public string UnsafeReason { get; set; }

        public List<IllustrationPart> Parts { get; set; } = new List<IllustrationPart>();

        public List<string> Palette { get; set; } = new List<string>();

        public bool IsSafe
        {
            get { return string.IsNullOrEmpty(UnsafeReason); }
        }

        /// <summary>True when at least one named part carries artwork that can be split out.</summary>
        public bool IsDecomposable
        {
            get { return Parts.Any(part => !part.IsPlaceholder); }
        }
    }

    /// <summary>
    /// Reads an illustration SVG the way the picker and the insertion path need it: which named parts
    /// exist, which colour roles they map to, which palette the artwork uses and whether the file is
    /// safe to ship (no scripts, no external references, no embedded raster or active content).
    /// Nothing here touches PowerPoint COM, so every rule stays unit-testable.
    /// </summary>
    public static class IllustrationSvgPlanner
    {
        private static readonly Regex UnsafeContent = new Regex(
            @"(?i)<script|javascript:|<foreignObject|@import|<image|(?:href|src)\s*=\s*['""]https?://",
            RegexOptions.Compiled);

        private static readonly Regex ColorToken = new Regex("#[0-9A-Fa-f]{3,8}\\b", RegexOptions.Compiled);

        private static readonly string[] ShapeElements = { "path", "rect", "circle", "ellipse", "polygon", "polyline", "line" };

        /// <summary>
        /// Basic CSS colour keywords count as their hex equivalents so keyword-coloured artwork can be
        /// recoloured through the same hex map as everything else.
        /// </summary>
        private static readonly Dictionary<string, string> NamedColors = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "black", "#000000" }, { "white", "#FFFFFF" }, { "red", "#FF0000" }, { "lime", "#00FF00" },
            { "blue", "#0000FF" }, { "yellow", "#FFFF00" }, { "cyan", "#00FFFF" }, { "aqua", "#00FFFF" },
            { "magenta", "#FF00FF" }, { "fuchsia", "#FF00FF" }, { "gray", "#808080" }, { "grey", "#808080" },
            { "silver", "#C0C0C0" }, { "maroon", "#800000" }, { "olive", "#808000" }, { "green", "#008000" },
            { "purple", "#800080" }, { "teal", "#008080" }, { "navy", "#000080" }, { "orange", "#FFA500" },
            { "brown", "#A52A2A" }, { "pink", "#FFC0CB" }
        };

        private static readonly Regex NamedColorUse = new Regex(
            @"(?i)\b(?:fill|stroke)\s*[:=]\s*['""]?\s*([a-z]{3,7})\b",
            RegexOptions.Compiled);

        public static IllustrationSvgPlan Describe(string svgText)
        {
            var plan = new IllustrationSvgPlan();

            if (string.IsNullOrWhiteSpace(svgText))
            {
                plan.UnsafeReason = "empty svg";
                return plan;
            }

            var unsafeMatch = UnsafeContent.Match(svgText);
            if (unsafeMatch.Success)
            {
                plan.UnsafeReason = "prohibited content: " + unsafeMatch.Value;
                return plan;
            }

            XDocument document;
            try
            {
                document = XDocument.Parse(svgText);
            }
            catch (Exception ex)
            {
                plan.UnsafeReason = "unparseable svg: " + ex.Message;
                return plan;
            }

            var root = document.Root;
            if (root != null)
            {
                var viewBox = root.Attribute("viewBox");
                plan.ViewBox = viewBox == null ? null : viewBox.Value;
            }

            foreach (var group in document.Descendants().Where(element => element.Name.LocalName == "g"))
            {
                var idAttribute = group.Attribute("id");
                if (idAttribute == null || string.IsNullOrWhiteSpace(idAttribute.Value)) continue;

                string id = idAttribute.Value.Trim();
                plan.Parts.Add(new IllustrationPart
                {
                    Id = id,
                    Name = DisplayName(id),
                    Role = ClassifyRole(id),
                    Depth = id.Count(character => character == '/'),
                    IsPlaceholder = IsPlaceholderPart(id),
                    Colors = CollectElementColors(group)
                });
            }

            plan.Palette = CollectPalette(svgText);
            return plan;
        }

        /// <summary>
        /// Maps a part id to its family inside the mix-and-match layout. The vocabulary follows the
        /// Open Peeps part scheme, which names heads, hairstyles, faces, poses and accessories.
        /// </summary>
        public static string ClassifyRole(string partId)
        {
            if (string.IsNullOrWhiteSpace(partId)) return "Other";

            var segments = partId.Split('/');
            string head = segments[0].Trim();

            if (head.Equals("head", StringComparison.OrdinalIgnoreCase) && segments.Length > 1) return "Hair";
            if (head.Equals("face", StringComparison.OrdinalIgnoreCase)) return "Face";
            if (head.Equals("facial-hair", StringComparison.OrdinalIgnoreCase) || head.Equals("facialhair", StringComparison.OrdinalIgnoreCase)) return "FacialHair";
            if (head.Equals("accessories", StringComparison.OrdinalIgnoreCase) || head.Equals("accessory", StringComparison.OrdinalIgnoreCase)) return "Accessory";
            if (head.Equals("pose", StringComparison.OrdinalIgnoreCase)) return "Pose";
            if (head.Equals("body", StringComparison.OrdinalIgnoreCase)) return "Body";
            if (head.Equals("images", StringComparison.OrdinalIgnoreCase)) return "Container";
            if (head.Equals("Head", StringComparison.OrdinalIgnoreCase)) return "Head";
            if (head.StartsWith("variations", StringComparison.OrdinalIgnoreCase)) return "VariantGroup";
            if (head.Equals("Introduction", StringComparison.OrdinalIgnoreCase)) return "Container";
            if (head.Equals("a-person", StringComparison.OrdinalIgnoreCase)) return "Container";

            return "Other";
        }

        public static bool IsPlaceholderPart(string partId)
        {
            if (string.IsNullOrWhiteSpace(partId)) return true;
            if (partId.IndexOf('*') >= 0) return true;

            string leaf = Leaf(partId);
            return leaf.EndsWith("-None", StringComparison.OrdinalIgnoreCase)
                || leaf.Equals("None", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>Readable name from the leaf segment: "pose/sitting/closed_legs-2" -> "Closed Legs 2".</summary>
        public static string DisplayName(string partId)
        {
            if (string.IsNullOrWhiteSpace(partId)) return string.Empty;

            string leaf = Leaf(partId);
            leaf = leaf.Replace('_', ' ').Replace('*', ' ').Replace('-', ' ').Trim();
            leaf = SplitWordBoundaries(leaf);

            while (leaf.IndexOf("  ", StringComparison.Ordinal) >= 0)
            {
                leaf = leaf.Replace("  ", " ");
            }

            if (leaf.Length == 0) return "None";

            var words = leaf.Split(' ');
            for (int index = 0; index < words.Length; index++)
            {
                if (words[index].Length > 0)
                {
                    words[index] = char.ToUpperInvariant(words[index][0]) + words[index].Substring(1);
                }
            }

            return string.Join(" ", words);
        }

        /// <summary>Distinct colours of the artwork, most frequent first, ties ordered by value.</summary>
        public static List<string> CollectPalette(string svgText)
        {
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrEmpty(svgText)) return new List<string>();

            foreach (Match match in ColorToken.Matches(svgText))
            {
                string token = NormalizeHex(match.Value);
                int current;
                counts.TryGetValue(token, out current);
                counts[token] = current + 1;
            }

            foreach (Match match in NamedColorUse.Matches(svgText))
            {
                string token;
                if (!NamedColors.TryGetValue(match.Groups[1].Value, out token)) continue;

                int current;
                counts.TryGetValue(token, out current);
                counts[token] = current + 1;
            }

            return counts.Keys
                .OrderByDescending(key => counts[key])
                .ThenBy(key => key, StringComparer.Ordinal)
                .ToList();
        }

        /// <summary>
        /// Replaces colour tokens through an explicit hex map. Only colour literals are touched, so
        /// geometry, ids and text content stay untouched and the result is still a plain SVG.
        /// </summary>
        public static string Recolor(string svgText, IDictionary<string, string> colorMap)
        {
            if (string.IsNullOrEmpty(svgText) || colorMap == null || colorMap.Count == 0) return svgText;

            var normalized = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in colorMap)
            {
                if (string.IsNullOrWhiteSpace(pair.Key) || string.IsNullOrWhiteSpace(pair.Value)) continue;
                normalized[NormalizeHex(pair.Key)] = NormalizeHex(pair.Value);
            }

            var recolored = ColorToken.Replace(svgText, match =>
            {
                string replacement;
                return normalized.TryGetValue(NormalizeHex(match.Value), out replacement) ? replacement : match.Value;
            });

            return NamedColorUse.Replace(recolored, match =>
            {
                string token;
                if (!NamedColors.TryGetValue(match.Groups[1].Value, out token)) return match.Value;

                string replacement;
                if (!normalized.TryGetValue(token, out replacement)) return match.Value;

                return match.Value.Replace(match.Groups[1].Value, replacement);
            });
        }

        /// <summary>Colours used by one part's own shapes (per-part recolouring contract).</summary>
        public static List<string> CollectElementColors(XElement element)
        {
            var colors = new List<string>();
            if (element == null) return colors;

            foreach (var shape in element.DescendantsAndSelf())
            {
                if (!ShapeElements.Contains(shape.Name.LocalName, StringComparer.OrdinalIgnoreCase)) continue;

                var fill = shape.Attribute("fill");
                if (fill != null)
                {
                    string token = ResolveColorToken(fill.Value);
                    if (token != null && !colors.Contains(token, StringComparer.OrdinalIgnoreCase)) colors.Add(token);
                }

                var stroke = shape.Attribute("stroke");
                if (stroke != null)
                {
                    string token = ResolveColorToken(stroke.Value);
                    if (token != null && !colors.Contains(token, StringComparer.OrdinalIgnoreCase)) colors.Add(token);
                }

                var style = shape.Attribute("style");
                if (style != null)
                {
                    foreach (Match match in ColorToken.Matches(style.Value))
                    {
                        string token = NormalizeHex(match.Value);
                        if (!colors.Contains(token, StringComparer.OrdinalIgnoreCase)) colors.Add(token);
                    }
                }
            }

            return colors;
        }

        /// <summary>Part roles that a user can expect to see as editable pieces.</summary>
        public static bool IsEditablePartRole(string role)
        {
            return role == "Head" || role == "Hair" || role == "Face" || role == "FacialHair"
                || role == "Accessory" || role == "Pose" || role == "Body" || role == "Other";
        }

        /// <summary>
        /// Part roles worth searching for. Structural containers and variant/pose groups stay out of
        /// the tag list so search results read like the artwork (hair, face, outfit, accessory).
        /// </summary>
        public static bool IsTaggableRole(string role)
        {
            return role == "Head" || role == "Hair" || role == "Face" || role == "FacialHair"
                || role == "Accessory" || role == "Body" || role == "Pose";
        }

        private static string Leaf(string partId)
        {
            return partId.Split('/').Last();
        }

        /// <summary>
        /// Breaks asset-style names into readable words: camel-case humps ("AlienMushroom") and
        /// letter/digit runs ("Blue500"). Lower-case stems such as "peep-4" are unaffected.
        /// </summary>
        private static string SplitWordBoundaries(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;

            var builder = new StringBuilder(value.Length + 8);
            for (int index = 0; index < value.Length; index++)
            {
                char current = value[index];

                if (index > 0 && !char.IsWhiteSpace(value[index - 1]))
                {
                    char previous = value[index - 1];
                    bool camelHump = char.IsLower(previous) && char.IsUpper(current);
                    bool numberEdge = char.IsDigit(current) != char.IsDigit(previous)
                        && (char.IsLetter(previous) || char.IsLetter(current));

                    if (camelHump || numberEdge) builder.Append(' ');
                }

                builder.Append(current);
            }

            return builder.ToString();
        }

        private static string NormalizeHex(string value)
        {
            return (value ?? string.Empty).Trim().ToUpperInvariant();
        }

        /// <summary>Hex value of a colour expression, accepting hex literals and CSS colour keywords.</summary>
        private static string ResolveColorToken(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;

            var match = ColorToken.Match(value);
            if (match.Success) return NormalizeHex(match.Value);

            string named;
            return NamedColors.TryGetValue(value.Trim(), out named) ? named : null;
        }
    }
}
