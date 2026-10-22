using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace oPenEfficiency.Models
{
    /// <summary>One picker category of the Insert Special Characters command (BtnSpecialChars).</summary>
    public sealed class SpecialCharacterCategory
    {
        public SpecialCharacterCategory(string header, string characters)
        {
            Header = header;
            Characters = characters;
        }

        public string Header { get; }

        public string Characters { get; }
    }

    /// <summary>
    /// Host-free catalog of the Insert Special Characters command: the symbol sets first, then the language sets, in
    /// the order the picker renders them. Extracted from the dialog so the catalog is production code that the
    /// acceptance scenarios can assert; the extraction also removed the duplicated '➔' from the arrows set, where the
    /// picker had been showing the same glyph twice.
    /// </summary>
    public static class SpecialCharacterCatalog
    {
        public static readonly IReadOnlyList<SpecialCharacterCategory> Symbols = new[]
        {
            new SpecialCharacterCategory("Mathematical", "¼½¾±≈≠≤≥∞∑∆∏π√∫‰ØΔμΣΩ"),
            new SpecialCharacterCategory("Business", "©®™£¥€¢¤§¶"),
            new SpecialCharacterCategory("Arrows & Process", "➔⇨⇒↔↕↖↗↘↙↺↻⟳➞➠"),
            new SpecialCharacterCategory("Status & Indicators", "⚠⚐🏁🎯🚀⌛⏳✅❌❓❗💡📌🔍🔗"),
            new SpecialCharacterCategory("Harvey Balls", "○◔◑◕●"),
            new SpecialCharacterCategory("Punctuation", "…—–•·«»“”"),
            new SpecialCharacterCategory("Miscellaneous", "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳✓✕✘⚡")
        };

        public static readonly IReadOnlyList<SpecialCharacterCategory> Languages = new[]
        {
            new SpecialCharacterCategory("German", "äöüßÄÖÜ"),
            new SpecialCharacterCategory("French", "àâçéèêëîïôûùÀÂÇÉÈÊËÎÏÔÛÙ"),
            new SpecialCharacterCategory("Spanish", "áéíóúüñÁÉÍÓÚÜÑ¿¡"),
            new SpecialCharacterCategory("Italian", "àèéìòùÀÈÉÌÒÙ"),
            new SpecialCharacterCategory("Portuguese", "áâãàçéêíóôõúÁÂÃÀÇÉÊÍÓÔÕÚ"),
            new SpecialCharacterCategory("Nordic", "åæøÅÆØðþÐÞ"),
            new SpecialCharacterCategory("Polish", "ąćęłńóśźżĄĆĘŁŃÓŚŹŻ"),
            new SpecialCharacterCategory("Turkish", "çğışöüÇĞİŞÖÜ"),
            new SpecialCharacterCategory("Czech/Slovak", "áčďéěíňóřšťúůýžÁČĎÉĚÍŇÓŘŠŤÚŮÝŽ"),
            new SpecialCharacterCategory("Hungarian", "áéíóöőúüűÁÉÍÓÖŐÚÜŰ"),
            new SpecialCharacterCategory("Romanian", "ăâîşţĂÂÎŞŢ")
        };

        /// <summary>The picker order: symbol sets first, then language sets.</summary>
        public static IEnumerable<SpecialCharacterCategory> Categories => Symbols.Concat(Languages);

        /// <summary>The language route of the picker matches category headers case-insensitively.</summary>
        public static bool MatchesHeader(string header, string language)
        {
            return string.Equals(header, language, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>The category a language-route request names, or null when it is unknown.</summary>
        public static SpecialCharacterCategory FindByHeader(string header)
        {
            return Categories.FirstOrDefault(category => MatchesHeader(category.Header, header));
        }

        /// <summary>
        /// The picker buttons of one category: one glyph per text element, so supplementary characters (emoji) are
        /// handed to the dialog whole instead of being split into two lone surrogates.
        /// </summary>
        public static IReadOnlyList<string> EnumerateGlyphs(string characters)
        {
            var glyphs = new List<string>();
            if (string.IsNullOrEmpty(characters)) return glyphs;

            var enumerator = StringInfo.GetTextElementEnumerator(characters);
            while (enumerator.MoveNext()) glyphs.Add(enumerator.GetTextElement());
            return glyphs;
        }
    }
}
