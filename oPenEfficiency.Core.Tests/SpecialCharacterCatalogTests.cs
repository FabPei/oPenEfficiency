using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA9 wave 3 (text and formatting): the host-free catalog behind Insert Special Characters (BtnSpecialChars).
    /// The dialog renders and inserts from <c>SpecialCharacterCatalog</c>; the COM leg (inserting into the text
    /// range) is office-automated and QA-P0-blocked.
    /// </summary>
    [TestClass]
    public class SpecialCharacterCatalogTests
    {
        /// <summary>Text elements of a set, i.e. what a user perceives as one character.</summary>
        private static List<string> TextElements(string characters)
        {
            var elements = new List<string>();
            var enumerator = StringInfo.GetTextElementEnumerator(characters);
            while (enumerator.MoveNext()) elements.Add(enumerator.GetTextElement());
            return elements;
        }

        [TestMethod]
        public void EveryCategoryIsNonEmptyAndFreeOfDuplicateCharacters()
        {
            Assert.IsTrue(SpecialCharacterCatalog.Symbols.Count > 0, "the symbol sets exist");
            Assert.IsTrue(SpecialCharacterCatalog.Languages.Count > 0, "the language sets exist");

            foreach (var category in SpecialCharacterCatalog.Categories)
            {
                Assert.IsFalse(string.IsNullOrWhiteSpace(category.Header), "every category carries a header");
                Assert.IsTrue(category.Characters.Length > 0, category.Header + " offers at least one character");

                var seen = new HashSet<string>();
                foreach (string element in TextElements(category.Characters))
                {
                    Assert.IsTrue(seen.Add(element),
                        category.Header + " lists '" + element + "' twice: the picker would show two identical buttons");
                }
            }
        }

        [TestMethod]
        public void CategoryLookupIsCaseInsensitiveAndUnknownNamesResolveToNothing()
        {
            Assert.IsTrue(SpecialCharacterCatalog.MatchesHeader("German", "german"));
            Assert.AreEqual("German", SpecialCharacterCatalog.FindByHeader("GERMAN").Header,
                "the language route matches headers case-insensitively");
            Assert.AreEqual("Czech/Slovak", SpecialCharacterCatalog.FindByHeader("czech/slovak").Header);
            Assert.IsNull(SpecialCharacterCatalog.FindByHeader("Klingon"), "an unknown language resolves to nothing");
            Assert.IsNull(SpecialCharacterCatalog.FindByHeader(string.Empty));
        }

        [TestMethod]
        public void DocumentedCharacterSetsArePresent()
        {
            Assert.AreEqual("äöüßÄÖÜ", SpecialCharacterCatalog.FindByHeader("German").Characters,
                "the German set is exactly the documented umlauts and sharp s");
            Assert.AreEqual("○◔◑◕●", SpecialCharacterCatalog.FindByHeader("Harvey Balls").Characters,
                "the picker offers the five-step Harvey ball progression");
            Assert.IsTrue(SpecialCharacterCatalog.FindByHeader("Spanish").Characters.Contains('¿'));
            Assert.IsTrue(SpecialCharacterCatalog.FindByHeader("Mathematical").Characters.Contains('∞'));
            Assert.IsTrue(SpecialCharacterCatalog.FindByHeader("Polish").Characters.Contains('ł'));

            string arrows = SpecialCharacterCatalog.FindByHeader("Arrows & Process").Characters;
            Assert.AreEqual(1, arrows.Count(character => character == '➔'),
                "the arrows set lists the arrow once - before the extraction the picker showed it twice");
        }

        [TestMethod]
        public void GlyphEnumerationKeepsSupplementaryCharactersWhole()
        {
            var status = SpecialCharacterCatalog.FindByHeader("Status & Indicators");
            var glyphs = SpecialCharacterCatalog.EnumerateGlyphs(status.Characters);
            Assert.AreEqual(15, glyphs.Count, "one picker button per text element, not per UTF-16 unit");
            Assert.IsTrue(glyphs.Contains("🏁"), "the flag emoji stays a single button");

            foreach (var category in SpecialCharacterCatalog.Categories)
            {
                foreach (string glyph in SpecialCharacterCatalog.EnumerateGlyphs(category.Characters))
                {
                    Assert.IsFalse(glyph.Length == 1 && char.IsSurrogate(glyph[0]),
                        category.Header + " would split a surrogate pair into a broken button: '" + glyph + "'");
                }
            }

            Assert.AreEqual(0, SpecialCharacterCatalog.EnumerateGlyphs(string.Empty).Count);
        }
    }
}
