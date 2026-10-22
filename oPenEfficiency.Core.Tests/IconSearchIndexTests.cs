using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 3 coverage for the Icon Browser kernel (QA9-BtnIconBrowser-002): the search
    /// filtering and result ordering the browser shows, the manifest cache invalidation, and the
    /// tolerant manifest parsing. Production implementation in IconCatalogComponents.
    /// </summary>
    [TestClass]
    public sealed class IconSearchIndexTests
    {
        private static IconCatalogItem Item(string name, string collection = "Core", string theme = "Light",
            string style = "Outline", bool favorite = false, string tag = null)
        {
            return new IconCatalogItem
            {
                Name = name,
                CollectionId = collection,
                Theme = theme,
                Style = style,
                IsFavorite = favorite,
                // SearchText is computed from these fields; tags extend what a needle can match.
                Tags = tag == null ? new System.Collections.Generic.List<string>() : new System.Collections.Generic.List<string> { tag }
            };
        }

        [TestMethod]
        public void EmptyQueryReturnsEverythingOrderedByNameThenStyle()
        {
            var index = new IconSearchIndex(new[]
            {
                Item("beta", style: "Solid"),
                Item("alpha", style: "Solid"),
                Item("alpha", style: "Outline"),
            });

            var results = index.Search();

            Assert.AreEqual(3, results.Count, "An empty query returns the whole catalog.");
            Assert.AreEqual("alpha", results[0].Name);
            Assert.AreEqual("Outline", results[0].Style, "Equal names order by style.");
            Assert.AreEqual("alpha", results[1].Name);
            Assert.AreEqual("beta", results[2].Name, "Names order case-insensitively by natural order.");
        }

        [TestMethod]
        public void FacetFiltersComposeCaseInsensitively()
        {
            var index = new IconSearchIndex(new[]
            {
                Item("a", "Core", "Light", "Outline", true),
                Item("b", "Core", "Dark", "Outline"),
                Item("c", "Flags", "Light", "Outline"),
                Item("d", "Core", "Light", "Filled"),
            });

            Assert.AreEqual(1, index.Search(collectionId: "core", theme: "LIGHT", style: "outline", favoritesOnly: true).Count,
                "All three facets and the favorite flag compose, case-insensitively.");
            Assert.AreEqual(3, index.Search(collectionId: "CORE").Count, "Collection filter alone (items a, b, d).");
            Assert.AreEqual(3, index.Search(theme: "light").Count, "Theme filter alone (items a, c, d).");
            Assert.AreEqual(3, index.Search(style: "outline", favoritesOnly: false).Count, "Style filter alone (items a, b, c).");
        }

        [TestMethod]
        public void TextSearchIsTrimmedAndCaseInsensitive()
        {
            var index = new IconSearchIndex(new[]
            {
                Item("Check Circle", tag: "tick mark"),
                Item("Arrow Right", tag: "chevron"),
            });

            var results = index.Search(search: "  CHECK  ");
            Assert.AreEqual(1, results.Count, "The needle is trimmed and lower-cased before matching.");
            Assert.AreEqual("Check Circle", results[0].Name);
        }

        [TestMethod]
        public void NullInputsAreHandledAndTheCacheIsStampKeyed()
        {
            var index = new IconSearchIndex(null);
            Assert.AreEqual(0, index.Search().Count, "A null catalog behaves like an empty one.");

            var cache = new IconManifestCache();
            int loads = 0;
            Func<string, IconCollectionManifest> loader = _ => { loads++; return new IconCollectionManifest(); };

            var first = cache.Get(@"C:\m.json", new DateTime(2026, 1, 1), loader);
            var same = cache.Get(@"C:\m.json", new DateTime(2026, 1, 1), loader);
            Assert.AreSame(first, same, "Same path and stamp reuse the cached manifest.");
            Assert.AreEqual(1, loads);

            var newer = cache.Get(@"C:\m.json", new DateTime(2026, 1, 2), loader);
            Assert.AreNotSame(first, newer, "A newer stamp reloads.");
            Assert.AreEqual(2, loads);

            cache.Invalidate();
            var afterInvalidate = cache.Get(@"C:\m.json", new DateTime(2026, 1, 2), loader);
            Assert.AreNotSame(newer, afterInvalidate, "Invalidate forces a reload.");
            Assert.AreEqual(3, loads);
        }

        [TestMethod]
        public void ParsingIsTolerantAndCloningIsIndependent()
        {
            var parser = new IconManifestParser();
            Assert.IsNotNull(parser.Parse(null), "A null manifest string yields an empty manifest.");
            Assert.IsNotNull(parser.Parse("   "), "A blank manifest string yields an empty manifest.");

            var source = new IconCollectionManifest { Version = "2.0" };
            var clone = parser.Clone(source);
            Assert.AreEqual("2.0", clone.Version, "Clone keeps the values.");
            clone.Version = "9.9";
            Assert.AreEqual("2.0", source.Version, "The clone is deep, not a shared reference.");
        }
    }
}
