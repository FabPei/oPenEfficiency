using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 3 coverage for the Quick Context Menu configuration rule
    /// (QA9-BtnQuickContextMenu-001): which configured feature ids survive into the popup, in which
    /// order, and how the item cap behaves. Production implementation used by
    /// QuickContextMenuService.
    /// </summary>
    [TestClass]
    public sealed class QuickContextMenuSettingsTests
    {
        [TestMethod]
        public void NormalizeKeepsOrderDropsUnknownAndDuplicates()
        {
            var known = new HashSet<string>(new[] { "BtnA", "BtnB", "BtnC" });
            var result = QuickContextMenuSettings.Normalize(
                new[] { "BtnB", "Unknown", "BtnA", "BtnB", "BtnC" }, known);

            CollectionAssert.AreEqual(new List<string> { "BtnB", "BtnA", "BtnC" }, result,
                "Normalize must keep the configured order, drop unknown ids and keep only the first duplicate.");
        }

        [TestMethod]
        public void NormalizeIsCaseSensitiveAgainstTheOrdinalKnownSet()
        {
            var known = new HashSet<string>(new[] { "BtnA" }, System.StringComparer.Ordinal);
            var result = QuickContextMenuSettings.Normalize(new[] { "btna", "BtnA" }, known);

            CollectionAssert.AreEqual(new List<string> { "BtnA" }, result,
                "The popup builds its known-id set with StringComparer.Ordinal; a differently cased id must be treated as unknown.");
        }

        [TestMethod]
        public void NormalizeTreatsNullKnownIdsAsAllowEverything()
        {
            var result = QuickContextMenuSettings.Normalize(new[] { "Anything", "Goes" }, null);

            CollectionAssert.AreEqual(new List<string> { "Anything", "Goes" }, result,
                "A null known-id set disables the unknown filter (historical behavior).");
        }

        [TestMethod]
        public void NormalizeSkipsBlankEntriesWithoutConsumingTheCap()
        {
            var known = new HashSet<string>(new[] { "BtnA", "BtnB" });
            var result = QuickContextMenuSettings.Normalize(new[] { null, "  ", "BtnA", "BtnB" }, known);

            CollectionAssert.AreEqual(new List<string> { "BtnA", "BtnB" }, result,
                "Null or whitespace ids must be skipped and must not eat into the item cap.");
        }

        [TestMethod]
        public void NormalizeCapsAtTwelveItemsAndKeepsOrder()
        {
            var ids = new List<string>();
            for (int i = 0; i < 20; i++) ids.Add("Btn" + i.ToString("00"));

            var result = QuickContextMenuSettings.Normalize(ids, null);

            Assert.AreEqual(QuickContextMenuSettings.MaximumItems, result.Count,
                "The popup must never show more than MaximumItems entries.");
            CollectionAssert.AreEqual(ids.GetRange(0, QuickContextMenuSettings.MaximumItems), result,
                "The cap must keep the first configured items in order.");
        }

        [TestMethod]
        public void DefaultsAreDistinctAndWithinTheCap()
        {
            var defaults = QuickContextMenuSettings.DefaultFeatureIds;

            Assert.IsTrue(defaults.Length >= QuickContextMenuSettings.MinimumItems,
                "The default popup must offer at least MinimumItems entries.");
            Assert.IsTrue(defaults.Length <= QuickContextMenuSettings.MaximumItems,
                "The default popup must stay within MaximumItems entries.");
            Assert.AreEqual(defaults.Length, new HashSet<string>(defaults).Count,
                "The default popup ids must be distinct.");

            var normalized = QuickContextMenuSettings.Normalize(defaults, new HashSet<string>(defaults));
            CollectionAssert.AreEqual(new List<string>(defaults), normalized,
                "The defaults must survive normalization unchanged (no unknown ids, no duplicates).");
        }
    }
}
