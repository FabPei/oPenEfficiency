using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 3 coverage for the Template Library rules (QA9-BtnTemplateManager-002): the
    /// supported template extensions and the three-source merge (priority order, first-wins
    /// de-duplication, blank filtering, case-insensitive blacklist), used by TemplateService.
    /// </summary>
    [TestClass]
    public sealed class TemplateLibraryPolicyTests
    {
        [TestMethod]
        public void SupportedExtensionsMatchTheLibraryContract()
        {
            Assert.IsTrue(TemplateLibraryPolicy.IsSupportedPath(@"C:\Templates\a.potx"));
            Assert.IsTrue(TemplateLibraryPolicy.IsSupportedPath(@"C:\Templates\a.POT"));
            Assert.IsTrue(TemplateLibraryPolicy.IsSupportedPath(@"C:\Templates\a.pptm"));
            Assert.IsTrue(TemplateLibraryPolicy.IsSupportedPath(@"C:\Templates\a.PPTX"));
            Assert.IsFalse(TemplateLibraryPolicy.IsSupportedPath(@"C:\Templates\a.docx"), "Word files are not templates.");
            Assert.IsFalse(TemplateLibraryPolicy.IsSupportedPath(@"C:\Templates\a.pptx.bak"), "Only the final extension counts.");
            Assert.IsFalse(TemplateLibraryPolicy.IsSupportedPath(null));
            Assert.IsFalse(TemplateLibraryPolicy.IsSupportedPath("   "));
        }

        [TestMethod]
        public void MergeKeepsPriorityOrderAndFirstWinsPerPath()
        {
            var merged = TemplateLibraryPolicy.MergePaths(
                registered: new[] { @"C:\T\a.potx" },
                watched: new[] { @"C:\T\a.potx", @"C:\W\b.pptx" },
                defaults: new[] { @"C:\D\c.potx" },
                blacklist: null);

            CollectionAssert.AreEqual(
                new List<string> { @"C:\T\a.potx", @"C:\W\b.pptx", @"C:\D\c.potx" },
                merged,
                "The registered entry wins over the watched duplicate and the order stays registered, watched, defaults.");
        }

        [TestMethod]
        public void DeDuplicationAndBlacklistAreCaseInsensitive()
        {
            var merged = TemplateLibraryPolicy.MergePaths(
                registered: new[] { @"C:\T\A.POTX" },
                watched: new[] { @"C:\T\a.potx", @"C:\T\SECRET.pptx" },
                defaults: null,
                blacklist: new[] { @"c:\t\secret.PPTX" });

            CollectionAssert.AreEqual(new List<string> { @"C:\T\A.POTX" }, merged,
                "Case-insensitive de-duplication and blacklisting must both apply.");
        }

        [TestMethod]
        public void BlankEntriesAreDroppedAndEmptySourcesAreAllowed()
        {
            var merged = TemplateLibraryPolicy.MergePaths(
                registered: new[] { null, "  ", @"C:\T\a.potx" },
                watched: null,
                defaults: new string[0],
                blacklist: new List<string>());

            CollectionAssert.AreEqual(new List<string> { @"C:\T\a.potx" }, merged);
            Assert.AreEqual(0, TemplateLibraryPolicy.MergePaths(null, null, null, null).Count,
                "All-empty inputs must merge to an empty library.");
        }
    }
}
