using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Core.Services;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 wave 3 coverage for the Export Wizard rules (QA9-BtnExportWizard-002): the filename
    /// sanitization contract (replace, not remove) and the overwrite-path comparison, used by
    /// ExportWizard.
    /// </summary>
    [TestClass]
    public sealed class ExportWizardPolicyTests
    {
        [TestMethod]
        public void SanitizationReplacesWhitespaceAndInvalidCharacters()
        {
            Assert.AreEqual("Q3_report_final", ExportWizardPolicy.SanitizeFilenameComponent("Q3 report final"));
            Assert.AreEqual("a_b_c", ExportWizardPolicy.SanitizeFilenameComponent("a:b*c"),
                "Invalid filename characters become underscores (replace, not remove).");
            Assert.AreEqual("__", ExportWizardPolicy.SanitizeFilenameComponent(" \t"),
                "Whitespace becomes underscores.");
            Assert.AreEqual("", ExportWizardPolicy.SanitizeFilenameComponent(null),
                "Null becomes an empty component (unlike the Asset Library 'unnamed' fallback).");
        }

        [TestMethod]
        public void PathComparisonIsCaseInsensitiveAndRejectsBlankInputs()
        {
            string a = Path.Combine(Path.GetTempPath(), "oPE-export-a.pdf");
            Assert.IsTrue(ExportWizardPolicy.PathsEqual(a, a.ToUpperInvariant()),
                "Target paths compare case-insensitively on full paths.");
            Assert.IsTrue(ExportWizardPolicy.PathsEqual(a, Path.Combine(Path.GetTempPath(), @".\oPE-export-a.pdf")),
                "Relative segments are normalized before comparison.");
            Assert.IsFalse(ExportWizardPolicy.PathsEqual(a, null));
            Assert.IsFalse(ExportWizardPolicy.PathsEqual("  ", a));
        }
    }
}
