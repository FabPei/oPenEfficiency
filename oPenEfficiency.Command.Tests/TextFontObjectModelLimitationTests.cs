using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace oPenEfficiency.Command.Tests
{
    /// <summary>
    /// Characterization pin for regression-ledger entry REG-TEXTFONTPROPS-001: the explicit
    /// Apply/Remove paths of FormatStrikethroughFeature and the CharacterSpacingFeature spacing
    /// presets are dead code because the PowerPoint Font dispatch exposes no StrikeThrough/Spacing
    /// member at all (PIA CS0117, COM late binding DISP_E_UNKNOWNNAME, dynamic
    /// RuntimeBinderException - all three measured on 2026-10-03). The documented contract was
    /// corrected to toggle/gallery-only (reviewer disposition (b), 2026-10-03): the dead explicit
    /// paths were removed. These tests pin the object-model fact and that the dead call sites stay
    /// removed, so a future redesign must update the ledger entry and this pin together.
    /// </summary>
    [TestClass]
    public sealed class TextFontObjectModelLimitationTests
    {
        [TestMethod]
        public void PowerPointFontExposesNoStrikeThroughOrSpacingMember()
        {
            // Require the complete test dependency set in output rather than relying on installed Office.
            string output = AppDomain.CurrentDomain.BaseDirectory;
            string officePath = System.IO.Path.Combine(output, "Office.dll");
            string piaPath = System.IO.Path.Combine(output, "Microsoft.Office.Interop.PowerPoint.dll");
            Assert.IsTrue(System.IO.File.Exists(officePath), "The test output must contain Office Core.");
            Assert.IsTrue(System.IO.File.Exists(piaPath), "The test output must contain the PowerPoint PIA.");
            System.Reflection.Assembly.LoadFrom(officePath);
            var pia = System.Reflection.Assembly.LoadFrom(piaPath);
            Type fontType = pia.GetType("Microsoft.Office.Interop.PowerPoint.Font");
            Assert.IsNotNull(fontType, "The PowerPoint PIA Font type could not be resolved.");
            var names = fontType.GetProperties().Select(p => p.Name).ToArray();

            Assert.IsFalse(names.Contains("StrikeThrough"),
                "The PIA now exposes Font.StrikeThrough - the dead strikethrough Apply/Remove paths " +
                "became fixable and REG-TEXTFONTPROPS-001 must be updated.");
            Assert.IsFalse(names.Contains("Spacing"),
                "The PIA now exposes Font.Spacing - the dead spacing presets became fixable and " +
                "REG-TEXTFONTPROPS-001 must be updated.");
        }

        [TestMethod]
        public void TheDeadCallSitesStayRemoved()
        {
            string root = FindRepositoryRoot();
            string source = System.IO.File.ReadAllText(System.IO.Path.Combine(
                root, "oPenEfficiency", "Features", "Text", "TextFormatFeatures.cs"));

            Assert.IsFalse(
                source.IndexOf("StrikeThrough = ", StringComparison.Ordinal) >= 0,
                "A strikethrough write call site is back - REG-TEXTFONTPROPS-001 must be updated.");
            Assert.IsFalse(
                source.IndexOf("Spacing = spacing", StringComparison.Ordinal) >= 0,
                "A spacing-preset call site is back - REG-TEXTFONTPROPS-001 must be updated.");
        }

        private static string FindRepositoryRoot()
        {
            var current = new System.IO.DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (current != null)
            {
                if (System.IO.File.Exists(System.IO.Path.Combine(current.FullName, "oPenEfficiency.slnx"))) return current.FullName;
                current = current.Parent;
            }

            throw new System.IO.DirectoryNotFoundException("Could not locate the repository root.");
        }
    }
}
