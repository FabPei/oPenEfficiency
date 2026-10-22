using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public class IllustrationImportPolicyTests
    {
        private const string CleanSvg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 100 100\">" +
            "<g id=\"Body\"><path fill=\"#263238\" d=\"M0 0h10v10z\"/></g>" +
            "</svg>";

        [TestMethod]
        public void CleanArtworkIsAcceptedWithItsFacts()
        {
            var decision = IllustrationImportPolicy.Evaluate("my-illustration.svg", CleanSvg);

            Assert.IsTrue(decision.Accepted, decision.Reason);
            Assert.AreEqual("My Illustration", decision.Name);
            Assert.AreEqual("0 0 100 100", decision.ViewBox);
            Assert.AreEqual(1, decision.PaletteSize);
            Assert.AreEqual(1, decision.PartCount);
            Assert.IsTrue(decision.Decomposable);
            Assert.IsNull(decision.Note);
        }

        [TestMethod]
        public void MissingViewBoxIsAcceptedWithARemark()
        {
            var decision = IllustrationImportPolicy.Evaluate(
                "no-viewbox.svg",
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"80\" height=\"60\"><path fill=\"#FF0000\" d=\"M0 0h1v1z\"/></svg>");

            Assert.IsTrue(decision.Accepted, decision.Reason);
            Assert.IsFalse(string.IsNullOrWhiteSpace(decision.Note), "a missing viewBox must be reported");
        }

        [TestMethod]
        public void NonSvgFilesAreRefused()
        {
            var decision = IllustrationImportPolicy.Evaluate("picture.png", CleanSvg);

            Assert.IsFalse(decision.Accepted);
            StringAssert.Contains(decision.Reason, ".svg");
        }

        [TestMethod]
        public void OversizedArtworkIsRefused()
        {
            string huge = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 1 1\"><path fill=\"#000000\" d=\"" +
                new string('M', IllustrationImportPolicy.MaxCharacters) + "\"/></svg>";

            var decision = IllustrationImportPolicy.Evaluate("huge.svg", huge);

            Assert.IsFalse(decision.Accepted);
            StringAssert.Contains(decision.Reason, "512 KB");
        }

        [TestMethod]
        public void ActiveContentIsRefused()
        {
            var decision = IllustrationImportPolicy.Evaluate(
                "sketchy.svg",
                "<svg xmlns=\"http://www.w3.org/2000/svg\"><script>alert(1)</script><path fill=\"#000000\" d=\"M0 0h1v1z\"/></svg>");

            Assert.IsFalse(decision.Accepted);
            StringAssert.Contains(decision.Reason, "does not insert");
        }

        [TestMethod]
        public void EmbeddedRasterIsRefused()
        {
            var decision = IllustrationImportPolicy.Evaluate(
                "embedded.svg",
                "<svg xmlns=\"http://www.w3.org/2000/svg\"><image href=\"data:image/png;base64,AAA\"/></svg>");

            Assert.IsFalse(decision.Accepted);
            StringAssert.Contains(decision.Reason, "does not insert");
        }

        [TestMethod]
        public void ArtworkWithoutColoursIsRefused()
        {
            var decision = IllustrationImportPolicy.Evaluate(
                "outline.svg",
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10 10\"><path fill=\"none\" stroke=\"none\" d=\"M0 0h1v1z\"/></svg>");

            Assert.IsFalse(decision.Accepted);
            StringAssert.Contains(decision.Reason, "colours");
        }

        [TestMethod]
        public void EmptyAndNonSvgPayloadsAreRefused()
        {
            Assert.IsFalse(IllustrationImportPolicy.Evaluate("blank.svg", "   ").Accepted);
            Assert.IsFalse(IllustrationImportPolicy.Evaluate("text.svg", "just some text, no markup").Accepted);
            Assert.IsFalse(IllustrationImportPolicy.Evaluate(null, CleanSvg).Accepted);
        }

        [TestMethod]
        public void ImportTrailRecordsSourceAndLicense()
        {
            var source = IllustrationSourceCatalog.Find("storyset");
            string trail = IllustrationImportPolicy.DescribeTrail(source, "team.svg", new DateTime(2026, 9, 25, 0, 0, 0, DateTimeKind.Utc));

            StringAssert.Contains(trail, "team.svg");
            StringAssert.Contains(trail, "2026-09-25");
            StringAssert.Contains(trail, "Storyset");
            StringAssert.Contains(trail, "credit required");
        }
    }
}
