using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public class IllustrationInsertionPlannerTests
    {
        private const string PartedSvg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 240 360\">" +
            "<g id=\"body/Blazer\"><path fill=\"#263238\" d=\"M0 0h1v1z\"/></g>" +
            "<g id=\"head/Bun\"><path fill=\"#E6E6E6\" d=\"M0 0h1v1z\"/></g>" +
            "<g id=\"facial-hair/*-None\"><path fill=\"#000000\" d=\"M0 0h1v1z\"/></g>" +
            "</svg>";

        [TestMethod]
        public void PlaceholderGroupsAreDroppedByDefault()
        {
            var result = IllustrationInsertionPlanner.CreateInsertSvg(PartedSvg, new IllustrationInsertOptions());

            Assert.IsFalse(result.Contains("facial-hair"), result);
            Assert.IsTrue(result.Contains("body/Blazer"), result);
            Assert.IsTrue(result.Contains("head/Bun"), result);
        }

        [TestMethod]
        public void PlaceholderRetentionIsPossible()
        {
            var options = new IllustrationInsertOptions { DropPlaceholders = false };

            var result = IllustrationInsertionPlanner.CreateInsertSvg(PartedSvg, options);

            Assert.IsTrue(result.Contains("facial-hair/*-None"), result);
        }

        [TestMethod]
        public void HiddenPartsAreRemovedButSiblingsStayIntact()
        {
            var options = new IllustrationInsertOptions();
            options.HiddenPartIds.Add("head/Bun");

            var result = IllustrationInsertionPlanner.CreateInsertSvg(PartedSvg, options);

            Assert.IsFalse(result.Contains("head/Bun"), result);
            Assert.IsTrue(result.Contains("body/Blazer"), result);
            Assert.IsTrue(result.Contains("<svg"), result);
        }

        [TestMethod]
        public void UnknownPartIdsChangeNothing()
        {
            var options = new IllustrationInsertOptions();
            options.HiddenPartIds.Add("does/not-exist");

            var result = IllustrationInsertionPlanner.CreateInsertSvg(PartedSvg, options);

            Assert.IsTrue(result.Contains("body/Blazer"), result);
            Assert.IsTrue(result.Contains("head/Bun"), result);
            Assert.IsFalse(result.Contains("facial-hair"), result);
        }

        [TestMethod]
        public void ColorMapIsAppliedAfterThePartsAreRemoved()
        {
            var options = new IllustrationInsertOptions();
            options.ColorMap["#263238"] = "#BA68C8";

            var result = IllustrationInsertionPlanner.CreateInsertSvg(PartedSvg, options);

            Assert.IsTrue(result.Contains("#BA68C8"), result);
            Assert.IsFalse(result.Contains("#263238"), result);
            Assert.IsTrue(result.Contains("#E6E6E6"), result);
        }

        [TestMethod]
        public void OriginalTextIsNeverMutated()
        {
            var options = new IllustrationInsertOptions();
            options.HiddenPartIds.Add("body/Blazer");
            options.ColorMap["#263238"] = "#123456";

            IllustrationInsertionPlanner.CreateInsertSvg(PartedSvg, options);

            Assert.IsTrue(PartedSvg.Contains("body/Blazer"), "the caller's text must stay untouched");
            Assert.IsTrue(PartedSvg.Contains("#263238"), "the caller's text must stay untouched");
        }

        [TestMethod]
        public void BrokenInputYieldsNullInsteadOfThrowing()
        {
            Assert.IsNull(IllustrationInsertionPlanner.CreateInsertSvg("<svg><g id=\"unterminated\">", new IllustrationInsertOptions()));
            Assert.IsNull(IllustrationInsertionPlanner.CreateInsertSvg(null, new IllustrationInsertOptions()));
        }

        [TestMethod]
        public void NoOpRoutesAreReportedExactly()
        {
            Assert.IsTrue(IllustrationInsertionPlanner.IsNoOp(null));
            Assert.IsTrue(IllustrationInsertionPlanner.IsNoOp(new IllustrationInsertOptions { DropPlaceholders = false }));

            var withHiddenPart = new IllustrationInsertOptions { DropPlaceholders = false };
            withHiddenPart.HiddenPartIds.Add("head/Bun");
            Assert.IsFalse(IllustrationInsertionPlanner.IsNoOp(withHiddenPart));

            var withColor = new IllustrationInsertOptions { DropPlaceholders = false };
            withColor.ColorMap["#000000"] = "#FFFFFF";
            Assert.IsFalse(IllustrationInsertionPlanner.IsNoOp(withColor));

            Assert.IsFalse(IllustrationInsertionPlanner.IsNoOp(new IllustrationInsertOptions()));

            // The second route: a copy that came out identical to the source needs no temp file either.
            Assert.IsTrue(IllustrationInsertionPlanner.ChangesNothing(PartedSvg, PartedSvg));

            var droppingOptions = new IllustrationInsertOptions();
            droppingOptions.HiddenPartIds.Add("head/Bun");
            string dropping = IllustrationInsertionPlanner.CreateInsertSvg(PartedSvg, droppingOptions);
            Assert.IsFalse(IllustrationInsertionPlanner.ChangesNothing(PartedSvg, dropping));

            Assert.IsFalse(IllustrationInsertionPlanner.ChangesNothing(null, null));
            Assert.IsFalse(IllustrationInsertionPlanner.ChangesNothing(string.Empty, string.Empty));
            Assert.IsFalse(IllustrationInsertionPlanner.ChangesNothing(PartedSvg, null));
        }

        [TestMethod]
        public void TheInsertedCopyWithholdsPlaceholdersAndHiddenPartsAndLeavesTheSourceUntouched()
        {
            var options = new IllustrationInsertOptions();
            options.HiddenPartIds.Add("head/Bun");

            string prepared = IllustrationInsertionPlanner.CreateInsertSvg(PartedSvg, options);

            // The placeholder variant and the switched-off part are withheld from the inserted copy ...
            Assert.IsFalse(prepared.Contains("facial-hair"), prepared);
            Assert.IsFalse(prepared.Contains("head/Bun"), prepared);
            Assert.IsTrue(prepared.Contains("body/Blazer"), prepared);

            // ... while the file on disk keeps every part it had.
            Assert.IsTrue(PartedSvg.Contains("facial-hair/*-None"), "the source keeps its placeholder group");
            Assert.IsTrue(PartedSvg.Contains("head/Bun"), "the source keeps the part the user switched off");
        }

        [TestMethod]
        public void SelectablePartsSkipPlaceholders()
        {
            var plan = IllustrationSvgPlanner.Describe(PartedSvg);
            var parts = IllustrationInsertionPlanner.SelectableParts(plan);

            Assert.AreEqual(2, parts.Count);
            CollectionAssert.AreEquivalent(
                new[] { "body/Blazer", "head/Bun" },
                parts.Select(part => part.Id).ToArray());
        }

        [TestMethod]
        public void ViewBoxDrivesTheAspectRatio()
        {
            float width;
            float height;

            Assert.IsTrue(IllustrationInsertionPlanner.TryReadViewBox(PartedSvg, out width, out height));
            Assert.AreEqual(240f, width);
            Assert.AreEqual(360f, height);

            Assert.IsFalse(IllustrationInsertionPlanner.TryReadViewBox("<svg xmlns=\"x\"><path d=\"M0 0\"/></svg>", out width, out height));
        }
    }
}
