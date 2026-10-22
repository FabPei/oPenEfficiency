using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public class IllustrationSvgPlannerTests
    {
        private const string PeepLikeSvg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 346 510\">" +
            "<g id=\"Introduction\"><g id=\"variations-sitting\"><g id=\"a-person/sitting\">" +
            "<g id=\"pose/sitting/closed_legs-2\"><path fill=\"#263238\" d=\"M0 0h1v1z\"/></g></g></g></g>" +
            "<g id=\"Head\"><path fill=\"#E6E6E6\" d=\"M0 0h1v1z\"/></g>" +
            "<g id=\"head/Afro\"><path style=\"fill:#263238\" d=\"M0 0h1v1z\"/></g>" +
            "<g id=\"face/Calm\"><path fill=\"#FFFFFF\" d=\"M0 0h1v1z\"/><path fill=\"#263238\" d=\"M0 0h1v1z\"/></g>" +
            "<g id=\"facial-hair/*-None\"><path fill=\"#000000\" d=\"M0 0h1v1z\"/></g>" +
            "</svg>";

        [TestMethod]
        public void DescribesNamedPartsWithRolesDepthAndPlaceholders()
        {
            var plan = IllustrationSvgPlanner.Describe(PeepLikeSvg);

            Assert.IsTrue(plan.IsSafe, plan.UnsafeReason);
            Assert.AreEqual("0 0 346 510", plan.ViewBox);
            Assert.AreEqual(8, plan.Parts.Count);

            var byId = plan.Parts.ToDictionary(part => part.Id, StringComparer.Ordinal);
            Assert.AreEqual("Head", byId["Head"].Role);
            Assert.AreEqual("Hair", byId["head/Afro"].Role);
            Assert.AreEqual("Face", byId["face/Calm"].Role);
            Assert.AreEqual("FacialHair", byId["facial-hair/*-None"].Role);
            Assert.AreEqual("Pose", byId["pose/sitting/closed_legs-2"].Role);
            Assert.AreEqual("Container", byId["a-person/sitting"].Role);
            Assert.AreEqual("VariantGroup", byId["variations-sitting"].Role);

            Assert.AreEqual(2, byId["pose/sitting/closed_legs-2"].Depth);
            Assert.AreEqual(0, byId["Head"].Depth);

            Assert.IsTrue(byId["facial-hair/*-None"].IsPlaceholder);
            Assert.IsFalse(byId["face/Calm"].IsPlaceholder);
            Assert.IsTrue(plan.IsDecomposable);
        }

        [TestMethod]
        public void PartNamesReadAsUserFacingLabels()
        {
            Assert.AreEqual("Closed Legs 2", IllustrationSvgPlanner.DisplayName("pose/sitting/closed_legs-2"));
            Assert.AreEqual("Afro", IllustrationSvgPlanner.DisplayName("head/Afro"));
            Assert.AreEqual("None", IllustrationSvgPlanner.DisplayName("facial-hair/*-None"));
            Assert.AreEqual("Peep 4", IllustrationSvgPlanner.DisplayName("peep-4"));
            Assert.AreEqual("Alien Mushroom Organic V M Blue 500", IllustrationSvgPlanner.DisplayName("AlienMushroom-Organic-V-M-Blue500"));
            Assert.AreEqual("Apple 1", IllustrationSvgPlanner.DisplayName("Apple 1"));
        }

        [TestMethod]
        public void StrokeOnlyArtworkStillReportsItsColors()
        {
            const string strokeOnlySvg =
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10 10\">" +
                "<g id=\"Cable\"><path fill=\"none\" stroke=\"#57534E\" d=\"M0 0h1v1z\"/></g>" +
                "</svg>";

            var plan = IllustrationSvgPlanner.Describe(strokeOnlySvg);
            var cable = plan.Parts.Single(part => part.Id == "Cable");

            CollectionAssert.AreEqual(new[] { "#57534E" }, cable.Colors.ToArray());
            CollectionAssert.AreEqual(new[] { "#57534E" }, plan.Palette.ToArray());
        }

        [TestMethod]
        public void PartColorsFollowThePartOwnShapesOnly()
        {
            var plan = IllustrationSvgPlanner.Describe(PeepLikeSvg);
            var face = plan.Parts.Single(part => part.Id == "face/Calm");
            var pose = plan.Parts.Single(part => part.Id == "pose/sitting/closed_legs-2");

            CollectionAssert.AreEqual(new[] { "#FFFFFF", "#263238" }, face.Colors.ToArray());
            CollectionAssert.AreEqual(new[] { "#263238" }, pose.Colors.ToArray());
        }

        [TestMethod]
        public void PaletteIsMostFrequentFirstAndDeterministic()
        {
            var palette = IllustrationSvgPlanner.CollectPalette(PeepLikeSvg);

            Assert.AreEqual("#263238", palette[0]);
            CollectionAssert.AreEqual(new[] { "#263238", "#000000", "#E6E6E6", "#FFFFFF" }, palette.ToArray());
        }

        [TestMethod]
        public void RecolorOnlyTouchesColorTokens()
        {
            var recolored = IllustrationSvgPlanner.Recolor(PeepLikeSvg, new Dictionary<string, string> { { "#263238", "#BA68C8" } });

            Assert.IsTrue(recolored.Contains("fill=\"#BA68C8\""));
            Assert.IsTrue(recolored.Contains("style=\"fill:#BA68C8\""));
            Assert.IsFalse(recolored.Contains("#263238"));
            Assert.IsTrue(recolored.Contains("id=\"head/Afro\""));
            Assert.IsTrue(recolored.Contains("fill=\"#FFFFFF\""));
        }

        [TestMethod]
        public void ActiveOrExternalContentIsRejected()
        {
            foreach (var payload in new[]
            {
                "<script>alert(1)</script>",
                "<image href=\"https://example.com/a.png\"/>",
                "<a xlink:href=\"http://example.com\">x</a>",
                "<style>@import url(x);</style>"
            })
            {
                var plan = IllustrationSvgPlanner.Describe("<svg xmlns=\"http://www.w3.org/2000/svg\">" + payload + "</svg>");
                Assert.IsFalse(plan.IsSafe, "should be rejected: " + payload);
                StringAssert.StartsWith(plan.UnsafeReason, "prohibited content");
            }
        }

        [TestMethod]
        public void NamedColorsCountAsTheirHexEquivalents()
        {
            const string keywordSvg =
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10 10\">" +
                "<g id=\"Dots\"><path fill=\"black\" d=\"M0 0h1v1z\"/><path stroke=\"none\" fill=\"white\" d=\"M0 0h1v1z\"/></g>" +
                "</svg>";

            var plan = IllustrationSvgPlanner.Describe(keywordSvg);
            var dots = plan.Parts.Single(part => part.Id == "Dots");

            CollectionAssert.AreEquivalent(new[] { "#000000", "#FFFFFF" }, dots.Colors.ToArray());
            CollectionAssert.AreEquivalent(new[] { "#000000", "#FFFFFF" }, plan.Palette.ToArray());

            var recolored = IllustrationSvgPlanner.Recolor(keywordSvg, new Dictionary<string, string> { { "#000000", "#BA68C8" } });
            Assert.IsTrue(recolored.Contains("fill=\"#BA68C8\""), recolored);
            Assert.IsTrue(recolored.Contains("fill=\"white\""), recolored);
        }

        [TestMethod]
        public void BrokenSvgIsReportedInsteadOfThrown()
        {
            var plan = IllustrationSvgPlanner.Describe("<svg><g id=\"unterminated\">");

            Assert.IsFalse(plan.IsSafe);
            StringAssert.StartsWith(plan.UnsafeReason, "unparseable svg");
        }
    }
}
