using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class StickerPreviewModelTests
    {
        [TestMethod]
        public void EmptySettingsProduceSafeRepresentativePreviews()
        {
            var sticky = StickerPreviewModel.CreateStickyNote(new SidebarConfig
            {
                StickyNoteDefaultText = " ",
                StickyNoteBackgroundColor = "invalid",
                StickyNoteFontColor = null
            }, "Aptos");
            var status = StickerPreviewModel.CreateStatusSticker(new SidebarConfig
            {
                IllustrativeStickerTexts = " , ",
                IllustrativeStickerColor = "broken"
            }, "Aptos", new DateTime(2026, 9, 16), "Reviewer");

            Assert.AreEqual("Sticky note", sticky.Text);
            Assert.AreEqual("Aptos", sticky.FontFamily);
            Assert.AreEqual("#FFF494", sticky.BackgroundColor);
            Assert.AreEqual("#323232", sticky.FontColor);
            Assert.AreEqual("Work in progress", status.Text);
            Assert.AreEqual("#FF0000", status.Color);
        }

        [TestMethod]
        public void StatusPreviewUsesFirstLabelAndConfiguredMetadata()
        {
            var state = StickerPreviewModel.CreateStatusSticker(new SidebarConfig
            {
                IllustrativeStickerTexts = "  Draft review, Final ",
                IllustrativeStickerIncludeDate = true,
                IllustrativeStickerIncludeAuthor = true,
                IllustrativeStickerFont = "Arial",
                IllustrativeStickerFontBold = false,
                IllustrativeStickerFontItalic = true,
                IllustrativeStickerPosition = "BottomLeft"
            }, "Aptos", new DateTime(2026, 9, 16), "Fabian");

            Assert.AreEqual("Draft review · 2026-09-16 · Fabian", state.Text);
            Assert.AreEqual("Arial", state.FontFamily);
            Assert.IsFalse(state.Bold);
            Assert.IsTrue(state.Italic);
            Assert.AreEqual("Bottom left", state.PositionLabel);
        }

        [TestMethod]
        public void StickyPreviewKeepsThreeSupportedVisualStylesDistinct()
        {
            var config = new SidebarConfig();
            var theme = StickerPreviewModel.CreateStickyNote(config, null, "Theme");
            var consulting = StickerPreviewModel.CreateStickyNote(config, null, "Consulting");
            var greenB = StickerPreviewModel.CreateStickyNote(config, null, "GreenB");

            Assert.AreEqual("#FFF494", theme.BackgroundColor);
            Assert.AreEqual("#323232", theme.FontColor);
            Assert.AreEqual("#FFFF99", consulting.BackgroundColor);
            Assert.AreEqual("#333333", consulting.FontColor);
            Assert.AreEqual("#FFFF00", greenB.BackgroundColor);
            Assert.AreEqual("#575757", greenB.FontColor);
        }

        [TestMethod]
        public void CustomColorsAreScopedToSelectedStickyStyle()
        {
            var config = new SidebarConfig();

            StickerPreviewModel.SetStickyNoteColors(config, "Consulting", "#123456", "#ABCDEF");

            Assert.AreEqual("#FFF494", StickerPreviewModel.CreateStickyNote(config, null, "Theme").BackgroundColor);
            Assert.AreEqual("#123456", StickerPreviewModel.CreateStickyNote(config, null, "Consulting").BackgroundColor);
            Assert.AreEqual("#ABCDEF", StickerPreviewModel.CreateStickyNote(config, null, "Consulting").FontColor);
            Assert.AreEqual("#FFFF00", StickerPreviewModel.CreateStickyNote(config, null, "GreenB").BackgroundColor);
        }

        [TestMethod]
        public void LegacyCustomColorIsRetainedForFormerlyActiveStyleOnly()
        {
            var config = new SidebarConfig
            {
                StickyNoteStyle = "GreenB",
                StickyNoteBackgroundColor = "#112233",
                StickyNoteFontColor = "#445566",
                StickyNoteThemeBackgroundColor = null,
                StickyNoteThemeFontColor = null,
                StickyNoteConsultingBackgroundColor = null,
                StickyNoteConsultingFontColor = null,
                StickyNoteGreenBBackgroundColor = null,
                StickyNoteGreenBFontColor = null
            };

            Assert.AreEqual("#112233", StickerPreviewModel.CreateStickyNote(config, null, "GreenB").BackgroundColor);
            Assert.AreEqual("#445566", StickerPreviewModel.CreateStickyNote(config, null, "GreenB").FontColor);
            Assert.AreEqual("#FFF494", StickerPreviewModel.CreateStickyNote(config, null, "Theme").BackgroundColor);
        }

        [TestMethod]
        public void LegacyPaletteNormalizationPinsCustomColorsToFormerActiveStyle()
        {
            var config = new SidebarConfig
            {
                StickyNoteStyle = "GreenB",
                StickyNoteBackgroundColor = "#112233",
                StickyNoteFontColor = "#445566",
                StickyNoteThemeBackgroundColor = null,
                StickyNoteThemeFontColor = null,
                StickyNoteConsultingBackgroundColor = null,
                StickyNoteConsultingFontColor = null,
                StickyNoteGreenBBackgroundColor = null,
                StickyNoteGreenBFontColor = null
            };

            Assert.IsTrue(StickerPreviewModel.NormalizeStickyNoteColors(config));
            Assert.AreEqual("#112233", config.StickyNoteGreenBBackgroundColor);
            Assert.AreEqual("#445566", config.StickyNoteGreenBFontColor);
            Assert.AreEqual("#FFF494", config.StickyNoteThemeBackgroundColor);
            Assert.AreEqual("#FFFF99", config.StickyNoteConsultingBackgroundColor);

            config.StickyNoteStyle = "Consulting";
            Assert.AreEqual("#FFFF99", StickerPreviewModel.CreateStickyNote(config, null).BackgroundColor);
            Assert.AreEqual("#112233", StickerPreviewModel.CreateStickyNote(config, null, "GreenB").BackgroundColor);
        }
    }
}
