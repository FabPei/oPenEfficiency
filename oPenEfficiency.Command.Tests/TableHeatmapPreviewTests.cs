using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Features.Tables;

namespace oPenEfficiency.Command.Tests
{
    /// <summary>
    /// REG-TABLEHEATMAP-PREVIEW-001: the dialog live preview and applied cell formatting share
    /// TableHeatmapApplier.ComputeColor / ComputeIcon. These characterization pins lock that shared
    /// math so a preview or applier change cannot silently drift the two apart.
    /// </summary>
    [TestClass]
    public sealed class TableHeatmapPreviewTests
    {
        private static TableHeatmapOptions Options(int colorPreset = 0, int iconPreset = 0) =>
            new TableHeatmapOptions { ColorPresetIndex = colorPreset, IconPresetIndex = iconPreset };

        [TestMethod]
        public void ThreeColorPresetInterpolatesThroughItsMiddleColor()
        {
            var options = Options(colorPreset: 0); // Red to Yellow to Green
            Assert.AreEqual(System.Drawing.Color.FromArgb(248, 105, 107), TableHeatmapApplier.ComputeColor(0.0, 0, options));
            Assert.AreEqual(System.Drawing.Color.FromArgb(255, 235, 132), TableHeatmapApplier.ComputeColor(0.5, 0, options));
            Assert.AreEqual(System.Drawing.Color.FromArgb(99, 190, 123), TableHeatmapApplier.ComputeColor(1.0, 0, options));
            // Half-way to the middle interpolates from low toward mid.
            Assert.AreEqual(System.Drawing.Color.FromArgb(251, 170, 119), TableHeatmapApplier.ComputeColor(0.25, 0, options));
        }

        [TestMethod]
        public void TwoColorPresetInterpolatesDirectlyWithoutMiddle()
        {
            var options = Options(colorPreset: 6); // White to Red
            Assert.AreEqual(System.Drawing.Color.FromArgb(255, 255, 255), TableHeatmapApplier.ComputeColor(0.0, 6, options));
            Assert.AreEqual(System.Drawing.Color.FromArgb(248, 105, 107), TableHeatmapApplier.ComputeColor(1.0, 6, options));
            Assert.AreEqual(System.Drawing.Color.FromArgb(251, 180, 181), TableHeatmapApplier.ComputeColor(0.5, 6, options));
        }

        [TestMethod]
        public void CustomPresetUsesTheDialogColors()
        {
            var options = new TableHeatmapOptions
            {
                ColorPresetIndex = 12,
                CustomLow = System.Drawing.Color.FromArgb(10, 20, 30),
                CustomMid = null,
                CustomHigh = System.Drawing.Color.FromArgb(110, 120, 130)
            };
            Assert.AreEqual(System.Drawing.Color.FromArgb(10, 20, 30), TableHeatmapApplier.ComputeColor(0.0, 12, options));
            Assert.AreEqual(System.Drawing.Color.FromArgb(60, 70, 80), TableHeatmapApplier.ComputeColor(0.5, 12, options));
            Assert.AreEqual(System.Drawing.Color.FromArgb(110, 120, 130), TableHeatmapApplier.ComputeColor(1.0, 12, options));
        }

        [TestMethod]
        public void IconPresetsUseTheDocumentedThresholds()
        {
            TableHeatmapApplier.ComputeIcon(1.0, 0, out string highIcon, out System.Drawing.Color highColor);
            TableHeatmapApplier.ComputeIcon(0.5, 0, out string midIcon, out System.Drawing.Color midColor);
            TableHeatmapApplier.ComputeIcon(0.0, 0, out string lowIcon, out System.Drawing.Color lowColor);
            Assert.AreEqual("\u25b2", highIcon);
            Assert.AreEqual("\u25b6", midIcon);
            Assert.AreEqual("\u25bc", lowIcon);
            Assert.AreEqual(System.Drawing.Color.FromArgb(99, 190, 123), highColor);
            Assert.AreEqual(System.Drawing.Color.FromArgb(255, 235, 132), midColor);
            Assert.AreEqual(System.Drawing.Color.FromArgb(248, 105, 107), lowColor);

            TableHeatmapApplier.ComputeIcon(0.74, 1, out string arrowIcon, out _); // 4 arrows: 0.50..0.75
            Assert.AreEqual("\u2197", arrowIcon);
            TableHeatmapApplier.ComputeIcon(0.33, 3, out string indicatorIcon, out _); // indicators: 0.33..0.66
            Assert.AreEqual("\u2757", indicatorIcon);
        }
    }
}
