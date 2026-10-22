using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.UI;
using oPenEfficiency.UI.Controls;
using oPenEfficiency.Utils;
using Shape = System.Windows.Shapes.Shape;
using WpfPath = System.Windows.Shapes.Path;

namespace oPenEfficiency.Command.Tests
{
    [TestClass]
    public sealed class FeatureIconParserTests
    {
        [TestMethod]
        public void NestedTransformsPreserveBaseAndAccentGeometry()
        {
            const string svg =
                "<svg xmlns=\"http://www.w3.org/2000/svg\" fill=\"none\" stroke=\"currentColor\">" +
                "<g transform=\"translate(10 2)\">" +
                "<rect x=\"0\" y=\"0\" width=\"2\" height=\"2\"/>" +
                "<g transform=\"translate(0 5)\" data-oe-role=\"accent\">" +
                "<circle cx=\"1\" cy=\"1\" r=\"1\"/>" +
                "</g></g></svg>";

            string path = WriteTemporarySvg(svg);
            try
            {
                var parsed = SvgParser.ParseSvgWithRoles(path);

                Assert.IsNotNull(parsed);
                Assert.IsNotNull(parsed.BaseGeometry);
                Assert.IsNotNull(parsed.AccentGeometry);
                Assert.AreEqual(10d, parsed.BaseGeometry.Bounds.X, 0.001d);
                Assert.AreEqual(12d, parsed.BaseGeometry.Bounds.Right, 0.001d);
                Assert.AreEqual(10d, parsed.AccentGeometry.Bounds.X, 0.001d);
                Assert.AreEqual(7d, parsed.AccentGeometry.Bounds.Y, 0.001d);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [TestMethod]
        public void PickFillColorPreservesElementPaintAndSemanticRole()
        {
            string root = FindRepositoryRoot();
            string path = Path.Combine(
                root,
                "oPenEfficiency",
                "UI",
                "Assets",
                "Icons",
                "BtnPickFillColor.svg");

            var parsed = SvgParser.ParseSvgWithRoles(path);
            var baseElements = parsed.Elements.Where(element => !element.IsAccent).ToList();
            var accent = parsed.Elements.Single(element => element.IsAccent);

            Assert.AreEqual(3, baseElements.Count);
            Assert.IsTrue(baseElements.All(element => element.HasStroke && !element.HasFill));
            Assert.IsTrue(accent.HasFill);
            Assert.IsFalse(accent.HasStroke);
            Assert.AreEqual(FeatureIconAccentRole.ContentColor, accent.AccentRole);
        }

        [DataTestMethod]
        [DataRow("BtnNewSlide", FeatureIconAccentRole.CreateGreen)]
        [DataRow("BtnSpellCheck", FeatureIconAccentRole.SuccessGreen)]
        [DataRow("BtnCheckbox", FeatureIconAccentRole.SuccessGreen)]
        [DataRow("BtnSlideNotes", null)]
        [DataRow("BtnStyleCheck", FeatureIconAccentRole.SuccessGreen)]
        [DataRow("BtnDeleteText", FeatureIconAccentRole.DestructiveRed)]
        [DataRow("BtnNumeration", FeatureIconAccentRole.ActionBlue)]
        [DataRow("BtnPickFillColor", FeatureIconAccentRole.ContentColor)]
        [DataRow("BtnTableHeatmap", FeatureIconAccentRole.ContentColor)]
        public void AuditedSvgCarriesSemanticAccentRole(
            string featureId,
            FeatureIconAccentRole? expectedRole)
        {
            string path = Path.Combine(
                FindRepositoryRoot(),
                "oPenEfficiency",
                "UI",
                "Assets",
                "Icons",
                featureId + ".svg");

            var parsed = SvgParser.ParseSvgWithRoles(path);
            var roles = parsed.Elements
                .Where(element => element.IsAccent)
                .Select(element => element.AccentRole)
                .Distinct()
                .ToList();

            Assert.IsTrue(parsed.Elements
                .Where(element => element.IsAccent)
                .All(element => element.HasExplicitAccentRole));
            CollectionAssert.AreEqual(
                expectedRole.HasValue ? new[] { expectedRole.Value } : Array.Empty<FeatureIconAccentRole>(),
                roles);
        }

        [TestMethod]
        public void HeatmapKeepsContentTonesAndRespectsDisabledAndFeatureColorModes()
        {
            var parsed = SvgParser.ParseSvgWithRoles(Path.Combine(FindRepositoryRoot(),
                "oPenEfficiency", "UI", "Assets", "Icons", "BtnTableHeatmap.svg"));
            var paths = FeatureIconPathFactory.CreatePaths(parsed, IconColorModes.PowerPoint,
                FeatureIconVisualState.Normal, Brushes.Magenta);
            var samples = parsed.Elements.Select((element, index) => new { element, index })
                .Where(item => item.element.IsAccent).ToList();
            Assert.AreEqual(3, samples.Select(item => item.element.FillPaint).Distinct().Count());
            foreach (var sample in samples)
            {
                Assert.AreEqual((Color)ColorConverter.ConvertFromString(sample.element.FillPaint),
                    ((SolidColorBrush)paths[sample.index].Fill).Color);
                paths[sample.index].Resources[FeatureIconColorPolicy.DisabledResourceKey] = Brushes.Gray;
            }
            var originalGeometry = paths.Select(path => path.Data).ToList();
            FeatureIconPathFactory.ApplyState(paths, parsed, IconColorModes.PowerPoint,
                FeatureIconVisualState.Disabled, Brushes.Magenta);
            Assert.IsTrue(samples.All(item => paths[item.index].Fill == Brushes.Gray));
            FeatureIconPathFactory.ApplyState(paths, parsed, IconColorModes.FeatureColors,
                FeatureIconVisualState.Normal, Brushes.Magenta);
            Assert.IsTrue(samples.All(item => paths[item.index].Fill == Brushes.Magenta));
            FeatureIconPathFactory.ApplyState(paths, parsed, IconColorModes.PowerPoint,
                FeatureIconVisualState.Hover, Brushes.Magenta);
            foreach (var sample in samples)
            {
                Assert.AreSame(originalGeometry[sample.index], paths[sample.index].Data);
                Assert.AreEqual((Color)ColorConverter.ConvertFromString(sample.element.FillPaint),
                    ((SolidColorBrush)paths[sample.index].Fill).Color);
            }
        }

        [TestMethod]
        public void SlideSnapshotsKeepPictureAndCreateRolesSeparate()
        {
            var parsed = SvgParser.ParseSvgWithRoles(Path.Combine(FindRepositoryRoot(),
                "oPenEfficiency", "UI", "Assets", "Icons", "BtnSlidePaste.svg"));
            var accents = parsed.Elements.Where(element => element.IsAccent).ToList();
            Assert.IsTrue(accents.All(element => element.HasExplicitAccentRole));
            CollectionAssert.AreEquivalent(new[] { FeatureIconAccentRole.ActionBlue,
                FeatureIconAccentRole.AttentionYellow, FeatureIconAccentRole.CreateGreen },
                accents.Select(element => element.AccentRole).Distinct().ToList());
            Assert.AreEqual(2, parsed.Elements.Count(element => !element.IsAccent),
                "Both the source-slide frame and the picture frame must retain neutral geometry.");
        }

        [TestMethod]
        public void ParserKeepsSuccessGreenIndependentFromCreateGreen()
        {
            const string svg =
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 24 24\" fill=\"currentColor\">" +
                "<rect x=\"18\" y=\"18\" width=\"4\" height=\"4\"/>" +
                "<rect x=\"1\" y=\"1\" width=\"4\" height=\"4\" data-oe-role=\"accent\" " +
                "data-oe-accent-role=\"CreateGreen\"/>" +
                "<rect x=\"10\" y=\"10\" width=\"4\" height=\"4\" data-oe-role=\"accent\" " +
                "data-oe-accent-role=\"SuccessGreen\"/></svg>";

            string path = WriteTemporarySvg(svg);
            try
            {
                var parsed = SvgParser.ParseSvgWithRoles(path);
                var roles = parsed.Elements
                    .Where(element => element.IsAccent)
                    .Select(element => element.AccentRole)
                    .ToList();

                CollectionAssert.AreEqual(
                    new[] { FeatureIconAccentRole.CreateGreen, FeatureIconAccentRole.SuccessGreen },
                    roles);
                Assert.IsTrue(parsed.Elements
                    .Where(element => element.IsAccent)
                    .All(element => element.HasExplicitAccentRole));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [TestMethod]
        public void PathFactoryPreservesAuthoredFillAndStroke()
        {
            string path = Path.Combine(
                FindRepositoryRoot(),
                "oPenEfficiency",
                "UI",
                "Assets",
                "Icons",
                "BtnPickFillColor.svg");
            var parsed = SvgParser.ParseSvgWithRoles(path);

            var paths = FeatureIconPathFactory.CreatePaths(
                parsed,
                IconColorModes.PowerPoint,
                FeatureIconVisualState.Normal,
                Brushes.Magenta);

            ParsedIconElement baseElement = parsed.Elements.First(element => !element.IsAccent);
            ParsedIconElement accentElement = parsed.Elements.Single(element => element.IsAccent);
            WpfPath basePath = paths[parsed.Elements.IndexOf(baseElement)];
            WpfPath accentPath = paths[parsed.Elements.IndexOf(accentElement)];

            Assert.AreEqual(Brushes.Transparent, basePath.Fill);
            Assert.AreNotEqual(DependencyProperty.UnsetValue, basePath.ReadLocalValue(Shape.StrokeProperty));
            Assert.AreEqual(2d, basePath.StrokeThickness, 0.001d);
            Assert.IsNull(accentPath.Stroke);
            Assert.AreNotEqual(DependencyProperty.UnsetValue, accentPath.ReadLocalValue(Shape.FillProperty));
            Assert.AreEqual(0d, accentPath.StrokeThickness, 0.001d);
        }

        [TestMethod]
        public void PathFactoryKeepsEveryElementOnTheShared24PixelCanvas()
        {
            string path = Path.Combine(
                FindRepositoryRoot(),
                "oPenEfficiency",
                "UI",
                "Assets",
                "Icons",
                "BtnNewSlide.svg");
            var parsed = SvgParser.ParseSvgWithRoles(path);

            var paths = FeatureIconPathFactory.CreatePaths(
                parsed,
                IconColorModes.PowerPoint,
                FeatureIconVisualState.Normal,
                Brushes.Magenta);

            Assert.IsTrue(paths.Count > 1);
            Assert.IsTrue(paths.All(item => item.Width == 24d && item.Height == 24d));
            Assert.IsTrue(paths.All(item => item.Stretch == Stretch.None));
            Assert.IsTrue(paths.All(item => item.Data.Bounds.X >= 0d));
            Assert.IsTrue(paths.All(item => item.Data.Bounds.Y >= 0d));
            Assert.IsTrue(paths.All(item => item.Data.Bounds.Right <= 24d));
            Assert.IsTrue(paths.All(item => item.Data.Bounds.Bottom <= 24d));
        }

        [TestMethod]
        public void ApplyStateReSkinsExistingPathsInsteadOfReplacingThem()
        {
            // Non-24 viewBox so the state round-trip also covers viewport-scaled strokes.
            const string svg =
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"100 200 200 100\" " +
                "fill=\"currentColor\" stroke=\"black\" stroke-width=\"4\">" +
                "<rect x=\"100\" y=\"200\" width=\"200\" height=\"100\"/>" +
                "<rect x=\"150\" y=\"225\" width=\"50\" height=\"50\" " +
                "data-oe-role=\"accent\" data-oe-accent-role=\"ActionBlue\"/></svg>";

            string path = WriteTemporarySvg(svg);
            try
            {
                var parsed = SvgParser.ParseSvgWithRoles(path);

                var paths = FeatureIconPathFactory.CreatePaths(
                    parsed,
                    IconColorModes.PowerPoint,
                    FeatureIconVisualState.Normal,
                    Brushes.Magenta);
                var originals = paths.ToList();
                var geometry = originals.Select(item => item.Data).ToList();
                var thickness = originals.Select(item => item.StrokeThickness).ToList();
                Assert.IsTrue(thickness.Any(value => value > 0d && Math.Abs(value - 4d) > 0.001d),
                    "expected at least one viewport-scaled stroke thickness");

                // State, color-mode, and stored-color changes must only re-skin the same instances;
                // replacing paths mid mouse-dispatch silently swallows clicks (PQ-015/PQ-016).
                FeatureIconPathFactory.ApplyState(
                    paths,
                    parsed,
                    IconColorModes.FeatureColors,
                    FeatureIconVisualState.Disabled,
                    Brushes.Lime);

                Assert.AreEqual(originals.Count, paths.Count);
                for (int i = 0; i < originals.Count; i++)
                {
                    Assert.AreSame(originals[i], paths[i], "ApplyState must not replace Path instances.");
                    Assert.AreSame(geometry[i], paths[i].Data, "ApplyState must not replace geometry.");
                }

                FeatureIconPathFactory.ApplyState(
                    paths,
                    parsed,
                    IconColorModes.PowerPoint,
                    FeatureIconVisualState.Normal,
                    Brushes.Magenta);

                for (int i = 0; i < originals.Count; i++)
                {
                    Assert.AreEqual(thickness[i], paths[i].StrokeThickness, 0.0001d,
                        "ApplyState must keep the viewport-scaled stroke thickness.");
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [TestMethod]
        public void PathFactoryNormalizesNon24ViewBoxIntoSharedCanvas()
        {
            const string svg =
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"100 200 200 100\" " +
                "fill=\"currentColor\">" +
                "<rect x=\"100\" y=\"200\" width=\"200\" height=\"100\"/>" +
                "<rect x=\"150\" y=\"225\" width=\"50\" height=\"50\" " +
                "data-oe-role=\"accent\" data-oe-accent-role=\"ActionBlue\"/>" +
                "</svg>";

            string path = WriteTemporarySvg(svg);
            try
            {
                var parsed = SvgParser.ParseSvgWithRoles(path);
                var paths = FeatureIconPathFactory.CreatePaths(
                    parsed,
                    IconColorModes.PowerPoint,
                    FeatureIconVisualState.Normal,
                    Brushes.Magenta);

                Assert.AreEqual(2, paths.Count);
                AssertRect(paths[0].Data.Bounds, 0d, 6d, 24d, 12d);
                AssertRect(paths[1].Data.Bounds, 6d, 9d, 6d, 6d);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [TestMethod]
        public void EveryCatalogCommandIconHasVisibleGeometryInsideSharedCanvas()
        {
            string root = FindRepositoryRoot();
            string iconDirectory = Path.Combine(root, "oPenEfficiency", "UI", "Assets", "Icons");
            var failures = new System.Collections.Generic.List<string>();

            foreach (var record in FeatureLibrary.Catalog.Where(record =>
                record.ExecutionKind == oPenEfficiency.Services.Attributes.FeatureExecutionKind.Command))
            {
                string path = Path.Combine(iconDirectory, record.Id + ".svg");
                var parsed = SvgParser.ParseSvgWithRoles(path);
                var paths = FeatureIconPathFactory.CreatePaths(
                    parsed,
                    IconColorModes.PowerPoint,
                    FeatureIconVisualState.Normal,
                    Brushes.Magenta);
                Rect bounds = Rect.Empty;
                foreach (WpfPath item in paths)
                    bounds.Union(item.Data.Bounds);

                if (paths.Count == 0 || bounds.IsEmpty || bounds.Width <= 0d || bounds.Height <= 0d ||
                    bounds.Right <= 0d || bounds.Bottom <= 0d || bounds.X >= 24d || bounds.Y >= 24d)
                {
                    failures.Add(record.Id + ": " + bounds);
                }
            }

            Assert.AreEqual(0, failures.Count,
                "Catalog icons without visible in-viewport geometry:" + Environment.NewLine +
                string.Join(Environment.NewLine, failures));
        }

        [TestMethod]
        public void ParserAppliesAuthoredTransformOrigin()
        {
            const string svg =
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 24 24\">" +
                "<g style=\"transform-origin: 12px 12px 0px;\" " +
                "transform=\"matrix(0, -1, 1, 0, 0, 0)\">" +
                "<rect x=\"10\" y=\"2\" width=\"4\" height=\"6\"/>" +
                "</g></svg>";

            string path = WriteTemporarySvg(svg);
            try
            {
                var parsed = SvgParser.ParseSvgWithRoles(path);

                AssertRect(parsed.Elements.Single().Geometry.Bounds, 2d, 10d, 6d, 4d);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [STATestMethod]
        public void EveryCatalogCommandIconRendersAtRequiredSizes()
        {
            string root = FindRepositoryRoot();
            string iconDirectory = Path.Combine(root, "oPenEfficiency", "UI", "Assets", "Icons");
            var failures = new System.Collections.Generic.List<string>();

            foreach (var record in FeatureLibrary.Catalog.Where(record =>
                record.ExecutionKind == oPenEfficiency.Services.Attributes.FeatureExecutionKind.Command))
            {
                string path = Path.Combine(iconDirectory, record.Id + ".svg");
                var parsed = SvgParser.ParseSvgWithRoles(path);
                var paths = FeatureIconPathFactory.CreatePaths(
                    parsed,
                    IconColorModes.PowerPoint,
                    FeatureIconVisualState.Normal,
                    Brushes.Magenta);

                foreach (int size in new[] { 16, 24, 48 })
                {
                    if (!RendersPixels(paths, parsed.Elements, size, Brushes.Black) ||
                        !RendersPixels(paths, parsed.Elements, size, Brushes.White))
                    {
                        failures.Add(record.Id + "@" + size);
                    }
                }
            }

            Assert.AreEqual(0, failures.Count,
                "Catalog icons with empty Light/Dark renders:" + Environment.NewLine +
                string.Join(Environment.NewLine, failures));
        }

        [TestMethod]
        public void FeatureIconDefaultColorModeIsPowerPoint()
        {
            var metadata = FeatureIcon.ColorModeProperty.GetMetadata(typeof(FeatureIcon));

            Assert.AreEqual(IconColorModes.PowerPoint, metadata.DefaultValue);
        }

        [TestMethod]
        public void CommandIconPipelineHasNoHoverAssetsOrRuntimeSuffix()
        {
            string root = FindRepositoryRoot();
            string iconDirectory = Path.Combine(
                root,
                "oPenEfficiency",
                "UI",
                "Assets",
                "Icons");
            string[] hoverAssets = Directory.GetFiles(iconDirectory, "*_hover.svg");

            Assert.AreEqual(0, hoverAssets.Length, "Command-icon hover assets must stay removed.");

            string iconService = File.ReadAllText(Path.Combine(
                root,
                "oPenEfficiency",
                "Services",
                "IconService.cs"));
            string featureIcon = File.ReadAllText(Path.Combine(
                root,
                "oPenEfficiency",
                "UI",
                "Controls",
                "FeatureIcon.xaml.cs"));

            Assert.IsFalse(iconService.Contains("_hover"));
            Assert.IsFalse(featureIcon.Contains("_hover"));
            Assert.IsFalse(featureIcon.Contains("UseHoverGeometry"));
        }

        private static string WriteTemporarySvg(string source)
        {
            string path = Path.Combine(
                Path.GetTempPath(),
                "oPenEfficiency-svg-test-" + Guid.NewGuid().ToString("N") + ".svg");
            File.WriteAllText(path, source);
            return path;
        }

        private static void AssertRect(Rect actual, double x, double y, double width, double height)
        {
            Assert.AreEqual(x, actual.X, 0.001d);
            Assert.AreEqual(y, actual.Y, 0.001d);
            Assert.AreEqual(width, actual.Width, 0.001d);
            Assert.AreEqual(height, actual.Height, 0.001d);
        }

        private static bool RendersPixels(
            System.Collections.Generic.IReadOnlyList<WpfPath> paths,
            System.Collections.Generic.IList<ParsedIconElement> elements,
            int size,
            Brush foreground)
        {
            var visual = new DrawingVisual();
            using (DrawingContext drawing = visual.RenderOpen())
            {
                drawing.PushTransform(new ScaleTransform(size / 24d, size / 24d));
                for (int index = 0; index < paths.Count; index++)
                {
                    ParsedIconElement element = elements[index];
                    drawing.DrawGeometry(
                        element.HasFill ? foreground : null,
                        element.HasStroke
                            ? new Pen(foreground, paths[index].StrokeThickness)
                            : null,
                        paths[index].Data);
                }
                drawing.Pop();
            }

            var bitmap = new RenderTargetBitmap(size, size, 96d, 96d, PixelFormats.Pbgra32);
            bitmap.Render(visual);
            var pixels = new byte[size * size * 4];
            bitmap.CopyPixels(pixels, size * 4, 0);
            for (int index = 3; index < pixels.Length; index += 4)
            {
                if (pixels[index] != 0) return true;
            }
            return false;
        }

        private static string FindRepositoryRoot()
        {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null)
            {
                if (File.Exists(Path.Combine(directory.FullName, "oPenEfficiency.slnx")))
                {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }

            throw new DirectoryNotFoundException("Repository root not found.");
        }
    }
}
