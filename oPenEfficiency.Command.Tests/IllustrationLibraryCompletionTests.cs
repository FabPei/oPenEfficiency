using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Command.Tests
{
    [TestClass]
    public sealed class IllustrationLibraryCompletionTests
    {
        [TestMethod]
        public void ThumbnailRendererPreservesSolidPaintAndEllipseGeometry()
        {
            const string svg =
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 40 20\">" +
                "<ellipse cx=\"20\" cy=\"10\" rx=\"18\" ry=\"8\" fill=\"#123456\" " +
                "stroke=\"#ABCDEF\" stroke-width=\"2\"/></svg>";
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".svg");
            File.WriteAllText(path, svg);
            try
            {
                var parsed = SvgParser.ParseSvgWithRoles(path);
                Assert.IsNotNull(parsed);
                Assert.AreEqual(1, parsed.Elements.Count);
                Assert.AreEqual("#123456", parsed.Elements[0].FillPaint);
                Assert.AreEqual("#ABCDEF", parsed.Elements[0].StrokePaint);
                Assert.AreEqual(36d, parsed.Elements[0].Geometry.Bounds.Width, 0.001d);
                Assert.AreEqual(16d, parsed.Elements[0].Geometry.Bounds.Height, 0.001d);

                var thumbnail = IllustrationThumbnailRenderer.Create(path);
                Assert.IsNotNull(thumbnail);
                Assert.IsTrue(thumbnail.IsFrozen);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [TestMethod]
        public void EveryBundledIllustrationCanProduceAVectorThumbnail()
        {
            string root = Path.Combine(FindRepositoryRoot(), "oPenEfficiency", "Resources", "Illustrations");
            var paths = Directory.GetFiles(root, "*.svg", SearchOption.AllDirectories);

            Assert.AreEqual(421, paths.Length, "The bundled illustration inventory changed; update the manifest gate and this preview contract together.");
            var failures = paths
                .Where(path => IllustrationThumbnailRenderer.Create(path) == null)
                .Select(path => path.Substring(root.Length).TrimStart(Path.DirectorySeparatorChar))
                .ToList();

            Assert.AreEqual(0, failures.Count,
                "These bundled SVGs have no renderable vector preview: " + string.Join(", ", failures));
        }

        [TestMethod]
        public void PickerKeepsThumbnailAndKeyboardInsertionContracts()
        {
            string root = FindRepositoryRoot();
            string xaml = File.ReadAllText(Path.Combine(root, "oPenEfficiency", "UI", "Dialogs", "IllustrationLibraryWindow.xaml"));
            string code = File.ReadAllText(Path.Combine(root, "oPenEfficiency", "UI", "Dialogs", "IllustrationLibraryWindow.xaml.cs"));

            StringAssert.Contains(xaml, "Source=\"{Binding Thumbnail}\"");
            StringAssert.Contains(xaml, "KeyDown=\"FigureSearchBox_KeyDown\"");
            StringAssert.Contains(xaml, "KeyDown=\"FigureList_KeyDown\"");
            StringAssert.Contains(xaml, "MouseDoubleClick=\"FigureList_MouseDoubleClick\"");
            StringAssert.Contains(code, "IllustrationThumbnailRenderer.Create(ThumbnailPath)");
            StringAssert.Contains(code, "InsertSelectedFigure();");
            StringAssert.Contains(code, "Keyboard.Modifiers.HasFlag(ModifierKeys.Control)");
        }

        [TestMethod]
        public void InsertionStartsOnePowerPointUndoUnitBeforeMutation()
        {
            string source = File.ReadAllText(Path.Combine(
                FindRepositoryRoot(), "oPenEfficiency", "Services", "IllustrationInsertionService.cs"));
            int undo = source.IndexOf("app.StartNewUndoEntry();", StringComparison.Ordinal);
            int insert = source.IndexOf("slide.Shapes.AddPicture(", StringComparison.Ordinal);

            Assert.IsTrue(undo >= 0, "The illustration insertion has no explicit PowerPoint undo boundary.");
            Assert.IsTrue(insert > undo, "The undo boundary must be established before the first slide mutation.");
            StringAssert.Contains(source, "IllustrationInsertionService.Insert.StartNewUndoEntry");
        }

        private static string FindRepositoryRoot()
        {
            var current = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (current != null)
            {
                if (File.Exists(Path.Combine(current.FullName, "oPenEfficiency.slnx"))) return current.FullName;
                current = current.Parent;
            }

            throw new DirectoryNotFoundException("Could not locate the repository root.");
        }
    }
}
