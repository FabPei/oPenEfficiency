using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public class IllustrationCatalogTests
    {
        private const string DecomposableSvg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 100 100\">" +
            "<g id=\"Head\"><path fill=\"#E6E6E6\" d=\"M0 0h1v1z\"/></g>" +
            "<g id=\"head/Afro\"><path fill=\"#263238\" d=\"M0 0h1v1z\"/></g>" +
            "<g id=\"pose/sitting/closed_legs-2\"><path fill=\"#263238\" d=\"M0 0h1v1z\"/></g>" +
            "</svg>";

        private const string FlatDoodleSvg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 1024 768\">" +
            "<g id=\"Illustrations/coffee\"><path fill=\"#000000\" d=\"M0 0h1v1z\"/></g>" +
            "</svg>";

        private const string UnsafeSvg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\"><g id=\"Head\"><script>x</script></g></svg>";

        [TestMethod]
        public void BuildDerivesItemsFromTheCollectionFolders()
        {
            var files = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "Peeps/peep-4.svg", DecomposableSvg },
                { "Peeps/peep-13.svg", DecomposableSvg },
                { "Doodles/coffee.svg", FlatDoodleSvg }
            };

            var result = BuildFrom(files);

            Assert.AreEqual(3, result.Items.Count);
            Assert.IsTrue(result.IsClean);

            var peep = result.Items.Single(item => item.Id == "open-peeps/peep-4");
            Assert.AreEqual("Peep 4", peep.Name);
            Assert.AreEqual("Peeps/peep-4.svg", peep.RelativePath);
            Assert.AreEqual(true, peep.Decomposable);
            CollectionAssert.Contains(peep.Parts, "head/Afro");
            CollectionAssert.Contains(peep.Tags, "afro");
            CollectionAssert.Contains(peep.Tags, "closed legs 2");

            var doodle = result.Items.Single(item => item.Id == "open-doodles/coffee");
            Assert.IsFalse(doodle.Decomposable);
        }

        [TestMethod]
        public void ItemsAreOrderedDeterministically()
        {
            var files = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "Peeps/peep-13.svg", DecomposableSvg },
                { "Peeps/peep-4.svg", DecomposableSvg },
                { "Doodles/coffee.svg", FlatDoodleSvg }
            };

            var ids = BuildFrom(files).Items.Select(item => item.Id).ToArray();

            CollectionAssert.AreEqual(new[] { "open-doodles/coffee", "open-peeps/peep-13", "open-peeps/peep-4" }, ids);
        }

        [TestMethod]
        public void UnsafeAssetsAreQuarantinedInsteadOfPublished()
        {
            var files = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "Peeps/peep-4.svg", DecomposableSvg },
                { "Peeps/peep-13.svg", UnsafeSvg }
            };

            var result = BuildFrom(files);

            Assert.AreEqual(1, result.Items.Count);
            Assert.AreEqual(1, result.UnsafeFiles.Count);
            StringAssert.Contains(result.UnsafeFiles[0], "Peeps/peep-13.svg");
            Assert.IsFalse(result.IsClean);
        }

        [TestMethod]
        public void MissingAssetsAreReported()
        {
            var files = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "Peeps/peep-4.svg", DecomposableSvg },
                { "Peeps/peep-13.svg", null }
            };

            var result = BuildFrom(files);

            Assert.AreEqual(1, result.Items.Count);
            CollectionAssert.AreEqual(new[] { "Peeps/peep-13.svg" }, result.MissingFiles.ToArray());
        }

        [TestMethod]
        public void SearchMatchesNamesTagsAndPartIds()
        {
            var files = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "Peeps/peep-4.svg", DecomposableSvg },
                { "Doodles/coffee.svg", FlatDoodleSvg }
            };
            var items = BuildFrom(files).Items;

            CollectionAssert.AreEqual(new[] { "open-peeps/peep-4" },
                IllustrationCatalog.Search(items, "afro").Select(item => item.Id).ToArray());
            CollectionAssert.AreEqual(new[] { "open-peeps/peep-4" },
                IllustrationCatalog.Search(items, "sitting legs").Select(item => item.Id).ToArray());
            CollectionAssert.AreEqual(new[] { "open-doodles/coffee" },
                IllustrationCatalog.Search(items, "doodle").Select(item => item.Id).ToArray());
            Assert.AreEqual(2, IllustrationCatalog.Search(items, string.Empty).Count);
            Assert.AreEqual(0, IllustrationCatalog.Search(items, "missing").Count);
        }

        [TestMethod]
        public void ManifestRoundTripsLicenseFacts()
        {
            const string json = "{\"Version\":\"1.0\",\"Collections\":[{\"Id\":\"open-peeps\",\"Name\":\"Open Peeps\"," +
                "\"License\":\"CC0 1.0 Universal\",\"LicenseUrl\":\"https://creativecommons.org/publicdomain/zero/1.0/\"," +
                "\"SourceUrl\":\"https://www.openpeeps.com/\",\"AttributionRequired\":false,\"AllowsBundling\":true," +
                "\"RelativeFolder\":\"Peeps\",\"SupportsPartDecomposition\":true}]}";

            var manifest = IllustrationCatalog.LoadManifest(json);

            Assert.AreEqual(1, manifest.Collections.Count);
            var collection = manifest.Collections[0];
            Assert.AreEqual("CC0 1.0 Universal", collection.License);
            Assert.IsFalse(collection.AttributionRequired);
            Assert.IsTrue(collection.AllowsBundling);
            Assert.IsTrue(collection.SupportsPartDecomposition);
        }

        private static IllustrationCatalogBuildResult BuildFrom(IDictionary<string, string> files)
        {
            var manifest = new IllustrationManifest
            {
                Collections = new List<IllustrationCollectionInfo>
                {
                    new IllustrationCollectionInfo
                    {
                        Id = "open-peeps",
                        Name = "Open Peeps",
                        RelativeFolder = "Peeps",
                        License = "CC0 1.0 Universal",
                        AllowsBundling = true,
                        SupportsPartDecomposition = true
                    },
                    new IllustrationCollectionInfo
                    {
                        Id = "open-doodles",
                        Name = "Open Doodles",
                        RelativeFolder = "Doodles",
                        License = "CC0 1.0 Universal",
                        AllowsBundling = true,
                        SupportsPartDecomposition = false
                    }
                }
            };

            return IllustrationCatalog.Build(
                manifest,
                folder => files.Keys.Where(path => path.StartsWith(folder + "/", StringComparison.Ordinal)).ToList(),
                path =>
                {
                    string value;
                    return files.TryGetValue(path, out value) ? value : null;
                });
        }
    }
}
