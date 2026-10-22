using System;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;

namespace oPenEfficiency.Command.Tests
{
    [TestClass]
    public sealed class JsonFileStoreTests
    {
        private string _directory;

        [TestInitialize]
        public void SetUp()
        {
            _directory = Path.Combine(Path.GetTempPath(), "oPE-json", Guid.NewGuid().ToString("N"));
        }

        [TestCleanup]
        public void TearDown()
        {
            if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
        }

        [TestMethod]
        public void JsonFileStoreRoundTripsUnicodeAndOverwritesAtomically()
        {
            var store = new JsonFileStore(_directory);
            Assert.IsTrue(store.TrySave(
                new SeriesConfig { SeriesId = "Änderung 😀", Count = 1 },
                "settings.json"));
            Assert.AreEqual("Änderung 😀", store.Load<SeriesConfig>("settings.json").SeriesId);

            Assert.IsTrue(store.TrySave(
                new SeriesConfig { SeriesId = "second", Count = 2 },
                "settings.json"));

            SeriesConfig loaded = store.Load<SeriesConfig>("settings.json");
            Assert.AreEqual("second", loaded.SeriesId);
            Assert.AreEqual(2, loaded.Count);
            Assert.AreEqual(1, Directory.GetFiles(_directory).Length);
            Assert.IsFalse(Directory.GetFiles(_directory).Any(path => path.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase)));
        }

        [TestMethod]
        public void JsonFileStoreReturnsDefaultForMissingOrMalformedData()
        {
            var store = new JsonFileStore(_directory);
            Assert.IsNull(store.Load<SeriesConfig>("missing.json"));

            Directory.CreateDirectory(_directory);
            File.WriteAllText(Path.Combine(_directory, "broken.json"), "{broken");
            Assert.IsNull(store.Load<SeriesConfig>("broken.json"));
        }

        [TestMethod]
        public void JsonFileStoreRejectsPathTraversalAndAbsolutePaths()
        {
            var store = new JsonFileStore(_directory);
            string outside = Path.Combine(Path.GetDirectoryName(_directory), "outside.json");

            Assert.IsFalse(store.TrySave(new SeriesConfig(), "..\\outside.json"));
            Assert.IsFalse(store.TrySave(new SeriesConfig(), outside));
            Assert.IsFalse(File.Exists(outside));
            Assert.IsNull(store.Load<SeriesConfig>("..\\outside.json"));
        }

        [TestMethod]
        public void SerializationFailurePreservesExistingJsonAndCleansTemporaryFiles()
        {
            var store = new JsonFileStore(_directory);
            Assert.IsTrue(store.TrySave(new SeriesConfig { SeriesId = "safe" }, "settings.json"));
            string original = File.ReadAllText(Path.Combine(_directory, "settings.json"));

            Assert.IsFalse(store.TrySave(new UnsupportedContract(), "settings.json"));

            Assert.AreEqual(original, File.ReadAllText(Path.Combine(_directory, "settings.json")));
            Assert.AreEqual(1, Directory.GetFiles(_directory).Length);
        }

        [DataContract]
        private sealed class UnsupportedContract
        {
            [DataMember]
            public Action Callback { get; set; } = () => { };
        }
    }
}
