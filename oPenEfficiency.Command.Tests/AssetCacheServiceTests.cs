using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Services;

namespace oPenEfficiency.Command.Tests
{
    [TestClass]
    public sealed class AssetCacheServiceTests
    {
        private string _root;
        private List<string> _sharedFolders;
        private int _saveCalls;

        [TestInitialize]
        public void SetUp()
        {
            _root = Path.Combine(Path.GetTempPath(), "oPenEfficiency-tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
            _sharedFolders = new List<string>();
            _saveCalls = 0;
        }

        [TestCleanup]
        public void TearDown()
        {
            if (Directory.Exists(_root)) Directory.Delete(_root, true);
        }

        [TestMethod]
        public void SharedFolderMatchingRespectsDirectoryBoundaries()
        {
            string shared = Path.Combine(_root, "Assets");
            _sharedFolders.Add(shared);
            var service = CreateService();

            Assert.IsTrue(service.IsInSharedFolder(Path.Combine(shared, "nested", "deck.pptx")));
            Assert.IsTrue(service.IsInSharedFolder(shared));
            Assert.IsFalse(service.IsInSharedFolder(Path.Combine(_root, "AssetsOld", "deck.pptx")));
            Assert.IsFalse(service.IsInSharedFolder(null));
        }

        [TestMethod]
        public void SharedFoldersAreDeduplicatedAndRemovedCaseInsensitively()
        {
            var service = CreateService();
            string path = Path.Combine(_root, "Assets");

            service.AddSharedFolder(path);
            service.AddSharedFolder(path.ToUpperInvariant());
            Assert.AreEqual(1, _sharedFolders.Count);
            Assert.AreEqual(1, _saveCalls);

            service.RemoveSharedFolder(path.ToUpperInvariant());
            Assert.AreEqual(0, _sharedFolders.Count);
            Assert.AreEqual(2, _saveCalls);
        }

        [TestMethod]
        public void CacheFolderIsDeterministicAndUsesCorrectStorageRoot()
        {
            string source = Path.Combine(_root, "source", "deck.pptx");
            var service = CreateService();
            string central = service.GetCacheFolder(source);

            Assert.AreEqual(central, service.GetCacheFolder(source));
            StringAssert.StartsWith(central, Path.Combine(_root, "central"));

            _sharedFolders.Add(Path.GetDirectoryName(source));
            string adjacent = service.GetCacheFolder(source);
            StringAssert.StartsWith(adjacent, Path.Combine(Path.GetDirectoryName(source), ".oPenEff_cache"));
            Assert.IsNull(service.GetCacheFolder(null));
        }

        [TestMethod]
        public void SourceRegistrationAndTimestampValidationRoundTrip()
        {
            string source = Path.Combine(_root, "source.pptx");
            File.WriteAllText(source, "one");
            var service = CreateService();
            string cache = service.GetCacheFolder(source);

            service.RegisterSourcePath(source, cache);
            Assert.AreEqual(source, service.GetSourcePath(cache));
            Assert.IsFalse(service.IsCacheValid(source));

            service.WriteTimestamp(source);
            Assert.IsTrue(service.IsCacheValid(source));

            File.SetLastWriteTimeUtc(source, File.GetLastWriteTimeUtc(source).AddSeconds(2));
            Assert.IsFalse(service.IsCacheValid(source));
        }

        [TestMethod]
        public void PartialTimestampMarksCacheAndRemainsValid()
        {
            string source = Path.Combine(_root, "partial.pptx");
            File.WriteAllText(source, "source");
            var service = CreateService();

            service.WritePartialTimestamp(source);

            string cache = service.GetCacheFolder(source);
            Assert.IsTrue(service.IsCacheValid(source));
            Assert.IsTrue(File.Exists(Path.Combine(cache, "cache.partial")));
        }

        [TestMethod]
        public void CleanupDeletesOnlyStaleCentralEntries()
        {
            string liveSource = Path.Combine(_root, "live.pptx");
            File.WriteAllText(liveSource, "live");
            var service = CreateService();
            string liveCache = service.GetCacheFolder(liveSource);
            service.RegisterSourcePath(liveSource, liveCache);

            string staleCache = Path.Combine(_root, "central", "stale");
            service.RegisterSourcePath(Path.Combine(_root, "missing.pptx"), staleCache);

            CacheCleanupReport report = service.CleanupStaleCache();

            Assert.AreEqual(1, report.StaleCacheFoldersDeleted);
            Assert.IsTrue(Directory.Exists(liveCache));
            Assert.IsFalse(Directory.Exists(staleCache));
        }

        [TestMethod]
        public void CacheSizeIncludesFilesAndFormatsUnits()
        {
            string file = Path.Combine(_root, "central", "entry", "payload.bin");
            Directory.CreateDirectory(Path.GetDirectoryName(file));
            File.WriteAllBytes(file, new byte[1536]);
            var service = CreateService();

            Assert.AreEqual(1536, service.GetCacheSizeBytes());
            Assert.AreEqual(1.5.ToString("F1", CultureInfo.CurrentCulture) + " KB", service.GetCacheSizeFormatted());
        }

        private AssetCacheService CreateService()
        {
            return new AssetCacheService(
                Path.Combine(_root, "central"),
                () => new List<string>(_sharedFolders),
                folders =>
                {
                    _saveCalls++;
                    _sharedFolders = new List<string>(folders);
                });
        }
    }
}
