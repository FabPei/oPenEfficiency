using System;
using System.IO;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;

namespace oPenEfficiency.Command.Tests
{
    [TestClass]
    public sealed class SecureSidebarConfigStoreTests
    {
        private string _directory;

        [TestInitialize]
        public void SetUp()
        {
            _directory = Path.Combine(Path.GetTempPath(), "oPE-secure-config", Guid.NewGuid().ToString("N"));
        }

        [TestCleanup]
        public void TearDown()
        {
            if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
        }

        [TestMethod]
        public void LegacyPlaintextCredentialsAreMigratedAndRemoved()
        {
            var jsonStore = new JsonFileStore(_directory);
            var legacy = SampleConfig();
            Assert.IsTrue(jsonStore.TrySave(legacy, "sidebar_config.json"));
            StringAssert.Contains(
                File.ReadAllText(Path.Combine(_directory, "sidebar_config.json")),
                "deepl-test-secret");

            var store = CreateStore(jsonStore);
            SidebarConfig loaded = store.Load();

            Assert.AreEqual("deepl-test-secret", loaded.DeepLApiKey);
            Assert.AreEqual("openai-test-secret", loaded.OpenAIApiKey);
            Assert.AreEqual("claude-test-secret", loaded.ClaudeApiKey);
            Assert.AreEqual("libre-test-secret", loaded.LibreTranslateApiKey);

            string persistedJson = File.ReadAllText(Path.Combine(_directory, "sidebar_config.json"));
            Assert.IsFalse(persistedJson.Contains("test-secret"));
            Assert.IsFalse(persistedJson.Contains("DeepLApiKey"));
            Assert.IsFalse(persistedJson.Contains("OpenAIApiKey"));
            Assert.IsFalse(persistedJson.Contains("ClaudeApiKey"));
            Assert.IsFalse(persistedJson.Contains("LibreTranslateApiKey"));

            string credentialPath = Path.Combine(_directory, DpapiCredentialBlobStore.DefaultFileName);
            Assert.IsTrue(File.Exists(credentialPath));
            Assert.IsFalse(Encoding.UTF8.GetString(File.ReadAllBytes(credentialPath)).Contains("test-secret"));

            SidebarConfig reloaded = CreateStore(jsonStore).Load();
            Assert.AreEqual("deepl-test-secret", reloaded.DeepLApiKey);
            Assert.AreEqual("openai-test-secret", reloaded.OpenAIApiKey);
            Assert.AreEqual("claude-test-secret", reloaded.ClaudeApiKey);
            Assert.AreEqual("libre-test-secret", reloaded.LibreTranslateApiKey);
        }

        [TestMethod]
        public void NormalSaveProtectsCredentialsWhilePortableSerializationRemainsExplicit()
        {
            var jsonStore = new JsonFileStore(_directory);
            var store = CreateStore(jsonStore);
            SidebarConfig config = SampleConfig();

            Assert.IsTrue(store.TrySave(config));
            string persistedJson = File.ReadAllText(Path.Combine(_directory, "sidebar_config.json"));
            Assert.IsFalse(persistedJson.Contains("test-secret"));

            string portableJson = JsonService.Serialize(config);
            StringAssert.Contains(portableJson, "deepl-test-secret");
            StringAssert.Contains(portableJson, "openai-test-secret");
        }

        [TestMethod]
        public void FailedJsonSaveRestoresThePreviousCredentialBlob()
        {
            Directory.CreateDirectory(_directory);
            Directory.CreateDirectory(Path.Combine(_directory, "sidebar_config.json"));
            var jsonStore = new JsonFileStore(_directory);
            var store = CreateStore(jsonStore);

            Assert.IsFalse(store.TrySave(SampleConfig()));
            Assert.IsFalse(File.Exists(Path.Combine(_directory, DpapiCredentialBlobStore.DefaultFileName)));
        }

        [TestMethod]
        public void ExistingProtectedCredentialsWinOverStaleLegacyValues()
        {
            var jsonStore = new JsonFileStore(_directory);
            var protectedStore = new DpapiCredentialBlobStore(_directory);
            var store = new SecureSidebarConfigStore(jsonStore, protectedStore);
            var protectedConfig = SampleConfig();
            protectedConfig.DeepLApiKey = "protected-current-value";
            Assert.IsTrue(store.TrySave(protectedConfig));

            var staleLegacy = SampleConfig();
            staleLegacy.DeepLApiKey = "stale-plaintext-value";
            Assert.IsTrue(jsonStore.TrySave(staleLegacy, "sidebar_config.json"));

            SidebarConfig loaded = store.Load();
            Assert.AreEqual("protected-current-value", loaded.DeepLApiKey);
            Assert.IsFalse(File.ReadAllText(Path.Combine(_directory, "sidebar_config.json"))
                .Contains("stale-plaintext-value"));
        }

        [TestMethod]
        public void FailedProtectedMigrationLeavesTheOnlyPlaintextCopyUntouched()
        {
            var jsonStore = new JsonFileStore(_directory);
            Assert.IsTrue(jsonStore.TrySave(SampleConfig(), "sidebar_config.json"));
            var store = new SecureSidebarConfigStore(jsonStore, new FailingCredentialBlobStore());

            SidebarConfig loaded = store.Load();

            Assert.AreEqual("deepl-test-secret", loaded.DeepLApiKey);
            StringAssert.Contains(
                File.ReadAllText(Path.Combine(_directory, "sidebar_config.json")),
                "deepl-test-secret");
        }

        [TestMethod]
        public void CorruptProtectedBlobCanBeRepairedFromLegacyCredentials()
        {
            Directory.CreateDirectory(_directory);
            var jsonStore = new JsonFileStore(_directory);
            Assert.IsTrue(jsonStore.TrySave(SampleConfig(), "sidebar_config.json"));
            File.WriteAllBytes(
                Path.Combine(_directory, DpapiCredentialBlobStore.DefaultFileName),
                new byte[] { 1, 2, 3, 4 });

            SidebarConfig loaded = CreateStore(jsonStore).Load();

            Assert.AreEqual("deepl-test-secret", loaded.DeepLApiKey);
            Assert.IsFalse(File.ReadAllText(Path.Combine(_directory, "sidebar_config.json"))
                .Contains("deepl-test-secret"));
            Assert.AreEqual("deepl-test-secret", CreateStore(jsonStore).Load().DeepLApiKey);
        }

        private SecureSidebarConfigStore CreateStore(JsonFileStore jsonStore)
        {
            return new SecureSidebarConfigStore(
                jsonStore,
                new DpapiCredentialBlobStore(_directory));
        }

        private static SidebarConfig SampleConfig()
        {
            return new SidebarConfig
            {
                BackgroundColor = "Transparent",
                DeepLApiKey = "deepl-test-secret",
                OpenAIApiKey = "openai-test-secret",
                ClaudeApiKey = "claude-test-secret",
                LibreTranslateApiKey = "libre-test-secret"
            };
        }

        private sealed class FailingCredentialBlobStore : ICredentialBlobStore
        {
            public CredentialReadState TryRead(out SidebarCredentials credentials)
            {
                credentials = null;
                return CredentialReadState.Missing;
            }

            public CredentialBlobSnapshot Capture()
            {
                return new CredentialBlobSnapshot { Captured = true };
            }

            public bool TryWrite(SidebarCredentials credentials)
            {
                return false;
            }

            public bool TryRestore(CredentialBlobSnapshot snapshot)
            {
                return true;
            }
        }
    }
}
