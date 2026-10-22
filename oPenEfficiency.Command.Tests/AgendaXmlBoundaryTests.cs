using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.Command.Tests
{
    [TestClass]
    public sealed class AgendaXmlBoundaryTests
    {
        private string _directory;

        [TestInitialize]
        public void SetUp()
        {
            _directory = Path.Combine(Path.GetTempPath(), "oPE-agenda-xml", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_directory);
        }

        [TestCleanup]
        public void TearDown()
        {
            if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
        }

        [TestMethod]
        public void MalformedAndEmptyXmlNeverPersistPartialState()
        {
            var store = new RecordingLayoutStore();
            string malformed = Write("AgendaCreator.xml", "<layouts><layout name='Broken'></layouts>");
            string empty = Write("AgendaWizard.xml", "<root><unrelated /></root>");

            Assert.AreEqual(0, new ExternalAgendaImportService(store).ImportFromXml(malformed).Count);
            Assert.AreEqual(0, new ExternalAgendaImportService(store).ImportFromXml(empty).Count);
            Assert.AreEqual(0, store.SaveCalls);
            Assert.AreEqual(1, store.LoadCalls, "Malformed XML should fail before persisted state is loaded.");
        }

        [TestMethod]
        public void AgendaXmlProhibitsDtdAndExternalEntities()
        {
            string externalPath = Path.Combine(_directory, "secret.txt");
            File.WriteAllText(externalPath, "must-not-be-read");
            string xml = "<!DOCTYPE layouts [<!ENTITY xxe SYSTEM '" +
                new Uri(externalPath).AbsoluteUri +
                "'>]><layouts><layout name='&xxe;'/></layouts>";
            var store = new RecordingLayoutStore();

            var imported = new ExternalAgendaImportService(store).ImportFromXml(Write("AgendaCreator.xml", xml));

            Assert.AreEqual(0, imported.Count);
            Assert.AreEqual(0, store.LoadCalls);
            Assert.AreEqual(0, store.SaveCalls);
        }

        [TestMethod]
        public void OversizedAgendaXmlIsRejectedBeforePersistence()
        {
            string oversizedName = new string('x', 5 * 1024 * 1024 + 1);
            var store = new RecordingLayoutStore();
            string path = Write(
                "AgendaCreator.xml",
                "<layouts><layout name='" + oversizedName + "'/></layouts>");

            var imported = new ExternalAgendaImportService(store).ImportFromXml(path);

            Assert.AreEqual(0, imported.Count);
            Assert.AreEqual(0, store.LoadCalls);
            Assert.AreEqual(0, store.SaveCalls);
        }

        [TestMethod]
        public void ImportedGeometryClampsNegativeOriginsAndUsesFallbackName()
        {
            var store = new RecordingLayoutStore();
            string xml = "<layouts><layout><position left='-20' top='-30'/>" +
                "<cases><case level='1' selected='1'/></cases></layout></layouts>";

            var imported = new ExternalAgendaImportService(store)
                .ImportFromXml(Write("AgendaCreator.xml", xml));

            CollectionAssert.AreEqual(new[] { "Imported Layout" }, imported);
            AgendaLayoutConfig layout = store.Saved[0];
            Assert.AreEqual(0, layout.ItemTextShape.Left);
            Assert.AreEqual(0, layout.ItemTextShape.Top);
            Assert.IsTrue(layout.ItemTextShape.Width > 0 && layout.ItemTextShape.Height > 0);
        }

        [TestMethod]
        public void DuplicateNamesWithinOneXmlAreMergedCaseInsensitively()
        {
            var store = new RecordingLayoutStore();
            string xml = "<layouts><layout name='Layout A'/><layout name='layout a'/>" +
                "<layout name='Layout B'/></layouts>";

            var imported = new ExternalAgendaImportService(store)
                .ImportFromXml(Write("AgendaCreator.xml", xml));

            CollectionAssert.AreEqual(new[] { "Layout A", "Layout B" }, imported);
            Assert.AreEqual(2, store.Saved.Count);
            Assert.AreEqual(1, store.SaveCalls);
        }

        private string Write(string name, string contents)
        {
            string path = Path.Combine(_directory, name);
            File.WriteAllText(path, contents);
            return path;
        }

        private sealed class RecordingLayoutStore : IAgendaLayoutStore
        {
            public int LoadCalls { get; private set; }
            public int SaveCalls { get; private set; }
            public List<AgendaLayoutConfig> Saved { get; private set; } = new List<AgendaLayoutConfig>();

            public List<AgendaLayoutConfig> Load()
            {
                LoadCalls++;
                return new List<AgendaLayoutConfig>();
            }

            public bool TrySave(List<AgendaLayoutConfig> layouts)
            {
                SaveCalls++;
                Saved = new List<AgendaLayoutConfig>(layouts);
                return true;
            }
        }
    }
}
