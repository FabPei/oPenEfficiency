using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.Command.Tests
{
    [TestClass]
    public sealed class AgendaLayoutStoreTests
    {
        private string _directory;

        [TestInitialize]
        public void SetUp()
        {
            _directory = Path.Combine(Path.GetTempPath(), "oPenEfficiency.AgendaLayoutStoreTests", Guid.NewGuid().ToString("N"));
        }

        [TestCleanup]
        public void TearDown()
        {
            if (Directory.Exists(_directory)) Directory.Delete(_directory, recursive: true);
        }

        [TestMethod]
        public void EmptyStoreProvidesOneUsableStandardAgenda()
        {
            List<AgendaLayoutConfig> layouts = new JsonAgendaLayoutStore(_directory).Load();

            Assert.AreEqual(1, layouts.Count);
            AgendaLayoutConfig standard = layouts.Single();
            Assert.AreEqual(AgendaLayoutCatalog.StandardAgendaName, standard.LayoutName);
            Assert.IsTrue(standard.IsBuiltIn);
            Assert.AreEqual(AgendaLayoutCatalog.StandardAgendaId, standard.BuiltInId);
            Assert.AreEqual(AgendaLayoutCatalog.StandardAgendaVersion, standard.BuiltInVersion);
            Assert.AreEqual(960f, standard.SlideShape.Width);
            Assert.AreEqual(540f, standard.SlideShape.Height);
            Assert.IsNotNull(standard.HeadlineShape);
            Assert.IsNotNull(standard.ItemTextShape);
            Assert.IsNotNull(standard.SubItemTextShape);
            Assert.IsTrue(standard.SubItemTextShape.Left > standard.ItemTextShape.Left);
            Assert.IsTrue(standard.HighlightBold);
            Assert.IsTrue(standard.HighlightFontColor);
        }

        [TestMethod]
        public void SaveOverwriteAndReopenRoundTripsUserLayoutInAgendaDirectory()
        {
            var store = new JsonAgendaLayoutStore(_directory);
            List<AgendaLayoutConfig> layouts = store.Load();
            layouts.Add(CustomLayout("Client Layout", 12f));

            Assert.IsTrue(store.TrySave(layouts));
            string expectedPath = Path.Combine(_directory, JsonAgendaLayoutStore.LayoutsConfigFileName);
            Assert.IsTrue(File.Exists(expectedPath));
            CollectionAssert.AreEquivalent(
                new[] { JsonAgendaLayoutStore.LayoutsConfigFileName },
                Directory.GetFiles(_directory).Select(Path.GetFileName).ToArray());

            var reopenedStore = new JsonAgendaLayoutStore(_directory);
            List<AgendaLayoutConfig> reopened = reopenedStore.Load();
            Assert.AreEqual(1, reopened.Count(layout => layout.IsBuiltIn));
            Assert.AreEqual(12f, reopened.Single(layout => layout.LayoutName == "Client Layout").ItemSpacing);

            reopened.Single(layout => layout.LayoutName == "Client Layout").ItemSpacing = 24f;
            Assert.IsTrue(reopenedStore.TrySave(reopened));
            List<AgendaLayoutConfig> overwritten = new JsonAgendaLayoutStore(_directory).Load();
            Assert.AreEqual(24f, overwritten.Single(layout => layout.LayoutName == "Client Layout").ItemSpacing);
            Assert.AreEqual(1, overwritten.Count(layout => layout.IsBuiltIn));
        }

        [TestMethod]
        public void ResettingUserLayoutsRestoresExactlyOneBuiltInLayout()
        {
            var store = new JsonAgendaLayoutStore(_directory);
            Assert.IsTrue(store.TrySave(new List<AgendaLayoutConfig> { CustomLayout("Temporary", 10f) }));
            Assert.IsTrue(store.TrySave(new List<AgendaLayoutConfig>()));

            List<AgendaLayoutConfig> layouts = new JsonAgendaLayoutStore(_directory).Load();
            Assert.AreEqual(1, layouts.Count);
            Assert.IsTrue(layouts[0].IsBuiltIn);
            Assert.AreEqual(AgendaLayoutCatalog.StandardAgendaName, layouts[0].LayoutName);
        }

        [TestMethod]
        public void UserLayoutCannotOverwriteBuiltInIdentity()
        {
            var store = new JsonAgendaLayoutStore(_directory);
            Assert.IsFalse(store.TrySave(new List<AgendaLayoutConfig>
            {
                CustomLayout(AgendaLayoutCatalog.StandardAgendaName, 99f)
            }));
            Assert.IsFalse(File.Exists(Path.Combine(_directory, JsonAgendaLayoutStore.LayoutsConfigFileName)));
        }

        private static AgendaLayoutConfig CustomLayout(string name, float spacing) =>
            new AgendaLayoutConfig
            {
                LayoutName = name,
                ItemSpacing = spacing,
                ItemTextShape = new ShapeFormatConfig
                {
                    Left = 100,
                    Top = 100,
                    Width = 400,
                    Height = 30,
                    FontName = "Calibri",
                    FontSize = 18
                }
            };
    }
}
