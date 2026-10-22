using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class SidebarMigrationBoundaryTests
    {
        [TestMethod]
        public void NewSidebarConfigTracksThemeBackgroundByDefault()
        {
            var config = new SidebarConfig();

            Assert.AreEqual("Transparent", config.BackgroundColor);
        }

        [TestMethod]
        public void ExplicitCustomBackgroundModeDoesNotDependOnColorLiteral()
        {
            var config = new SidebarConfig
            {
                BackgroundColor = "#292929",
                BackgroundMode = SidebarBackgroundModes.Custom
            };

            Assert.IsFalse(SidebarBackgroundModes.UsesTheme(config));
        }

        [TestMethod]
        public void LegacyCustomBackgroundIsMigratedToExplicitCustomMode()
        {
            var config = new SidebarConfig
            {
                BackgroundColor = "#123456",
                BackgroundMode = null
            };

            bool changed = SidebarBackgroundModes.Normalize(config);

            Assert.IsTrue(changed);
            Assert.AreEqual(SidebarBackgroundModes.Custom, config.BackgroundMode);
            Assert.IsFalse(SidebarBackgroundModes.UsesTheme(config));
        }

        [TestMethod]
        public void LegacySerializedCustomBackgroundKeepsMissingModeForMigration()
        {
            const string legacyJson = "{\"BackgroundColor\":\"#123456\"}";
            var serializer = new DataContractJsonSerializer(typeof(SidebarConfig));
            SidebarConfig config;
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(legacyJson)))
            {
                config = (SidebarConfig)serializer.ReadObject(stream);
            }

            Assert.IsNull(config.BackgroundMode,
                "Legacy JSON must retain a missing mode so color-based migration can distinguish it from an explicit theme choice.");
            Assert.IsTrue(SidebarBackgroundModes.Normalize(config));
            Assert.AreEqual(SidebarBackgroundModes.Custom, config.BackgroundMode);
            Assert.AreEqual("#123456", config.BackgroundColor);
        }

        [TestMethod]
        public void LegacyDefaultBackgroundIsMigratedToThemeMode()
        {
            var config = new SidebarConfig
            {
                BackgroundColor = "#252525",
                BackgroundMode = null
            };

            bool changed = SidebarBackgroundModes.Normalize(config);

            Assert.IsTrue(changed);
            Assert.AreEqual(SidebarBackgroundModes.Theme, config.BackgroundMode);
            Assert.AreEqual("Transparent", config.BackgroundColor);
        }

        [TestMethod]
        public void NewSidebarConfigTracksThemeSectionHeaderByDefault()
        {
            var config = new SidebarConfig();

            Assert.IsTrue(SidebarSectionFontColorModes.UsesTheme(config));
            Assert.AreEqual(SidebarSectionFontColorModes.Theme, config.SectionFontColorMode);
        }

        [TestMethod]
        public void LegacyDefaultSectionHeaderIsMigratedToThemeMode()
        {
            var config = new SidebarConfig
            {
                SectionFontColor = "#9CA3AF",
                SectionFontColorMode = null
            };

            bool changed = SidebarSectionFontColorModes.Normalize(config);

            Assert.IsTrue(changed);
            Assert.AreEqual(SidebarSectionFontColorModes.Theme, config.SectionFontColorMode);
        }

        [TestMethod]
        public void LegacyCustomSectionHeaderIsPreservedAsCustom()
        {
            var config = new SidebarConfig
            {
                SectionFontColor = "#123456",
                SectionFontColorMode = null
            };

            bool changed = SidebarSectionFontColorModes.Normalize(config);

            Assert.IsTrue(changed);
            Assert.AreEqual(SidebarSectionFontColorModes.Custom, config.SectionFontColorMode);
            Assert.AreEqual("#123456", config.SectionFontColor);
        }

        [TestMethod]
        public void ThemeSectionHeaderModeRetainsSavedCustomColor()
        {
            var config = new SidebarConfig
            {
                SectionFontColor = "#123456",
                SectionFontColorMode = SidebarSectionFontColorModes.Theme
            };

            bool changed = SidebarSectionFontColorModes.Normalize(config);

            Assert.IsFalse(changed);
            Assert.AreEqual("#123456", config.SectionFontColor);
        }

        [TestMethod]
        public void MigrationMapsLegacyIdsRemovesObsoleteEntriesAndReportsUnknowns()
        {
            var config = Config(
                new SidebarFeature { Id = "BtnAlignToTable", Name = "Old name" },
                new SidebarFeature { Id = "BtnSwapHorizontal", Name = "Obsolete" },
                new SidebarFeature { Id = "UnknownZ", Name = "Unknown" },
                new SidebarFeature { Id = "unknownA", Name = "Unknown" },
                null);
            var catalog = new[]
            {
                new SidebarFeature { Id = "BtnAlignToTableCell", Name = "Align to Table Cell" }
            };

            SidebarLayoutMigrationResult result = SidebarLayoutMigration.Migrate(config, catalog);

            Assert.IsTrue(result.Changed);
            Assert.AreEqual(2, result.RemovedObsoleteCount);
            Assert.AreEqual(1, result.RefreshedNameCount);
            CollectionAssert.AreEqual(new[] { "unknownA", "UnknownZ" }, result.UnknownFeatureIds.ToArray());
            Assert.AreEqual("BtnAlignToTableCell", config.Sections[0].Features[0].Id);
            Assert.AreEqual("Align to Table Cell", config.Sections[0].Features[0].Name);
        }

        [TestMethod]
        public void MigrationIsIdempotentAfterFirstSuccessfulPass()
        {
            var config = Config(new SidebarFeature { Id = "btnaligntotablecell", Name = "Old" });
            var catalog = new[]
            {
                new SidebarFeature { Id = "BtnAlignToTableCell", Name = "Align to Table Cell" }
            };

            SidebarLayoutMigrationResult first = SidebarLayoutMigration.Migrate(config, catalog);
            SidebarLayoutMigrationResult second = SidebarLayoutMigration.Migrate(config, catalog);

            Assert.IsTrue(first.Changed);
            Assert.IsFalse(second.Changed);
            Assert.AreEqual(0, second.RefreshedNameCount);
            Assert.AreEqual(0, second.RemovedObsoleteCount);
        }

        [TestMethod]
        public void MigrationHandlesNullStructuresWithoutThrowing()
        {
            Assert.IsFalse(SidebarLayoutMigration.Migrate(null, null).Changed);
            Assert.IsFalse(SidebarLayoutMigration.Migrate(new SidebarConfig { Sections = null }, null).Changed);

            var config = new SidebarConfig
            {
                Sections = new List<SidebarSection>
                {
                    null,
                    new SidebarSection { Name = "Empty", Features = null }
                }
            };
            Assert.IsFalse(SidebarLayoutMigration.Migrate(config, null).Changed);
        }

        private static SidebarConfig Config(params SidebarFeature[] features)
        {
            return new SidebarConfig
            {
                Sections = new List<SidebarSection>
                {
                    new SidebarSection { Name = "Test", Features = features.ToList() }
                }
            };
        }
    }
}
