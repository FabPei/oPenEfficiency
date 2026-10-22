using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.Models;
using oPenEfficiency.Services;

namespace oPenEfficiency.Command.Tests
{
    /// <summary>
    /// Contract coverage for the built-in "Standard Agenda" sample layout required by
    /// docs/product/feat-009-agenda-manager-2-todo.md ("Required built-in sample agenda
    /// layout"): a conservative, immediately usable business agenda template that is
    /// available on a clean profile and never persisted with user layouts.
    /// </summary>
    [TestClass]
    public sealed class StandardAgendaSampleTests
    {
        private const float SlideWidth = 960f;
        private const float SlideHeight = 540f;
        private const int White = 16777215;

        private string _directory;

        [TestInitialize]
        public void SetUp()
        {
            _directory = Path.Combine(Path.GetTempPath(), "oPenEfficiency.StandardAgendaSampleTests", Guid.NewGuid().ToString("N"));
        }

        [TestCleanup]
        public void TearDown()
        {
            if (Directory.Exists(_directory)) Directory.Delete(_directory, recursive: true);
        }

        [TestMethod]
        public void StandardAgendaMatchesTheDocumentedSampleContract()
        {
            AgendaLayoutConfig layout = new JsonAgendaLayoutStore(_directory).Load().Single(item => item.IsBuiltIn);

            // "16:9 slide; clean white background"
            Assert.AreEqual(SlideWidth, layout.SlideShape.Width);
            Assert.AreEqual(SlideHeight, layout.SlideShape.Height);
            Assert.IsTrue(layout.SlideShape.HasFill);
            Assert.AreEqual(White, layout.SlideShape.FillColorRgb);
            Assert.IsFalse(layout.SlideShape.HasLine);

            // Every assigned role shape stays inside the 16:9 canvas.
            foreach (KeyValuePair<string, ShapeFormatConfig> role in Roles(layout))
            {
                ShapeFormatConfig shape = role.Value;
                Assert.IsNotNull(shape, role.Key + " must be part of the sample");
                Assert.IsTrue(shape.Left >= 0 && shape.Top >= 0, role.Key + " must start on the slide");
                Assert.IsTrue(shape.Left + shape.Width <= SlideWidth, role.Key + " must not overflow the slide width");
                Assert.IsTrue(shape.Top + shape.Height <= SlideHeight, role.Key + " must not overflow the slide height");
            }

            // "simple title/headline at the top"
            Assert.IsTrue(layout.HeadlineShape.Top < layout.ItemTextShape.Top, "the headline sits above the agenda items");
            Assert.IsTrue(layout.HeadlineShape.FontBold);
            Assert.IsTrue(layout.SubHeadlineShape.Top > layout.HeadlineShape.Top);
            Assert.IsTrue(layout.SubHeadlineShape.FontSize < layout.HeadlineShape.FontSize);

            // "vertically stacked numbered agenda items; agenda item text as the dominant element"
            Assert.AreEqual("Vertical", layout.ListOrientation);
            Assert.IsNotNull(layout.ItemNumberingShape);
            Assert.AreEqual(layout.ItemNumberingShape.Top, layout.ItemTextShape.Top, "number and text share the row baseline");
            Assert.IsTrue(layout.ItemNumberingShape.Left < layout.ItemTextShape.Left, "the number column precedes the text");
            Assert.IsTrue(layout.ItemTextShape.Width > layout.ItemNumberingShape.Width);
            Assert.IsTrue(layout.ItemBackgroundShape.Top < layout.ItemTextShape.Top);
            Assert.IsTrue(layout.ItemBackgroundShape.Top + layout.ItemBackgroundShape.Height > layout.ItemTextShape.Top + layout.ItemTextShape.Height,
                "the row background encloses the item text row");

            // "current agenda item highlighted with one restrained accent treatment"
            Assert.IsTrue(layout.HighlightBold);
            Assert.IsTrue(layout.HighlightFontColor);
            Assert.IsFalse(layout.HighlightFillColor, "the sample must not highlight with a fill");
            Assert.IsFalse(layout.HighlightItalic);
            Assert.IsFalse(layout.HighlightUnderline);
            Assert.AreEqual(layout.ItemNumberingShape.FontColorRgb, layout.HighlightFontColorRgb,
                "the single accent treatment is reused, not duplicated with a second colour");

            // "subtle neutral divider/separator treatment"
            Assert.IsTrue(layout.ItemBackgroundShape.HasLine);
            Assert.IsFalse(layout.ItemBackgroundShape.HasFill);
            Assert.IsTrue(layout.ItemBackgroundShape.LineWeight > 0 && layout.ItemBackgroundShape.LineWeight <= 1);
            Assert.IsTrue(layout.SubItemBackgroundShape.HasLine);
            Assert.IsTrue(layout.SubItemBackgroundShape.LineWeight <= layout.ItemBackgroundShape.LineWeight);

            // "optional page number/time/responsible fields ... visually secondary". The
            // built-in intentionally omits page numbers; custom layouts can still define them.
            Assert.IsNull(layout.ItemPageNumberShape);
            Assert.IsNull(layout.SubItemPageNumberShape);
            foreach (ShapeFormatConfig metadata in new[]
            {
                layout.ItemResponsibleShape,
                layout.ItemTimeShape,
                layout.ItemDurationShape
            })
            {
                Assert.IsTrue(metadata.FontSize < layout.ItemTextShape.FontSize, "metadata text stays secondary");
                Assert.AreEqual(layout.ItemResponsibleShape.FontColorRgb, metadata.FontColorRgb);
            }
            Assert.AreNotEqual(layout.ItemTextShape.FontColorRgb, layout.ItemResponsibleShape.FontColorRgb,
                "metadata uses the secondary text colour");
            Assert.AreEqual(1, layout.ItemTimeShape.TextAlignment, "time ranges stay left-to-right");
            Assert.IsTrue(layout.ItemTimeShape.Width >= 120, "time ranges fit without reflowing their endpoints");

            // "sub-items visibly indented below their parent item"
            Assert.IsTrue(layout.SubItemNumberingShape.Left > layout.ItemNumberingShape.Left);
            Assert.IsTrue(layout.SubItemTextShape.Left > layout.ItemTextShape.Left);
            Assert.IsTrue(layout.SubItemTextShape.FontSize < layout.ItemTextShape.FontSize);
            Assert.IsFalse(layout.SubItemTextShape.FontBold, "sub-items are visibly lighter than their parent");
            Assert.IsTrue(layout.SubItemBackgroundShape.Left > layout.ItemBackgroundShape.Left,
                "the sub-item separator is inset against the parent separator");

            // "no decorative graphics, gradients, stock imagery, or brand-specific assets"
            Assert.IsFalse(layout.ItemTextShape.HasFill);
            Assert.IsFalse(layout.HeadlineShape.HasFill);
            Assert.IsFalse(layout.SubItemTextShape.HasFill);

            // "use theme-safe/default Office typography and colors where practical"
            foreach (ShapeFormatConfig shape in Roles(layout).Values)
                Assert.AreEqual("Calibri", shape.FontName);

            // "no decorative ... gradients" - the model has no gradient or image field, but the
            // only filled surface must remain the slide background itself.
            Assert.IsTrue(Roles(layout)
                .Where(role => role.Value.HasFill)
                .All(role => role.Key == "SlideShape"));
        }

        [TestMethod]
        public void StandardAgendaIsNotPersistedAndStaleCopiesDoNotDuplicate()
        {
            var store = new JsonAgendaLayoutStore(_directory);
            Assert.IsTrue(store.TrySave(new List<AgendaLayoutConfig> { CustomLayout("Client Layout", 10f) }));

            string path = Path.Combine(_directory, JsonAgendaLayoutStore.LayoutsConfigFileName);
            string storedFile = File.ReadAllText(path);
            Assert.IsFalse(storedFile.Contains(AgendaLayoutCatalog.StandardAgendaName),
                "the built-in sample must never be materialized into the user layout file");

            // A stale built-in copy written by an older installation must not duplicate the sample.
            string staleJson = "[" + storedFile.Trim().TrimStart('[').TrimEnd(']') +
                ",{\"layoutName\":\"" + AgendaLayoutCatalog.StandardAgendaName + "\",\"itemSpacing\":12}]";
            File.WriteAllText(path, staleJson);

            List<AgendaLayoutConfig> layouts = new JsonAgendaLayoutStore(_directory).Load();
            Assert.AreEqual(2, layouts.Count, "exactly one built-in sample plus the user layout");
            Assert.AreEqual(1, layouts.Count(layout => layout.IsBuiltIn));
            Assert.AreEqual(AgendaLayoutCatalog.StandardAgendaId, layouts.Single(layout => layout.IsBuiltIn).BuiltInId);
            Assert.AreEqual("Client Layout", layouts.Single(layout => !layout.IsBuiltIn).LayoutName);
        }

        [TestMethod]
        public void DefaultSelectionPrefersTheBuiltInSample()
        {
            // Clean profile / empty layout store.
            List<AgendaLayoutConfig> cleanProfile = new JsonAgendaLayoutStore(_directory).Load();
            Assert.AreEqual(0, AgendaLayoutCatalog.GetDefaultSelectionIndex(cleanProfile));
            Assert.IsTrue(cleanProfile[AgendaLayoutCatalog.GetDefaultSelectionIndex(cleanProfile)].IsBuiltIn);

            // Adding personal layouts keeps the sample selectable and still selected.
            Assert.IsTrue(new JsonAgendaLayoutStore(_directory).TrySave(
                new List<AgendaLayoutConfig> { CustomLayout("Client Layout", 10f) }));
            List<AgendaLayoutConfig> withUserLayouts = new JsonAgendaLayoutStore(_directory).Load();
            int selected = AgendaLayoutCatalog.GetDefaultSelectionIndex(withUserLayouts);
            Assert.IsTrue(withUserLayouts[selected].IsBuiltIn);

            // After the user deletes/resets the personal layouts the sample remains.
            Assert.IsTrue(new JsonAgendaLayoutStore(_directory).TrySave(new List<AgendaLayoutConfig>()));
            List<AgendaLayoutConfig> afterReset = new JsonAgendaLayoutStore(_directory).Load();
            Assert.AreEqual(1, afterReset.Count);
            Assert.IsTrue(afterReset[0].IsBuiltIn);

            Assert.AreEqual(-1, AgendaLayoutCatalog.GetDefaultSelectionIndex(new List<AgendaLayoutConfig>()));
            Assert.AreEqual(-1, AgendaLayoutCatalog.GetDefaultSelectionIndex(null));
        }

        [TestMethod]
        public void PlanningUsesMonotonicStartToEndTimeSlots()
        {
            var items = new List<AgendaPlanItem>
            {
                new AgendaPlanItem { Topic = "Opening", Duration = 15, ShouldGenerate = true },
                new AgendaPlanItem { Topic = "Review", Duration = 15, ShouldGenerate = true }
            };

            new AgendaPlanningService().Build(items, TimeSpan.FromHours(9), 1, "HH:mm", null);

            Assert.AreEqual("09:00 - 09:15", items[0].TimeSlot);
            Assert.AreEqual("09:15 - 09:30", items[1].TimeSlot);
        }

        private static Dictionary<string, ShapeFormatConfig> Roles(AgendaLayoutConfig layout) =>
            new Dictionary<string, ShapeFormatConfig>(StringComparer.Ordinal)
            {
                { "SlideShape", layout.SlideShape },
                { "HeadlineShape", layout.HeadlineShape },
                { "SubHeadlineShape", layout.SubHeadlineShape },
                { "ItemNumberingShape", layout.ItemNumberingShape },
                { "ItemTextShape", layout.ItemTextShape },
                { "ItemBackgroundShape", layout.ItemBackgroundShape },
                { "ItemResponsibleShape", layout.ItemResponsibleShape },
                { "ItemTimeShape", layout.ItemTimeShape },
                { "ItemDurationShape", layout.ItemDurationShape },
                { "SubItemNumberingShape", layout.SubItemNumberingShape },
                { "SubItemTextShape", layout.SubItemTextShape },
                { "SubItemBackgroundShape", layout.SubItemBackgroundShape },
                { "SubItemResponsibleShape", layout.SubItemResponsibleShape },
                { "SubItemTimeShape", layout.SubItemTimeShape },
                { "SubItemDurationShape", layout.SubItemDurationShape }
            };

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
