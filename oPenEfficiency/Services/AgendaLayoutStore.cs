using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using oPenEfficiency.Models;

namespace oPenEfficiency.Services
{
    public interface IAgendaLayoutStore
    {
        List<AgendaLayoutConfig> Load();
        bool TrySave(List<AgendaLayoutConfig> layouts);
    }

    public sealed class JsonAgendaLayoutStore : IAgendaLayoutStore
    {
        public const string LayoutsConfigFileName = "agenda_layouts.json";
        private readonly JsonFileStore _jsonStore;

        public JsonAgendaLayoutStore()
            : this(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "oPenEfficiency",
                "AgendaLayouts"))
        {
        }

        public JsonAgendaLayoutStore(string layoutsDirectory)
        {
            _jsonStore = new JsonFileStore(layoutsDirectory);
        }

        public List<AgendaLayoutConfig> Load()
        {
            var userLayouts = _jsonStore.Load<List<AgendaLayoutConfig>>(LayoutsConfigFileName) ??
                              new List<AgendaLayoutConfig>();
            return AgendaLayoutCatalog.WithBuiltInLayouts(userLayouts);
        }

        public bool TrySave(List<AgendaLayoutConfig> layouts)
        {
            if (layouts == null) throw new ArgumentNullException(nameof(layouts));
            if (layouts.Any(layout =>
                    layout != null &&
                    !layout.IsBuiltIn &&
                    AgendaLayoutCatalog.IsBuiltInName(layout.LayoutName)))
                return false;

            var userLayouts = layouts
                .Where(layout => layout != null && !layout.IsBuiltIn)
                .ToList();
            return _jsonStore.TrySave(userLayouts, LayoutsConfigFileName);
        }
    }

    /// <summary>
    /// Owns immutable built-in agenda layouts. Built-ins are merged in memory and
    /// are deliberately never persisted into the user-owned JSON file.
    /// </summary>
    public static class AgendaLayoutCatalog
    {
        public const string StandardAgendaName = "Standard Agenda";
        public const string StandardAgendaId = "builtin.standard-agenda";
        public const int StandardAgendaVersion = 1;

        public static bool IsBuiltInName(string layoutName) =>
            string.Equals(layoutName, StandardAgendaName, StringComparison.OrdinalIgnoreCase);

        public static List<AgendaLayoutConfig> WithBuiltInLayouts(IEnumerable<AgendaLayoutConfig> userLayouts)
        {
            var result = new List<AgendaLayoutConfig> { CreateStandardAgenda() };
            if (userLayouts == null) return result;
            result.AddRange(userLayouts.Where(layout =>
                layout != null &&
                !layout.IsBuiltIn &&
                !IsBuiltInName(layout.LayoutName)));
            return result;
        }

        /// <summary>
        /// Layout that a layout picker selects by default. The built-in sample comes
        /// first so a clean profile can create an agenda without running the Agenda
        /// Layout Designer first. Returns -1 when no layout can be selected.
        /// </summary>
        public static int GetDefaultSelectionIndex(IReadOnlyList<AgendaLayoutConfig> layouts)
        {
            if (layouts == null || layouts.Count == 0) return -1;
            for (int index = 0; index < layouts.Count; index++)
            {
                AgendaLayoutConfig layout = layouts[index];
                if (layout != null && (layout.IsBuiltIn || IsBuiltInName(layout.LayoutName)))
                    return index;
            }
            return 0;
        }

        public static AgendaLayoutConfig CreateStandardAgenda()
        {
            int darkText = OfficeRgb(45, 52, 61);
            int secondaryText = OfficeRgb(99, 107, 118);
            int accent = OfficeRgb(47, 84, 150);
            int separator = OfficeRgb(218, 222, 228);

            return new AgendaLayoutConfig
            {
                LayoutName = StandardAgendaName,
                IsBuiltIn = true,
                BuiltInId = StandardAgendaId,
                BuiltInVersion = StandardAgendaVersion,
                SlideShape = Shape(0, 0, 960, 540, 1, darkText, hasFill: true, fillColor: OfficeRgb(255, 255, 255)),
                HeadlineShape = Shape(80, 42, 800, 46, 30, darkText, bold: true),
                SubHeadlineShape = Shape(80, 94, 800, 26, 15, secondaryText),
                ItemNumberingShape = Shape(90, 166, 45, 34, 18, accent, bold: true, alignment: 2),
                ItemTextShape = Shape(150, 166, 430, 34, 20, darkText, bold: true),
                ItemBackgroundShape = Shape(80, 162, 855, 42, 1, darkText,
                    hasLine: true, lineColor: separator, lineWeight: 0.75f),
                ItemResponsibleShape = Shape(600, 166, 120, 34, 12, secondaryText, alignment: 3),
                ItemTimeShape = Shape(735, 166, 120, 34, 12, secondaryText),
                ItemDurationShape = Shape(865, 166, 70, 34, 12, secondaryText, alignment: 3),
                SubItemNumberingShape = Shape(120, 166, 45, 30, 14, secondaryText, alignment: 2),
                SubItemTextShape = Shape(180, 166, 400, 30, 16, secondaryText),
                SubItemBackgroundShape = Shape(110, 163, 825, 36, 1, darkText,
                    hasLine: true, lineColor: separator, lineWeight: 0.5f),
                SubItemResponsibleShape = Shape(600, 166, 120, 30, 11, secondaryText, alignment: 3),
                SubItemTimeShape = Shape(735, 166, 120, 30, 11, secondaryText),
                SubItemDurationShape = Shape(865, 166, 70, 30, 11, secondaryText, alignment: 3),
                ListAlignment = "Left",
                ListOrientation = "Vertical",
                ItemSpacing = 10f,
                AlignItemsMiddle = true,
                AlignInfoShapesWithItem = true,
                HighlightBold = true,
                HighlightFontColor = true,
                HighlightFontColorRgb = accent,
                HighlightFillColor = false
            };
        }

        private static ShapeFormatConfig Shape(
            float left,
            float top,
            float width,
            float height,
            float fontSize,
            int fontColor,
            bool bold = false,
            int alignment = 1,
            bool hasFill = false,
            int fillColor = 0,
            bool hasLine = false,
            int lineColor = 0,
            float lineWeight = 0)
        {
            return new ShapeFormatConfig
            {
                Left = left,
                Top = top,
                Width = width,
                Height = height,
                HasFill = hasFill,
                FillColorRgb = fillColor,
                HasLine = hasLine,
                LineColorRgb = lineColor,
                LineWeight = lineWeight,
                FontName = "Calibri",
                FontSize = fontSize,
                FontColorRgb = fontColor,
                FontBold = bold,
                TextAlignment = alignment,
                MarginLeft = 4,
                MarginRight = 4,
                MarginTop = 2,
                MarginBottom = 2,
                Orientation = 1
            };
        }

        private static int OfficeRgb(byte red, byte green, byte blue) =>
            red | (green << 8) | (blue << 16);
    }
}
