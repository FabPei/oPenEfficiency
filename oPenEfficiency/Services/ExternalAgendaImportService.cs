using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using oPenEfficiency.Models;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Service for importing external Agenda Creator configurations from other tools' XML format.
    /// Parses XML files like AgendaCreator.xml and converts them to internal AgendaLayoutConfig format.
    /// </summary>
    public class ExternalAgendaImportService
    {
        private const long MaximumXmlCharacters = 5 * 1024 * 1024;
        private const float DefaultImportedItemWidth = 360f;
        private const float DefaultImportedItemHeight = 28f;
        private readonly IAgendaLayoutStore _layoutStore;

        public ExternalAgendaImportService()
            : this(new JsonAgendaLayoutStore())
        {
        }

        public ExternalAgendaImportService(IAgendaLayoutStore layoutStore)
        {
            _layoutStore = layoutStore ?? throw new ArgumentNullException(nameof(layoutStore));
        }

        /// <summary>
        /// Imports an AgendaCreator.xml file and saves the configuration to oPenEfficiency's format.
        /// </summary>
        /// <param name="xmlFilePath">Path to the external AgendaCreator.xml file</param>
        /// <returns>List of imported layout names, or empty list if failed</returns>
        public List<string> ImportFromXml(string xmlFilePath)
        {
            var importedLayouts = new List<string>();

            try
            {
                if (!File.Exists(xmlFilePath))
                {
                    System.Diagnostics.Debug.WriteLine($"ExternalAgendaImportService: File not found: {xmlFilePath}");
                    return importedLayouts;
                }

                var doc = new XmlDocument { XmlResolver = null };
                var readerSettings = new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    MaxCharactersInDocument = MaximumXmlCharacters
                };
                using (var reader = XmlReader.Create(xmlFilePath, readerSettings))
                    doc.Load(reader);

                // Load existing layouts
                var existingLayouts = _layoutStore.Load() ?? new List<AgendaLayoutConfig>();

                // Parse layouts from XML
                var layoutNodes = doc.SelectNodes("//layouts/layout");
                if (layoutNodes == null || layoutNodes.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("ExternalAgendaImportService: No layout nodes found in XML");
                    return importedLayouts;
                }

                foreach (XmlNode layoutNode in layoutNodes)
                {
                    try
                    {
                        var layout = ParseLayout(layoutNode);
                        if (layout != null)
                        {
                            // Check for duplicate and skip
                            if (existingLayouts.Any(l => string.Equals(
                                l.LayoutName,
                                layout.LayoutName,
                                StringComparison.OrdinalIgnoreCase)))
                            {
                                System.Diagnostics.Debug.WriteLine($"Skipping duplicate layout: {layout.LayoutName}");
                                continue;
                            }

                            existingLayouts.Add(layout);
                            importedLayouts.Add(layout.LayoutName);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error parsing layout: {ex.Message}");
                    }
                }

                if (importedLayouts.Count == 0) return importedLayouts;

                // Save all layouts atomically and report success only after persistence succeeded.
                if (!_layoutStore.TrySave(existingLayouts))
                {
                    System.Diagnostics.Debug.WriteLine("ExternalAgendaImportService: Imported layouts could not be saved");
                    importedLayouts.Clear();
                    return importedLayouts;
                }

                System.Diagnostics.Debug.WriteLine($"ExternalAgendaImportService: Imported {importedLayouts.Count} layouts successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExternalAgendaImportService: Import failed: {ex.Message}");
            }

            return importedLayouts;
        }

        /// <summary>
        /// Imports from multiple XML files in a folder.
        /// </summary>
        /// <param name="folderPath">Path containing AgendaCreator.xml or legacy AgendaWizard.xml files</param>
        /// <returns>Total number of layouts imported</returns>
        public List<string> ImportFromFolder(string folderPath)
        {
            var allImported = new List<string>();

            try
            {
                if (!Directory.Exists(folderPath))
                {
                    System.Diagnostics.Debug.WriteLine($"ExternalAgendaImportService: Folder not found: {folderPath}");
                    return allImported;
                }

                var xmlFiles = FindSupportedXmlFiles(folderPath);
                foreach (var xmlFile in xmlFiles)
                {
                    var imported = ImportFromXml(xmlFile);
                    allImported.AddRange(imported);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExternalAgendaImportService: Folder import failed: {ex.Message}");
            }

            return allImported;
        }

        /// <summary>
        /// Gets a list of all available external AgendaCreator.xml files.
        /// </summary>
        public List<string> FindAgendaCreatorFiles()
        {
            var files = new List<string>();

            try
            {
                // Common installation paths for other tools with Agenda Creator
                var searchPaths = new List<string>();

                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                string programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

                // Search in specific common locations where other PowerPoint tools might be installed
                // We search only top-level and specific known subfolders to avoid UnauthorizedAccessException in system folders
                var potentialRoots = new List<string>();
                potentialRoots.Add(Path.Combine(appData));
                potentialRoots.Add(Path.Combine(localAppData));
                potentialRoots.Add(Path.Combine(programFiles));
                potentialRoots.Add(Path.Combine(programFilesX86));

                foreach (var root in potentialRoots)
                {
                    if (Directory.Exists(root))
                    {
                        try
                        {
                            // Get first-level subdirectories only to avoid deep recursive access into protected system folders
                            var subDirs = Directory.GetDirectories(root, "*", SearchOption.TopDirectoryOnly);
                            foreach (var subDir in subDirs)
                            {
                                try
                                {
                                    var foundFiles = FindSupportedXmlFiles(subDir);
                                    files.AddRange(foundFiles);
                                }
                                catch (UnauthorizedAccessException) { /* Skip protected subfolders */ }
                                catch (Exception) { /* Skip other errors */ }
                            }
                        }
                        catch (UnauthorizedAccessException) { /* Skip protected roots */ }
                        catch (Exception) { /* Skip other errors */ }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExternalAgendaImportService: Find failed: {ex.Message}");
            }

            return files.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static IEnumerable<string> FindSupportedXmlFiles(string root)
        {
            return Directory
                .EnumerateFiles(root, "AgendaCreator.xml", SearchOption.AllDirectories)
                .Concat(Directory.EnumerateFiles(root, "AgendaWizard.xml", SearchOption.AllDirectories))
                .Where(AgendaWorkflowPolicy.IsSupportedExternalLayoutFile);
        }

        #region Private Parsing Methods

        private AgendaLayoutConfig ParseLayout(XmlNode layoutNode)
        {
            var config = new AgendaLayoutConfig();

            // Get layout name and ID
            config.LayoutName = GetAttributeValue(layoutNode, "name", "Imported Layout");
            string layoutId = GetAttributeValue(layoutNode, "id", "");

            // Parse standard formatting (default for all elements)
            var standardNode = layoutNode.SelectSingleNode("standard");
            var standardFont = ParseFont(standardNode?.SelectSingleNode("font"));
            var standardParagraph = standardNode?.SelectSingleNode("paragraphformat");

            // Parse agenda settings
            var agendaNode = layoutNode.SelectSingleNode("agenda");
            if (agendaNode != null)
            {
                // Parse agenda-specific settings for reference
                string fontSizeStr = GetAttributeValue(agendaNode, "fontSize", "16");
                string title = GetAttributeValue(agendaNode, "title", "Agenda");
                string subtitle = GetAttributeValue(agendaNode, "subtitle", "");
            }

            // Parse columns configuration
            var columnsNode = layoutNode.SelectSingleNode("columns");
            var columns = new Dictionary<string, XmlNode>();
            if (columnsNode != null)
            {
                foreach (XmlNode colNode in columnsNode.SelectNodes("column"))
                {
                    string field = GetAttributeValue(colNode, "field", "");
                    if (!string.IsNullOrEmpty(field))
                    {
                        columns[field] = colNode;
                    }
                }
            }

            // Parse position
            var positionNode = layoutNode.SelectSingleNode("position");
            float leftMargin = 0, topMargin = 0;
            if (positionNode != null)
            {
                leftMargin = Math.Max(0, GetFloatAttribute(positionNode, "left", 0));
                topMargin = Math.Max(0, GetFloatAttribute(positionNode, "top", 0));
            }

            // Parse settings
            var settingsNode = layoutNode.SelectSingleNode("settings");
            if (settingsNode != null)
            {
                string slideLayout = GetAttributeValue(settingsNode, "slideLayout", "");
                string customLayoutName = GetAttributeValue(settingsNode, "customLayoutName", "");
            }

            // Parse cases (formatting for different levels and states)
            var casesNode = layoutNode.SelectSingleNode("cases");
            var caseNodes = casesNode?.SelectNodes("case");

            var level1Selected = caseNodes == null ? null : FindSelectedCase(caseNodes, 1);
            var level2Selected = caseNodes == null ? null : FindSelectedCase(caseNodes, 2);

            // Parse title shape format (from standard) - map to HeadlineShape for backwards compatibility
            config.HeadlineShape = new ShapeFormatConfig
            {
                Left = leftMargin,
                Top = Math.Max(0, topMargin - 36f),
                Width = DefaultImportedItemWidth,
                Height = 30f,
                FontName = standardFont?.Name ?? "Segoe UI",
                FontSize = standardFont?.Size ?? 16,
                FontColorRgb = standardFont?.Color ?? 0x000000,
                FontBold = standardFont?.Bold ?? false,
                FontItalic = standardFont?.Italic ?? false,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                MarginBottom = 0,
                TextAlignment = GetParagraphAlignment(standardParagraph)
            };

            config.ItemTextShape = ParseItemFormat(
                level1Selected,
                standardFont,
                standardParagraph,
                leftMargin,
                topMargin,
                false);

            config.SubItemTextShape = ParseItemFormat(
                level2Selected,
                standardFont,
                standardParagraph,
                leftMargin,
                topMargin,
                true);

            // Default alignments
            config.ListAlignment = "Top";
            config.ListOrientation = "Vertical";
            config.ItemSpacing = 15f;
            config.HighlightBold = true;

            return config;
        }

        private ShapeFormatConfig ParseItemFormat(
            XmlNode caseNode,
            FontInfo standardFont,
            XmlNode standardParagraph,
            float leftMargin,
            float topMargin,
            bool isSubItem)
        {
            // External files describe formatting but do not necessarily contain PowerPoint
            // shape bounds. Use a conservative, usable fallback anchored at their position.
            var format = new ShapeFormatConfig
            {
                Left = leftMargin + (isSubItem ? 20f : 0f),
                Top = topMargin,
                Width = DefaultImportedItemWidth - (isSubItem ? 20f : 0f),
                Height = DefaultImportedItemHeight
            };

            // Start with standard font settings
            format.FontName = standardFont?.Name ?? "Segoe UI";
            format.FontSize = standardFont?.Size ?? 14;
            format.FontColorRgb = standardFont?.Color ?? 0x000000;
            format.FontBold = standardFont?.Bold ?? false;
            format.FontItalic = standardFont?.Italic ?? false;

            // Parse elements within the case
            var elements = caseNode?.SelectNodes("element");
            if (elements != null)
            {
                foreach (XmlNode elemNode in elements)
                {
                    string type = GetAttributeValue(elemNode, "type", "");
                    string field = GetAttributeValue(elemNode, "field", "");

                    // Parse font overrides
                    var fontNode = elemNode.SelectSingleNode("font");
                    if (fontNode != null)
                    {
                        if (fontNode.Attributes["name"] != null)
                            format.FontName = fontNode.Attributes["name"].Value;
                        if (fontNode.Attributes["bold"] != null && fontNode.Attributes["bold"].Value == "1")
                            format.FontBold = true;
                        if (fontNode.Attributes["italic"] != null && fontNode.Attributes["italic"].Value == "1")
                            format.FontItalic = true;
                        if (fontNode.Attributes["color"] != null)
                        {
                            // Color is stored as theme color index, convert to RGB later
                            format.FontColorRgb = ParseThemeColor(fontNode.Attributes["color"].Value);
                        }
                    }

                    // Parse textframe margins
                    var textFrameNode = elemNode.SelectSingleNode("textframe");
                    if (textFrameNode != null)
                    {
                        format.MarginLeft = GetFloatAttribute(textFrameNode, "marginLeft", 0);
                        format.MarginRight = GetFloatAttribute(textFrameNode, "marginRight", 0);
                        format.MarginTop = GetFloatAttribute(textFrameNode, "marginTop", 0);
                        format.MarginBottom = GetFloatAttribute(textFrameNode, "marginBottom", 0);
                    }

                    // Parse paragraph alignment
                    var paraNode = elemNode.SelectSingleNode("paragraphformat");
                    if (paraNode != null)
                    {
                        format.TextAlignment = GetParagraphAlignment(paraNode);
                    }

                    // Parse fill
                    var fillNode = elemNode.SelectSingleNode("fill");
                    if (fillNode != null)
                    {
                        string visible = GetAttributeValue(fillNode, "visible", "0");
                        format.HasFill = visible == "1";
                        if (fillNode.Attributes["foreColor"] != null)
                        {
                            format.FillColorRgb = ParseThemeColor(fillNode.Attributes["foreColor"].Value);
                        }
                    }

                    // Parse line/border
                    var lineNode = elemNode.SelectSingleNode("line");
                    if (lineNode != null)
                    {
                        string visible = GetAttributeValue(lineNode, "visible", "0");
                        format.HasLine = visible == "1";
                    }
                }
            }

            return format;
        }

        private FontInfo ParseFont(XmlNode fontNode)
        {
            if (fontNode == null) return null;

            return new FontInfo
            {
                Name = GetAttributeValue(fontNode, "name", "Segoe UI"),
                Size = GetFloatAttribute(fontNode, "size", 14),
                Bold = GetAttributeValue(fontNode, "bold", "0") == "1",
                Italic = GetAttributeValue(fontNode, "italic", "0") == "1",
                Color = ParseThemeColor(GetAttributeValue(fontNode, "color", "0"))
            };
        }

        private int GetParagraphAlignment(XmlNode paraNode)
        {
            if (paraNode == null) return 1; // Default left

            string align = GetAttributeValue(paraNode, "alignment", "1");
            // PpParagraphAlignment: 1=Left, 2=Center, 3=Right, 4=Justify
            if (int.TryParse(align, out int result))
            {
                return result;
            }
            return 1;
        }

        private XmlNode FindSelectedCase(XmlNodeList caseNodes, int level)
        {
            foreach (XmlNode caseNode in caseNodes)
            {
                string levelStr = GetAttributeValue(caseNode, "level", "1");
                string selectedStr = GetAttributeValue(caseNode, "selected", "0");

                if (int.TryParse(levelStr, out int caseLevel) && caseLevel == level && selectedStr == "1")
                {
                    return caseNode;
                }
            }
            return null;
        }

        private int ParseThemeColor(string colorIndex)
        {
            // External tools use theme color indices
            // Convert common indices to RGB approximations
            // This is a simplified mapping - full implementation would need PowerPoint theme access
            switch (colorIndex)
            {
                case "1": return 0x000000;      // Dark 1 - Black
                case "2": return 0xFFFFFF;      // Light 1 - White
                case "3": return 0x4472C4;      // Dark 2 - Navy
                case "4": return 0xEDEDED;      // Light 2 - Light Gray
                case "5": return 0xED7D31;      // Accent 1 - Orange
                case "6": return 0x70AD47;      // Accent 2 - Green
                case "7": return 0x5B9BD5;      // Accent 3 - Blue
                case "8": return 0x264D64;      // Accent 4 - Dark Blue
                case "9": return 0x7030A0;      // Accent 5 - Purple
                case "10": return 0xFFC000;     // Accent 6 - Amber
                case "11": return 0xFF6600;     // Hyperlink - Orange
                case "12": return 0x954F07;     // Followed Hyperlink
                case "13": return 0x000000;     // Common text color
                case "14": return 0x000000;     // Common text color
                case "15": return 0xFFFFFF;     // White
                case "16": return 0xC00000;     // Red
                case "17": return 0x70AD47;     // Green
                case "18": return 0x1F4E79;     // Dark Blue
                case "19": return 0xFFC000;     // Amber
                default: return 0x000000;       // Default black
            }
        }

        private string GetAttributeValue(XmlNode node, string attributeName, string defaultValue)
        {
            return node?.Attributes?[attributeName]?.Value ?? defaultValue;
        }

        private float GetFloatAttribute(XmlNode node, string attributeName, float defaultValue)
        {
            string value = node?.Attributes?[attributeName]?.Value;
            if (float.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float result))
            {
                return result;
            }
            return defaultValue;
        }

        #endregion

        #region Helper Classes

        private class FontInfo
        {
            public string Name { get; set; }
            public float Size { get; set; }
            public bool Bold { get; set; }
            public bool Italic { get; set; }
            public int Color { get; set; }
        }

        #endregion
    }
}
