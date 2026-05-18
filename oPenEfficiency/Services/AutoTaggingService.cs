using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.Services
{
    public class AutoTaggingService
    {
        private static TagConfig _config;

        public static List<string> IdentifyTags(Slide slide)
        {
            LoadConfig();
            if (_config == null) return new List<string>();

            var tags = new List<string>();
            string slideText = ExtractTextFromSlide(slide);
            int wordCount = CountWords(slideText);

            // 1. Structural Tags
            foreach (var entry in _config.StructuralTags)
            {
                if (MatchesRule(entry.Value, slideText, wordCount, slide))
                {
                    tags.Add(FormatTagName(entry.Key));
                }
            }

            // 2. Content Tags
            foreach (var entry in _config.ContentTags)
            {
                if (MatchesRule(entry.Value, slideText, wordCount, slide))
                {
                    tags.Add(FormatTagName(entry.Key));
                }
            }

            // 3. Industry Context
            foreach (var entry in _config.IndustryContext)
            {
                int matchCount = 0;
                foreach (var keyword in entry.Value)
                {
                    if (slideText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                        matchCount++;
                }
                
                // If more than 2 keywords match, add industry tag
                if (matchCount >= 2)
                {
                    tags.Add(FormatTagName(entry.Key));
                }
            }

            return tags.Distinct().ToList();
        }

        private static bool MatchesRule(TagRule rule, string text, int wordCount, Slide slide)
        {
            // Keyword match (Full word matching for better accuracy)
            if (rule.Keywords != null)
            {
                foreach (var kw in rule.Keywords)
                {
                    if (string.IsNullOrEmpty(kw)) continue;
                    // Use regex for whole-word matching to avoid partial matches (e.g. "age" matching "agenda")
                    string pattern = @"\b" + Regex.Escape(kw) + @"\b";
                    if (Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase))
                        return true;
                }
            }

            // Pattern match (Regex)
            if (rule.Patterns != null)
            {
                foreach (var pattern in rule.Patterns)
                {
                    try
                    {
                        if (string.IsNullOrEmpty(pattern)) continue;
                        if (Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase))
                            return true;
                    }
                    catch { }
                }
            }

            // Logic heuristics
            if (!string.IsNullOrEmpty(rule.Logic))
            {
                if (rule.Logic.Contains("low word count") && wordCount > 0 && wordCount < 15) return true;
                if (rule.Logic.Contains("slide 2 or 3") && (slide.SlideIndex == 2 || slide.SlideIndex == 3)) return true;
            }

            return false;
        }

        private static string ExtractTextFromSlide(Slide slide)
        {
            var sb = new StringBuilder();
            foreach (Shape shape in slide.Shapes)
            {
                ExtractTextFromShape(shape, sb);
            }
            return sb.ToString();
        }

        private static void ExtractTextFromShape(Shape shape, StringBuilder sb)
        {
            try
            {
                // 1. Standard Text Frames
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    sb.AppendLine(shape.TextFrame.TextRange.Text);
                }

                // 2. Tables
                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    var table = shape.Table;
                    for (int r = 1; r <= table.Rows.Count; r++)
                    {
                        for (int c = 1; c <= table.Columns.Count; c++)
                        {
                            var cell = table.Cell(r, c);
                            if (cell.Shape.HasTextFrame == Office.MsoTriState.msoTrue)
                                sb.AppendLine(cell.Shape.TextFrame.TextRange.Text);
                        }
                    }
                }

                // 3. Groups (Recursive)
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    foreach (Shape subShape in shape.GroupItems)
                    {
                        ExtractTextFromShape(subShape, sb);
                    }
                }

                // 4. SmartArt (peek into nodes if possible)
                if (shape.HasSmartArt == Office.MsoTriState.msoTrue)
                {
                    foreach (Office.SmartArtNode node in shape.SmartArt.AllNodes)
                    {
                        if (node.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                            sb.AppendLine(node.TextFrame2.TextRange.Text);
                    }
                }
            }
            catch { }
        }

        private static int CountWords(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;
            return text.Split(new[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length;
        }

        private static string FormatTagName(string key)
        {
            // Convert agenda_slide to "Agenda Slide"
            var parts = key.Split('_');
            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i].Length > 0)
                    parts[i] = char.ToUpper(parts[i][0]) + parts[i].Substring(1);
            }
            return string.Join(" ", parts);
        }

        private static void LoadConfig()
        {
            if (_config != null) return;

            try
            {
                // Find Resources/SlideTaggingRules.json relative to the add-in path or known locations
                // For simplicity, let's assume it's deployed or we can find it in the dev env.
                // In a real VSTO, you'd embed it as a resource or put it in AppData.
                
                string assemblyPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                // Check possible locations
                string[] paths = {
                    Path.Combine(assemblyPath, "Resources", "SlideTaggingRules.json"),
                    Path.Combine(assemblyPath, "SlideTaggingRules.json")
                };

                string configPath = paths.FirstOrDefault(File.Exists);
                if (configPath == null) return;

                string json = File.ReadAllText(configPath);
                
                // Manual parsing since DataContractJsonSerializer is a pain with dynamic keys
                // We'll use a simplified regex-based parser or just deserialize into a Dictionary
                _config = ParseConfig(json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("AutoTaggingService.LoadConfig error: " + ex.Message);
            }
        }

        private static TagConfig ParseConfig(string json)
        {
            // Extremely simplified "JSON" parser for this specific file structure
            // In a production app, use Newtonsoft.Json
            var config = new TagConfig();
            
            // Extract sections
            config.StructuralTags = ParseSection(json, "structural_tags");
            config.ContentTags = ParseSection(json, "content_tags");
            
            // Industry context is slightly different (list of strings)
            config.IndustryContext = ParseIndustrySection(json);

            return config;
        }

        private static Dictionary<string, TagRule> ParseSection(string json, string sectionName)
        {
            var dict = new Dictionary<string, TagRule>();
            var sectionMatch = Regex.Match(json, "\"" + sectionName + "\":\\s*\\{(.*?)\\}\\s*,\\s*\"[a-z_]+\":", RegexOptions.Singleline);
            if (!sectionMatch.Success)
            {
                // Try if it's the last section
                sectionMatch = Regex.Match(json, "\"" + sectionName + "\":\\s*\\{(.*?)\\}\\s*\\}", RegexOptions.Singleline);
            }

            if (sectionMatch.Success)
            {
                string sectionContent = sectionMatch.Groups[1].Value;
                // Find each tag object
                var tagMatches = Regex.Matches(sectionContent, "\"([a-z_]+)\":\\s*\\{(.*?)\\}", RegexOptions.Singleline);
                foreach (Match m in tagMatches)
                {
                    string tagName = m.Groups[1].Value;
                    string tagContent = m.Groups[2].Value;
                    var rule = new TagRule();
                    
                    // Keywords
                    var kwMatch = Regex.Match(tagContent, "\"keywords\":\\s*\\[(.*?)\\]", RegexOptions.Singleline);
                    if (kwMatch.Success)
                    {
                        rule.Keywords = kwMatch.Groups[1].Value.Split(',').Select(s => s.Trim(' ', '"', '\r', '\n')).ToList();
                    }

                    // Patterns
                    var ptMatch = Regex.Match(tagContent, "\"patterns\":\\s*\\[(.*?)\\]", RegexOptions.Singleline);
                    if (ptMatch.Success)
                    {
                        rule.Patterns = ptMatch.Groups[1].Value.Split(',').Select(s => s.Trim(' ', '"', '\r', '\n')).ToList();
                    }

                    // Logic
                    var lgMatch = Regex.Match(tagContent, "\"logic\":\\s*\"(.*?)\"", RegexOptions.Singleline);
                    if (lgMatch.Success)
                    {
                        rule.Logic = lgMatch.Groups[1].Value;
                    }

                    dict[tagName] = rule;
                }
            }
            return dict;
        }

        private static Dictionary<string, List<string>> ParseIndustrySection(string json)
        {
            var dict = new Dictionary<string, List<string>>();
            var sectionMatch = Regex.Match(json, "\"industry_context\":\\s*\\{(.*?)\\}\\s*\\}", RegexOptions.Singleline);
            if (sectionMatch.Success)
            {
                string content = sectionMatch.Groups[1].Value;
                var tagMatches = Regex.Matches(content, "\"([a-z_]+)\":\\s*\\[(.*?)\\]", RegexOptions.Singleline);
                foreach (Match m in tagMatches)
                {
                    string name = m.Groups[1].Value;
                    var list = m.Groups[2].Value.Split(',').Select(s => s.Trim(' ', '"', '\r', '\n')).ToList();
                    dict[name] = list;
                }
            }
            return dict;
        }
    }

    public class TagConfig
    {
        public Dictionary<string, TagRule> StructuralTags { get; set; } = new Dictionary<string, TagRule>();
        public Dictionary<string, TagRule> ContentTags { get; set; } = new Dictionary<string, TagRule>();
        public Dictionary<string, List<string>> IndustryContext { get; set; } = new Dictionary<string, List<string>>();
    }

    public class TagRule
    {
        public List<string> Keywords { get; set; }
        public List<string> Patterns { get; set; }
        public string Logic { get; set; }
    }
}
