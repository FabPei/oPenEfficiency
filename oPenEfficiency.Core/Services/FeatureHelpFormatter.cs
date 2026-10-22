using System;
using System.Text.RegularExpressions;

namespace oPenEfficiency.Services
{
    public static class FeatureHelpFormatter
    {
        public static string ToReadablePlainText(string helpText)
        {
            if (string.IsNullOrWhiteSpace(helpText)) return string.Empty;

            string text = helpText.Replace("\r\n", "\n");
            text = Regex.Replace(text, @"(?m)^\s{0,3}#{1,6}\s*", string.Empty);
            text = Regex.Replace(text, @"\*\*(.+?)\*\*", "$1");
            text = Regex.Replace(text, @"__(.+?)__", "$1");
            text = Regex.Replace(text, @"(?<!\*)\*([^*\r\n]+)\*(?!\*)", "$1");
            text = Regex.Replace(text, @"(?<!_)_([^_\r\n]+)_(?!_)", "$1");
            text = Regex.Replace(text, @"`([^`\r\n]+)`", "$1");
            text = Regex.Replace(text, @"(?m)^\s*[-*]\s+", "• ");
            text = Regex.Replace(text, @"\n{3,}", "\n\n");
            return text.Trim().Replace("\n", Environment.NewLine);
        }
    }
}
