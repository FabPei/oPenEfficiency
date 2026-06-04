using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnAnonymize",
        Name = "Anonymize Text",
        Tooltip = "Anonymize text",
        IconData = "M12,2A10,10 0 0,1 22,12A10,10 0 0,1 12,22A10,10 0 0,1 2,12A10,10 0 0,1 12,2M12,4A8,8 0 0,0 4,12A8,8 0 0,0 12,20A8,8 0 0,0 20,12A8,8 0 0,0 12,4M12,6A6,6 0 0,1 18,12A6,6 0 0,1 12,18A6,6 0 0,1 6,12A6,6 0 0,1 12,6Z",
        Color = "#6B7280",
        Description = "Replaces all text in selected slides with Lorem Ipsum words, preserving word count.",
        DetailedHelpText = "### Anonymize Text\nReplaces all identifiable text across the presentation (names, emails, client references) with anonymized placeholder strings.",
        Keywords = "lorem ipsum, fake text, sanitize text, hide text, replace content",
        MinSelection = 0)]
    public static class AnonymizeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var slides = manager.GetSelectedSlidesForOperation();
                foreach (Slide slide in slides)
                {
                    foreach (Shape shape in slide.Shapes)
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            var textRange = shape.TextFrame.TextRange;
                            int wordCount = textRange.Words().Count;
                            textRange.Text = GenerateLoremIpsum(wordCount);
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AnonymizeFeature.Execute");
                return false;
            }
        }

        private static string GenerateLoremIpsum(int wordCount)
        {
            string[] words = { "lorem", "ipsum", "dolor", "sit", "amet", "consectetur", "adipiscing", "elit", "sed", "do", "eiusmod", "tempor", "incididunt", "ut", "labore", "et", "dolore", "magna", "aliqua" };
            var random = new Random();
            var result = new List<string>();
            for (int i = 0; i < wordCount; i++)
            {
                result.Add(words[random.Next(words.Length)]);
            }
            if (result.Count > 0)
            {
                string s = string.Join(" ", result);
                return char.ToUpper(s[0]) + s.Substring(1) + ".";
            }
            return "";
        }
    }
}
