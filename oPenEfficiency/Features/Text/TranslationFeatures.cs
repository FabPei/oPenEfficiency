using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Text
{
    [FeatureMetadata(
        Id = "BtnTranslate",
        Name = "Translate",
        Tooltip = "Translate text in selected shapes",
        IconData = "M5 8l6 6 M4 14h7 M2 5h12 M7 2l3 3-3 3 M22 22l-5-10-5 10 M14 18h6",
        Color = "#EAB308",
        Description = "Opens PowerPoint's built-in Translator pane to translate text in selected shapes.",
        DetailedHelpText = "### Translate\nOpens the Microsoft Translator pane integrated in PowerPoint. Select shapes or text and use this feature to translate content into any supported language.\n\nFor AI-powered translation (DeepL, OpenAI, Claude, LibreTranslate), configure the translation provider in Settings.",
        Keywords = "translator, translation, deepL, ai translate, convert language",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class TranslationFeatures
    {
        private class TextTranslationRequest
        {
            public TextRange Range { get; set; }
            public string OriginalText { get; set; }
            public string TranslatedText { get; set; }
            
            // Formatting to preserve
            public float FontSize { get; set; }
            public string FontName { get; set; }
            public int FontColor { get; set; }
            public Microsoft.Office.Core.MsoTriState FontBold { get; set; }
            public Microsoft.Office.Core.MsoTriState FontItalic { get; set; }
        }

        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var app = manager.GetApplication();
                
                // Ensure the window is focused
                try { if (app.ActiveWindow != null) app.ActiveWindow.Activate(); } catch { }

                // Pass 1: Try modern ExecuteMso
                string[] msoIds = { "Translator", "Translate", "MicrosoftTranslator", "ReviewTranslate", "ReviewSelectionTranslate" };
                foreach (var id in msoIds)
                {
                    try 
                    { 
                        if (app.CommandBars.GetEnabledMso(id)) 
                        { 
                            app.CommandBars.ExecuteMso(id); 
                            return true; 
                        } 
                    } 
                    catch { }
                }

                // Pass 2: Legacy CommandBars approach (showing the Research pane)
                try
                {
                    var researchBar = app.CommandBars["Research"];
                    if (researchBar != null)
                    {
                        researchBar.Visible = true;
                        return true;
                    }
                }
                catch { }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Translation trigger failed: {ex.Message}");
                return false;
            }
        }

        public static async void TranslateSelectionSafe(PowerPointManager manager, SidebarConfig config)
        {
            var app = manager.GetApplication();
            if (app == null || app.ActiveWindow == null) return;
            
            var selection = app.ActiveWindow.Selection;
            if (selection == null) return;

            var textsToTranslate = new List<TextTranslationRequest>();
            
            try
            {
                // UI Visual feedback
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                
                // Synchronously capture everything on main thread
                if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    CaptureTextRangeForTranslation(selection.TextRange, textsToTranslate);
                }
                else if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    for (int i = 1; i <= selection.ShapeRange.Count; i++)
                    {
                        var shape = selection.ShapeRange[i];
                        if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            CaptureTextRangeForTranslation(shape.TextFrame.TextRange, textsToTranslate);
                        }
                    }
                }

                if (textsToTranslate.Count == 0) return;

                // Fire async network calls on thread-pool to prevent hanging UI
                foreach (var req in textsToTranslate)
                {
                    req.TranslatedText = await Task.Run(() => GetTranslatedTextAsync(req.OriginalText, config));
                }

                // Await returns us to main thread. Resume synchronously applying results.
                ApplyTranslatedText(textsToTranslate);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Translation Error: {ex.Message}", "oPenEfficiency", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            finally
            {
                try { System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default; } catch { }
            }
        }

        public static async void TranslateWithFreeApiSafe(PowerPointManager manager, string targetLang)
        {
            var app = manager.GetApplication();
            if (app == null || app.ActiveWindow == null) return;
            
            var selection = app.ActiveWindow.Selection;
            if (selection == null) return;

            var textsToTranslate = new List<TextTranslationRequest>();
            
            try
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                
                if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    CaptureTextRangeForTranslation(selection.TextRange, textsToTranslate);
                }
                else if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    for (int i = 1; i <= selection.ShapeRange.Count; i++)
                    {
                        var shape = selection.ShapeRange[i];
                        if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            CaptureTextRangeForTranslation(shape.TextFrame.TextRange, textsToTranslate);
                        }
                    }
                }

                if (textsToTranslate.Count == 0) return;

                foreach (var req in textsToTranslate)
                {
                    req.TranslatedText = await Task.Run(() => GetMyMemoryTranslationAsync(req.OriginalText, targetLang));
                }

                ApplyTranslatedText(textsToTranslate);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Free API Translation Error:\n{ex.Message}", "oPenEfficiency", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
            finally
            {
                try { System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default; } catch { }
            }
        }

        private static void CaptureTextRangeForTranslation(TextRange textRange, List<TextTranslationRequest> list)
        {
            if (textRange == null || string.IsNullOrWhiteSpace(textRange.Text)) return;
            
            var req = new TextTranslationRequest 
            { 
                Range = textRange, 
                OriginalText = textRange.Text 
            };

            // Store initial formatting to reapply later
            try
            {
                req.FontSize = textRange.Font.Size;
                req.FontName = textRange.Font.Name;
                req.FontBold = textRange.Font.Bold;
                req.FontItalic = textRange.Font.Italic;
                
                // Avoid crashing on color extraction if varied
                try { req.FontColor = textRange.Font.Color.RGB; } catch { req.FontColor = 0; }
            }
            catch { /* Ignore formatting extraction errors */ }

            list.Add(req);
        }

        private static void ApplyTranslatedText(List<TextTranslationRequest> requests)
        {
            foreach (var req in requests)
            {
                if (!string.IsNullOrEmpty(req.TranslatedText) && req.Range != null)
                {
                    req.Range.Text = req.TranslatedText;
                    
                    try
                    {
                        req.Range.Font.Size = req.FontSize;
                        req.Range.Font.Name = req.FontName;
                        req.Range.Font.Bold = req.FontBold;
                        req.Range.Font.Italic = req.FontItalic;
                        if (req.FontColor != 0) 
                        {
                            req.Range.Font.Color.RGB = req.FontColor;
                        }
                    }
                    catch { /* Ignore format restore failures */ }
                }
            }
        }

        private static async Task<string> GetTranslatedTextAsync(string text, SidebarConfig config)
        {
            if (config.SelectedTranslationTool == "DeepL")
            {
                var deepL = new DeepLService(config.DeepLApiKey);
                return await deepL.TranslateAsync(text, config.TranslationTargetLanguage);
            }
            else if (config.SelectedTranslationTool == "OpenAI")
            {
                var llm = new LLMTranslationService();
                return await llm.TranslateWithOpenAIAsync(text, config.TranslationTargetLanguage, config.OpenAIApiKey, config.OpenAIModel);
            }
            else if (config.SelectedTranslationTool == "Claude")
            {
                var llm = new LLMTranslationService();
                return await llm.TranslateWithClaudeAsync(text, config.TranslationTargetLanguage, config.ClaudeApiKey, config.ClaudeModel);
            }
            else if (config.SelectedTranslationTool == "LibreTranslate")
            {
                var lt = new LibreTranslateService(config.LibreTranslateBaseUrl, config.LibreTranslateApiKey);
                return await lt.TranslateAsync(text, config.TranslationTargetLanguage);
            }
            return null;
        }

        private static async Task<string> GetMyMemoryTranslationAsync(string text, string targetLang)
        {
            try
            {
                // Force TLS 1.2 for modern API compatibility in .NET 4.8
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

                using (var client = new HttpClient())
                {
                    string target = targetLang.Split('-')[0].ToLower();
                    string url = $"https://api.mymemory.translated.net/get?q={Uri.EscapeDataString(text)}&langpair=autodetect|{target}";
                    
                    var response = await client.GetStringAsync(url);
                    int start = response.IndexOf("\"translatedText\":\"") + 18;
                    int end = response.IndexOf("\"", start);
                    if (start > 17 && end > start)
                    {
                        string translated = response.Substring(start, end - start);
                        return WebUtility.HtmlDecode(translated);
                    }
                }
            }
            catch { }
            return null;
        }
    }
}
