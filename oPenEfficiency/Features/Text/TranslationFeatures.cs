using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Text
{
    public static class TranslationFeatures
    {
        private static bool _translationInProgress;

        private class TextTranslationRequest
        {
            public TextRange Range { get; set; }
            public string OriginalText { get; set; }
            public string TranslatedText { get; set; }
            
            // Formatting to preserve
            public float FontSize { get; set; }
            public string FontName { get; set; }
            public int FontColor { get; set; }
            public bool HasFontColor { get; set; }
            public Microsoft.Office.Core.MsoTriState FontBold { get; set; }
            public Microsoft.Office.Core.MsoTriState FontItalic { get; set; }
        }

        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app?.CommandBars == null) return false;
                
                // Ensure the window is focused
                try { if (app.ActiveWindow != null) app.ActiveWindow.Activate(); }
                catch (Exception optionalEx) { ExceptionPolicy.ReportOptionalComFailure(optionalEx, "TranslationFeatures.TranslateSelection.ActivateWindow"); }

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
                    catch (Exception optionalEx) { ExceptionPolicy.ReportOptionalComFailure(optionalEx, "TranslationFeatures.TranslateSelection.StartUndoEntry"); }
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
                catch (Exception optionalEx) { ExceptionPolicy.ReportOptionalComFailure(optionalEx, "TranslationFeatures.TranslateSelection.RestoreSelection"); }

                System.Windows.MessageBox.Show(
                    "PowerPoint's translation interface is not available in the current context.",
                    "Translate Text",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TranslationFeatures.Execute");
                return false;
            }
        }

        public static async void TranslateSelectionSafe(PowerPointManager manager, SidebarConfig config)
        {
            if (manager == null || config == null) return;
            var app = manager.GetApplication();
            if (app == null || app.ActiveWindow == null) return;
            var selection = app.ActiveWindow.Selection;
            if (selection == null) return;
            if (_translationInProgress)
            {
                System.Windows.MessageBox.Show("A translation is already in progress.", "Translation");
                return;
            }
            _translationInProgress = true;
            var textsToTranslate = new List<TextTranslationRequest>();
            
            try
            {
                var sourcePresentation = app.ActivePresentation;
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
                if (app.ActivePresentation != sourcePresentation)
                {
                    System.Windows.MessageBox.Show(
                        "Translation was canceled because the active presentation changed.",
                        "Translation",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                    return;
                }

                if (!textsToTranslate.Exists(request => !string.IsNullOrWhiteSpace(request.TranslatedText)))
                {
                    System.Windows.MessageBox.Show(
                        "The configured provider returned no translation. The presentation was not changed.",
                        "Translation",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                    return;
                }

                ApplyTranslatedText(app, textsToTranslate);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TranslationFeatures.TranslateSelectionSafe");
                System.Windows.MessageBox.Show($"Translation Error: {ex.Message}", "openEfficiency", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
            finally
            {
                _translationInProgress = false;
                try { System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default; }
                catch (Exception cleanupEx) { ExceptionPolicy.ReportOptionalComFailure(cleanupEx, "TranslationFeatures.TranslateSelection.RestoreCursor"); }
            }
        }

        public static async void TranslateWithFreeApiSafe(PowerPointManager manager, string targetLang)
        {
            if (manager == null || string.IsNullOrWhiteSpace(targetLang)) return;
            var app = manager.GetApplication();
            if (app == null || app.ActiveWindow == null) return;
            var selection = app.ActiveWindow.Selection;
            if (selection == null) return;
            if (_translationInProgress)
            {
                System.Windows.MessageBox.Show("A translation is already in progress.", "Translation");
                return;
            }
            _translationInProgress = true;
            var textsToTranslate = new List<TextTranslationRequest>();
            
            try
            {
                var sourcePresentation = app.ActivePresentation;
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

                if (app.ActivePresentation != sourcePresentation)
                {
                    System.Windows.MessageBox.Show(
                        "Translation was canceled because the active presentation changed.",
                        "Translation",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                    return;
                }

                if (!textsToTranslate.Exists(request => !string.IsNullOrWhiteSpace(request.TranslatedText)))
                {
                    System.Windows.MessageBox.Show(
                        "The free translation service returned no translation. The presentation was not changed.",
                        "Translation",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning);
                    return;
                }

                ApplyTranslatedText(app, textsToTranslate);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TranslationFeatures.TranslateWithFreeApiSafe");
                System.Windows.MessageBox.Show($"Free API Translation Error:\n{ex.Message}", "openEfficiency", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
            finally
            {
                _translationInProgress = false;
                try { System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default; }
                catch (Exception cleanupEx) { ExceptionPolicy.ReportOptionalComFailure(cleanupEx, "TranslationFeatures.TranslateText.RestoreCursor"); }
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
                try
                {
                    req.FontColor = textRange.Font.Color.RGB;
                    req.HasFontColor = true;
                }
                catch { req.HasFontColor = false; }
            }
            catch { /* Ignore formatting extraction errors */ }

            list.Add(req);
        }

        private static void ApplyTranslatedText(Application app, List<TextTranslationRequest> requests)
        {
            var changedRequests = new List<TextTranslationRequest>();
            try
            {
                try { app?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "TranslationFeatures.StartNewUndoEntry"); }

                foreach (var req in requests)
                {
                    if (!string.IsNullOrEmpty(req.TranslatedText) && req.Range != null)
                    {
                        req.Range.Text = req.TranslatedText;
                        changedRequests.Add(req);
                        RestoreBaseFormatting(req);
                    }
                }
            }
            catch (Exception ex)
            {
                for (int i = changedRequests.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        var req = changedRequests[i];
                        req.Range.Text = req.OriginalText;
                        RestoreBaseFormatting(req);
                    }
                    catch (Exception rollbackEx)
                    {
                        ExceptionLogger.Log(rollbackEx, "TranslationFeatures.ApplyTranslatedText.Rollback");
                    }
                }
                ExceptionLogger.Log(ex, "TranslationFeatures.ApplyTranslatedText");
                throw;
            }
        }

        private static void RestoreBaseFormatting(TextTranslationRequest request)
        {
            if (request == null || request.Range == null) return;
            if (request.FontSize > 0) request.Range.Font.Size = request.FontSize;
            if (!string.IsNullOrWhiteSpace(request.FontName)) request.Range.Font.Name = request.FontName;
            if (request.FontBold != Microsoft.Office.Core.MsoTriState.msoTriStateMixed)
                request.Range.Font.Bold = request.FontBold;
            if (request.FontItalic != Microsoft.Office.Core.MsoTriState.msoTriStateMixed)
                request.Range.Font.Italic = request.FontItalic;
            if (request.HasFontColor) request.Range.Font.Color.RGB = request.FontColor;
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
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TranslationFeatures.GetMyMemoryTranslationAsync");
            }
            return null;
        }
    }
}
