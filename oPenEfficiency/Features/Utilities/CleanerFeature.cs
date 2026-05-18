using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Clean up presentations by removing hidden slides, animations, notes, comments, etc.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnCleaner",
        Name = "Cleaner",
        Tooltip = "Cleaner",
        IconData = "M19.36,2.72L20.78,4.14L15.06,9.85C15.5,10.5 15.81,11.23 15.96,12.03L19.09,8.91L20.5,10.32L16.04,14.78C15.16,15.66 13.94,16.24 12.61,16.37L10.38,18.6L8.96,17.18L11.19,14.95C11.06,13.62 11.64,12.4 12.52,11.52L16.98,7.06L18.39,8.47L15.27,11.6C15.47,10.27 15.08,8.87 14.2,8L19.36,2.72M5.21,10C4.71,10.81 4.35,11.7 4.15,12.65L2.74,11.24L1.32,12.65L5.21,16.54L9.1,12.65L7.68,11.24L6.28,12.65C6.18,11.85 6.22,11.05 6.4,10.27L5.21,10Z",
        Color = "#F43F5E",
        Description = "Cleans up presentations by removing hidden slides, animations, notes, and comments.",
        DetailedHelpText = "### Cleaner\nPerforms a comprehensive cleanup of the presentation. Left-click to open the main cleanup dialog.\n\n**Right-Click Options:**\n* **Remove Hidden Slides**: Deletes all slides marked as hidden.\n* **Remove Animations**: Strip all animations (All Slides / Selected).\n* **Remove Notes**: Wipe Speaker or Developer notes.\n* **Anonymize**: Remove all personal metadata and author tags.",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class CleanerFeature
    {
        public enum CleanAction
        {
            RemoveHiddenSlides,
            RemoveAnimationsAll,
            RemoveAnimationsSelected,
            RemoveTransitionsAll,
            RemoveTransitionsSelected,
            RemoveSpeakerNotesAll,
            RemoveSpeakerNotesSelected,
            RemoveCommentsAll,
            RemoveCommentsSelected,
            RemoveUnusedMasters,
            DeleteStickyNotes,
            AnonymizeSelected,
            RemoveUneditedFootersAll,
            RemoveUneditedFootersSelected,
            DetectSuperscriptWithoutFooter,
            RemoveDoubleSpacesAll,
            RemoveDoubleSpacesSelected,
            ConvertSlidesToImagesAll,
            ConvertSlidesToImagesSelected,
            RemoveDeveloperNotesAll,
            RemoveDeveloperNotesSelected
        }

        /// <summary>
        /// Wrapper for auto-discovery - returns false as Cleaner requires dialog for action selection.
        /// Manual switch in MainSidebar.xaml.cs handles opening the Cleaner dialog.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            // Requires dialog for clean action selection
            // Handled by manual switch in MainSidebar.xaml.cs
            return false;
        }

        public static bool Execute(PowerPointManager manager, CleanAction action)
        {
            if (manager == null) return false;

            switch (action)
            {
                case CleanAction.RemoveHiddenSlides: return RemoveHiddenSlides(manager);
                case CleanAction.RemoveAnimationsAll: return RemoveAnimations(manager, false);
                case CleanAction.RemoveAnimationsSelected: return RemoveAnimations(manager, true);
                case CleanAction.RemoveTransitionsAll: return RemoveTransitions(manager, false);
                case CleanAction.RemoveTransitionsSelected: return RemoveTransitions(manager, true);
                case CleanAction.RemoveSpeakerNotesAll: return RemoveSpeakerNotes(manager, false);
                case CleanAction.RemoveSpeakerNotesSelected: return RemoveSpeakerNotes(manager, true);
                case CleanAction.RemoveCommentsAll: return RemoveComments(manager, false);
                case CleanAction.RemoveCommentsSelected: return RemoveComments(manager, true);
                case CleanAction.RemoveUnusedMasters: return RemoveUnusedMasters(manager);
                case CleanAction.DeleteStickyNotes: return DeleteAllStickyNotes(manager);
                case CleanAction.AnonymizeSelected: return AnonymizeFeature.Execute(manager);
                case CleanAction.RemoveUneditedFootersAll: return RemoveUneditedFooters(manager, false);
                case CleanAction.RemoveUneditedFootersSelected: return RemoveUneditedFooters(manager, true);
                case CleanAction.DetectSuperscriptWithoutFooter: return DetectSuperscriptWithoutFooter(manager);
                case CleanAction.RemoveDoubleSpacesAll: return RemoveDoubleSpaces(manager, false);
                case CleanAction.RemoveDoubleSpacesSelected: return RemoveDoubleSpaces(manager, true);
                case CleanAction.ConvertSlidesToImagesAll: return ConvertSlidesToImages(manager, false);
                case CleanAction.ConvertSlidesToImagesSelected: return ConvertSlidesToImages(manager, true);
                case CleanAction.RemoveDeveloperNotesAll: return RemoveDeveloperNotes(manager, false);
                case CleanAction.RemoveDeveloperNotesSelected: return RemoveDeveloperNotes(manager, true);
                default: return false;
            }
        }

        public static bool RemoveDeveloperNotes(PowerPointManager manager, bool selectedOnly)
        {
            try
            {
                var app = manager.GetApplication();
                IEnumerable<Slide> slides;

                if (selectedOnly)
                {
                    var sel = app.ActiveWindow.Selection;
                    if (sel.Type == PpSelectionType.ppSelectionNone) return false;
                    var list = new List<Slide>();
                    for (int i = 1; i <= sel.SlideRange.Count; i++) list.Add(sel.SlideRange[i]);
                    slides = list;
                }
                else
                {
                    var list = new List<Slide>();
                    foreach (Slide s in app.ActivePresentation.Slides) list.Add(s);
                    slides = list;
                }

                foreach (Slide slide in slides)
                {
                    try { slide.Tags.Delete("oPE_SlideNotes"); } catch { }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.RemoveDeveloperNotes");
                return false;
            }
        }

        public static bool RemoveHiddenSlides(PowerPointManager manager)
        {
            try
            {
                var pres = manager.GetApplication().ActivePresentation;
                for (int i = pres.Slides.Count; i >= 1; i--)
                {
                    if (pres.Slides[i].SlideShowTransition.Hidden == Office.MsoTriState.msoTrue)
                        pres.Slides[i].Delete();
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.RemoveHiddenSlides");
                return false;
            }
        }

        public static bool RemoveAnimations(PowerPointManager manager, bool selectedOnly)
        {
            try
            {
                var slides = GetTargetSlides(manager, selectedOnly);
                foreach (Slide slide in slides)
                {
                    while (slide.TimeLine.MainSequence.Count > 0)
                        slide.TimeLine.MainSequence[1].Delete();
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.RemoveAnimations");
                return false;
            }
        }

        public static bool RemoveTransitions(PowerPointManager manager, bool selectedOnly)
        {
            try
            {
                var slides = GetTargetSlides(manager, selectedOnly);
                foreach (Slide slide in slides)
                {
                    try { ((dynamic)slide.SlideShowTransition).EntryEffect = 0; } catch (Exception ex) { ExceptionLogger.Log(ex, "CleanerFeature.RemoveTransitions.EntryEffect"); }
                    slide.SlideShowTransition.Duration = 0;
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.RemoveTransitions");
                return false;
            }
        }

        public static bool RemoveSpeakerNotes(PowerPointManager manager, bool selectedOnly)
        {
            try
            {
                var slides = GetTargetSlides(manager, selectedOnly);
                foreach (Slide slide in slides)
                {
                    if (slide.HasNotesPage == Office.MsoTriState.msoTrue)
                    {
                        foreach (Shape shape in slide.NotesPage.Shapes)
                        {
                            if (shape.Type == Office.MsoShapeType.msoPlaceholder && shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody)
                            {
                                shape.TextFrame.TextRange.Text = "";
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.RemoveSpeakerNotes");
                return false;
            }
        }

        public static bool RemoveComments(PowerPointManager manager, bool selectedOnly)
        {
            try
            {
                var slides = GetTargetSlides(manager, selectedOnly);
                foreach (Slide slide in slides)
                {
                    for (int i = slide.Comments.Count; i >= 1; i--)
                        slide.Comments[i].Delete();
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.RemoveComments");
                return false;
            }
        }

        public static bool RemoveUnusedMasters(PowerPointManager manager)
        {
            try
            {
                var pres = manager.GetApplication().ActivePresentation;
                pres.RemoveDocumentInformation(PpRemoveDocInfoType.ppRDIAll);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.RemoveUnusedMasters");
                return false;
            }
        }

        public static bool RemoveUneditedFooters(PowerPointManager manager, bool selectedOnly)
        {
            try
            {
                var slides = GetTargetSlides(manager, selectedOnly);
                foreach (Slide slide in slides)
                {
                    for (int i = slide.Shapes.Count; i >= 1; i--)
                    {
                        var shape = slide.Shapes[i];
                        if (shape.Type == Office.MsoShapeType.msoPlaceholder)
                        {
                            var p = shape.PlaceholderFormat;
                            if (p.Type == PpPlaceholderType.ppPlaceholderFooter ||
                                p.Type == PpPlaceholderType.ppPlaceholderHeader ||
                                p.Type == PpPlaceholderType.ppPlaceholderDate)
                            {
                                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                                {
                                    string text = shape.TextFrame.TextRange.Text.Trim();
                                    if (string.IsNullOrEmpty(text) || (text.StartsWith("<") && text.EndsWith(">") && text.Length > 1))
                                    {
                                        shape.Delete();
                                    }
                                }
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.RemoveUneditedFooters");
                return false;
            }
        }

        public static bool DeleteAllStickyNotes(PowerPointManager manager)
        {
            try
            {
                var pres = manager.GetApplication().ActivePresentation;
                foreach (Slide slide in pres.Slides)
                {
                    for (int i = slide.Shapes.Count; i >= 1; i--)
                    {
                        var shape = slide.Shapes[i];
                        if (shape.Name.StartsWith("StickyNote", StringComparison.OrdinalIgnoreCase))
                            shape.Delete();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.DeleteAllStickyNotes");
                return false;
            }
        }

        private static IEnumerable<Slide> GetTargetSlides(PowerPointManager manager, bool selectedOnly)
        {
            var app = manager.GetApplication();
            var pres = app.ActivePresentation;
            if (!selectedOnly)
            {
                foreach (Slide slide in pres.Slides) yield return slide;
                yield break;
            }

            if (app.ActiveWindow != null && app.ActiveWindow.Selection != null)
            {
                var selection = app.ActiveWindow.Selection;
                if (selection.Type == PpSelectionType.ppSelectionSlides)
                {
                    foreach (Slide slide in selection.SlideRange) yield return slide;
                }
                else if (app.ActiveWindow.View.Slide != null)
                {
                    yield return (Slide)app.ActiveWindow.View.Slide;
                }
            }
        }

        /// <summary>
        /// Detects superscript numbers in text that do not have a corresponding footer reference.
        /// Reports findings to the user via message box.
        /// </summary>
        public static bool DetectSuperscriptWithoutFooter(PowerPointManager manager)
        {
            try
            {
                var app = manager.GetApplication();
                var pres = app.ActivePresentation;
                var issuesFound = new List<string>();
                int totalSuperscripts = 0;
                int totalWithoutFooter = 0;

                // First, collect all footer text from all slides to check for references
                var footerTexts = new List<string>();
                foreach (Slide slide in pres.Slides)
                {
                    string footerText = GetFooterText(slide);
                    if (!string.IsNullOrEmpty(footerText))
                    {
                        footerTexts.Add(footerText.ToLower());
                    }
                }

                // Combine all footer texts for searching
                string allFooters = string.Join(" ", footerTexts);

                // Scan all slides for superscript numbers
                int slideIndex = 0;
                foreach (Slide slide in pres.Slides)
                {
                    slideIndex++;
                    var superscriptsInSlide = FindSuperscriptNumbers(slide);

                    foreach (var superInfo in superscriptsInSlide)
                    {
                        totalSuperscripts++;
                        string superNumber = superInfo.Number;

                        // Check if this number exists in any footer
                        bool foundInFooter = FooterContainsNumber(allFooters, superNumber);

                        if (!foundInFooter)
                        {
                            totalWithoutFooter++;
                            issuesFound.Add($"Slide {slideIndex}: Superscript '{superNumber}' found in \"{superInfo.Context}\" - No matching footer reference");
                        }
                    }
                }

                // Report results
                if (issuesFound.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        $"No issues found. All {totalSuperscripts} superscript number(s) have corresponding footer references.",
                        "Superscript Footnote Check Complete",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
                else
                {
                    string message = $"Found {totalWithoutFooter} superscript number(s) without corresponding footer references:\n\n" +
                                     string.Join("\n", issuesFound) +
                                     $"\n\nTotal superscripts found: {totalSuperscripts}";

                    System.Windows.Forms.MessageBox.Show(
                        message,
                        "Superscript Footnote Warning",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.DetectSuperscriptWithoutFooter");
                System.Windows.Forms.MessageBox.Show(
                    $"Error detecting superscript footnotes: {ex.Message}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// Finds all superscript numbers in a slide's text shapes.
        /// </summary>
        private static List<(string Number, string Context)> FindSuperscriptNumbers(Slide slide)
        {
            var results = new List<(string Number, string Context)>();

            try
            {
                foreach (Shape shape in slide.Shapes)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        var textRange = shape.TextFrame.TextRange;
                        string fullText = textRange.Text;

                        for (int i = 1; i <= fullText.Length; i++)
                        {
                            try
                            {
                                var ch = textRange.Characters(i, 1);
                                // Check if character is superscript and is a number
                                if (ch.Font.Superscript == Office.MsoTriState.msoTrue)
                                {
                                    string text = ch.Text;
                                    // Check if it's a digit or looks like a superscript number reference
                                    if (System.Text.RegularExpressions.Regex.IsMatch(text.Trim(), "^\\d+$"))
                                    {
                                        // Get context (surrounding text)
                                        int start = Math.Max(1, i - 10);
                                        int contextLen = Math.Min(30, fullText.Length - start + 1);
                                        string context = textRange.Characters(start, contextLen).Text.Trim();

                                        results.Add((text.Trim(), context));
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.FindSuperscriptNumbers");
            }

            return results;
        }

        /// <summary>
        /// Gets all footer text from a slide.
        /// </summary>
        private static string GetFooterText(Slide slide)
        {
            try
            {
                var footerTexts = new List<string>();

                foreach (Shape shape in slide.Shapes)
                {
                    if (shape.Type == Office.MsoShapeType.msoPlaceholder)
                    {
                        var p = shape.PlaceholderFormat;
                        if (p.Type == PpPlaceholderType.ppPlaceholderFooter)
                        {
                            if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                            {
                                footerTexts.Add(shape.TextFrame.TextRange.Text.Trim());
                            }
                        }
                    }
                }

                return string.Join(" ", footerTexts);
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Checks if the footer text contains a reference to the given number.
        /// </summary>
        private static bool FooterContainsNumber(string footerText, string number)
        {
            if (string.IsNullOrEmpty(footerText) || string.IsNullOrEmpty(number))
                return false;

            // Check for various reference patterns like "[1]", "1.", "^1", etc.
            // Also check for just the number followed by typical footnote patterns
            var patterns = new[]
            {
                $@"\[{number}\]",      // [1]
                $@"\^{number}",         // ^1
                $@"{number}\.",         // 1.
                $@"{number}\)",         // 1)
                $@"\b{number}\b",       // Just the number as whole word
            };

            foreach (var pattern in patterns)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(footerText, pattern))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Removes double (and multiple) spaces from text in shapes.
        /// </summary>
        public static bool RemoveDoubleSpaces(PowerPointManager manager, bool selectedOnly)
        {
            try
            {
                var slides = GetTargetSlides(manager, selectedOnly);
                int totalReplacements = 0;

                foreach (Slide slide in slides)
                {
                    foreach (Shape shape in slide.Shapes)
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            var textRange = shape.TextFrame.TextRange;
                            string originalText = textRange.Text;
                            string newText = System.Text.RegularExpressions.Regex.Replace(originalText, "  +", " ");

                            if (newText != originalText)
                            {
                                textRange.Text = newText;
                                totalReplacements++;
                            }
                        }
                    }
                }

                System.Windows.Forms.MessageBox.Show(
                    $"Removed double spaces from {totalReplacements} shape(s).",
                    "Clean Complete",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.RemoveDoubleSpaces");
                System.Windows.Forms.MessageBox.Show(
                    $"Error removing double spaces: {ex.Message}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// Converts slides to high-resolution PNG images and creates a new uneditable presentation.
        /// Original presentation remains unchanged.
        /// </summary>
        public static bool ConvertSlidesToImages(PowerPointManager manager, bool selectedOnly)
        {
            try
            {
                var app = manager.GetApplication();
                var pres = app.ActivePresentation;
                var slides = GetTargetSlides(manager, selectedOnly);

                // Get temp folder for exporting images
                string tempFolder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "oPenEfficiency_SlideImages");
                if (!System.IO.Directory.Exists(tempFolder))
                {
                    System.IO.Directory.CreateDirectory(tempFolder);
                }

                // Create a new presentation for the uneditable version
                Presentation newPres = app.Presentations.Add(Office.MsoTriState.msoFalse);
                int processedCount = 0;
                var slideList = new List<Slide>();
                foreach (Slide s in slides) slideList.Add(s);

                // Process slides in order
                foreach (Slide slide in slideList)
                {
                    int slideIndex = slide.SlideNumber;

                    try
                    {
                        // Export slide as high-res PNG (3x scale for ~300 DPI)
                        string imagePath = System.IO.Path.Combine(tempFolder, $"slide_{slideIndex}.png");

                        // Export the slide as PNG
                        slide.Export(imagePath, "PNG", (int)pres.PageSetup.SlideWidth * 3, (int)pres.PageSetup.SlideHeight * 3);

                        // Get slide dimensions (from original presentation)
                        float slideWidth = pres.PageSetup.SlideWidth;
                        float slideHeight = pres.PageSetup.SlideHeight;

                        // Copy page setup to new presentation
                        newPres.PageSetup.SlideWidth = slideWidth;
                        newPres.PageSetup.SlideHeight = slideHeight;

                        // Add a blank slide and insert the image
                        var newSlide = newPres.Slides.Add(newPres.Slides.Count + 1, PpSlideLayout.ppLayoutBlank);
                        var picture = newSlide.Shapes.AddPicture(
                            imagePath,
                            Office.MsoTriState.msoFalse,
                            Office.MsoTriState.msoCTrue,
                            0,
                            0,
                            slideWidth,
                            slideHeight);

                        // Lock the image position and aspect ratio
                        picture.LockAspectRatio = Office.MsoTriState.msoTrue;

                        processedCount++;
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, $"CleanerFeature.ConvertSlidesToImages.Slide{slideIndex}");
                    }
                }

                // Clean up temp files
                try
                {
                    if (System.IO.Directory.Exists(tempFolder))
                    {
                        System.IO.Directory.Delete(tempFolder, true);
                    }
                }
                catch { }

                // Generate output filename based on original presentation
                string originalPath = pres.FullName;
                string originalDir = System.IO.Path.GetDirectoryName(originalPath);
                string originalName = System.IO.Path.GetFileNameWithoutExtension(originalPath);
                string uneditablePath = System.IO.Path.Combine(originalDir, $"{originalName}_UNEDITABLE.pptx");

                // Handle case where file already exists
                if (System.IO.File.Exists(uneditablePath))
                {
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    uneditablePath = System.IO.Path.Combine(originalDir, $"{originalName}_UNEDITABLE_{timestamp}.pptx");
                }

                // Save the new presentation
                newPres.SaveAs(uneditablePath);

                System.Windows.Forms.MessageBox.Show(
                    $"Created uneditable presentation with {processedCount} slide(s) converted to images.\n\n" +
                    $"Saved to:\n{uneditablePath}\n\n" +
                    $"Original presentation remains unchanged.",
                    "Conversion Complete",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.ConvertSlidesToImages");
                System.Windows.Forms.MessageBox.Show(
                    $"Error converting slides to images: {ex.Message}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }
    }
}
