using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using CleanAction = oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Clean up presentations by removing hidden slides, animations, notes, comments, etc.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnCleaner",
        Name = "Presentation Cleaner",
        Tooltip = "Review and clean presentation content",
        SearchAliases = new[] { "Deep Presentation Cleaner", "Cleaner" },
        IconData = "M19.36,2.72L20.78,4.14L15.06,9.85C15.5,10.5 15.81,11.23 15.96,12.03L19.09,8.91L20.5,10.32L16.04,14.78C15.16,15.66 13.94,16.24 12.61,16.37L10.38,18.6L8.96,17.18L11.19,14.95C11.06,13.62 11.64,12.4 12.52,11.52L16.98,7.06L18.39,8.47L15.27,11.6C15.47,10.27 15.08,8.87 14.2,8L19.36,2.72M5.21,10C4.71,10.81 4.35,11.7 4.15,12.65L2.74,11.24L1.32,12.65L5.21,16.54L9.1,12.65L7.68,11.24L6.28,12.65C6.18,11.85 6.22,11.05 6.4,10.27L5.21,10Z",
        Color = "#F43F5E",
        Description = "Reviews, sanitizes, or removes presentation artifacts with explicit scope and confirmation for destructive actions.",
        DetailedHelpText = "### Presentation Cleaner\nOpens cleanup and review actions for hidden slides, animations, transitions, notes, comments, unused masters and layouts, sticky notes, footers, repeated spaces, accessibility text, shape locks, ink, external data links, invisible or off-slide objects, document information, unprotected custom XML, and visible numbers. Destructive actions show their scope and require confirmation; PowerPoint Undo is started where supported. Shape cleanup and XML cleanup preserve oPE-owned objects and recognizable Efficient Elements or think-cell content. Convert Slides to Images creates a separate presentation and leaves the source unchanged. Anonymization and number sanitization have separate confirmations.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class CleanerFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - returns false as Cleaner requires dialog for action selection.
        /// Manual switch in MainSidebar.xaml.cs handles opening the Cleaner dialog.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            // Requires dialog for clean action selection
            // Handled by the command registry in MainSidebar.Commands.cs.
            return false;
        }

        public static bool Execute(PowerPointManager manager, CleanAction action)
        {
            if (manager == null) return false;

            if (RequiresConfirmation(action))
            {
                var result = System.Windows.Forms.MessageBox.Show(
                    GetConfirmationText(action),
                    "Presentation Cleaner",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                if (result != System.Windows.Forms.DialogResult.Yes) return false;

                try { manager.GetApplication().StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "CleanerFeature.StartNewUndoEntry"); }
            }

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
                case CleanAction.RemoveAlternativeTextAll: return RemoveAlternativeText(manager, false);
                case CleanAction.RemoveAlternativeTextSelected: return RemoveAlternativeText(manager, true);
                case CleanAction.UnlockShapesAll: return UnlockShapes(manager, false);
                case CleanAction.UnlockShapesSelected: return UnlockShapes(manager, true);
                case CleanAction.RemoveInkAll: return RemoveInk(manager, false);
                case CleanAction.RemoveInkSelected: return RemoveInk(manager, true);
                case CleanAction.BreakDataLinksAll: return BreakDataLinks(manager, false);
                case CleanAction.BreakDataLinksSelected: return BreakDataLinks(manager, true);
                case CleanAction.RemoveInvisibleOrOffSlideAll: return RemoveInvisibleOrOffSlide(manager, false);
                case CleanAction.RemoveInvisibleOrOffSlideSelected: return RemoveInvisibleOrOffSlide(manager, true);
                case CleanAction.RemoveDocumentInformation: return RemoveDocumentInformation(manager);
                case CleanAction.RemoveUnusedLayouts: return RemoveUnusedLayouts(manager);
                case CleanAction.RemoveNonOpeCustomXml: return RemoveNonOpeCustomXml(manager);
                case CleanAction.MakeInvisibleShapesVisibleAll: return MakeInvisibleShapesVisible(manager, false);
                case CleanAction.MakeInvisibleShapesVisibleSelected: return MakeInvisibleShapesVisible(manager, true);
                case CleanAction.SanitizeNumbersAll: return SanitizeNumbers(manager, false);
                case CleanAction.SanitizeNumbersSelected: return SanitizeNumbers(manager, true);
                default: return false;
            }
        }

        private static bool RequiresConfirmation(CleanAction action)
        {
            return oPenEfficiency.Core.Services.CleanerActionPolicy.RequiresConfirmation(action);
        }

        private static string GetConfirmationText(CleanAction action)
        {
            return oPenEfficiency.Core.Services.CleanerActionPolicy.GetConfirmationText(action);
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
                    PrivateCommentsService.DeleteForSlides(app.ActivePresentation, new[] { slide });
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
                var usedDesignNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (Slide slide in pres.Slides)
                {
                    if (slide.Design != null) usedDesignNames.Add(slide.Design.Name);
                }

                int deletedCount = 0;
                for (int i = pres.Designs.Count; i >= 1; i--)
                {
                    var design = pres.Designs[i];
                    if (!usedDesignNames.Contains(design.Name))
                    {
                        design.Delete();
                        deletedCount++;
                    }
                }

                System.Windows.Forms.MessageBox.Show(
                    deletedCount == 0 ? "No unused slide masters were found." : $"Removed {deletedCount} unused slide master(s).",
                    "Presentation Cleaner",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
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
                                if (StickyNoteFeature.IsStickyNote(shape))
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

        public static bool RemoveAlternativeText(PowerPointManager manager, bool selectedOnly)
        {
            return ProcessShapes(manager, selectedOnly, "CleanerFeature.RemoveAlternativeText", shape =>
            {
                if (IsProtectedShape(shape)) return 0;
                int changed = 0;
                if (!string.IsNullOrEmpty(shape.AlternativeText)) { shape.AlternativeText = string.Empty; changed = 1; }
                if (!string.IsNullOrEmpty(shape.Title)) { shape.Title = string.Empty; changed = 1; }
                return changed;
            }, "Removed accessibility text from {0} shape(s); protected smart objects were preserved.");
        }

        public static bool UnlockShapes(PowerPointManager manager, bool selectedOnly)
        {
            return ProcessShapes(manager, selectedOnly, "CleanerFeature.UnlockShapes", shape =>
            {
                if (IsProtectedShape(shape)) return 0;
                try
                {
                    dynamic candidate = shape;
                    if ((int)candidate.Locked == 0) return 0;
                    candidate.Locked = Office.MsoTriState.msoFalse;
                    ShapeMetadataService.RemoveTag(shape, "oPE_Locked");
                    return 1;
                }
                catch (System.Runtime.InteropServices.COMException) { return 0; }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { return 0; }
            }, "Unlocked {0} shape(s); protected third-party objects were preserved.");
        }

        public static bool RemoveInk(PowerPointManager manager, bool selectedOnly)
        {
            return ProcessTopLevelShapesReverse(manager, selectedOnly, "CleanerFeature.RemoveInk", shape =>
            {
                int type = (int)shape.Type;
                if (type != 22 && type != 23) return 0; // Office.MsoShapeType.msoInk / msoInkComment
                if (IsProtectedShape(shape)) return 0;
                shape.Delete();
                return 1;
            }, "Removed {0} ink object(s).");
        }

        public static bool BreakDataLinks(PowerPointManager manager, bool selectedOnly)
        {
            return ProcessShapes(manager, selectedOnly, "CleanerFeature.BreakDataLinks", shape =>
            {
                if (IsProtectedShape(shape)) return 0;
                if (shape.Type != Office.MsoShapeType.msoLinkedOLEObject &&
                    shape.Type != Office.MsoShapeType.msoLinkedPicture) return 0;
                shape.LinkFormat.BreakLink();
                return 1;
            }, "Converted {0} externally linked object(s) to local content.");
        }

        public static bool RemoveInvisibleOrOffSlide(PowerPointManager manager, bool selectedOnly)
        {
            var presentation = manager?.GetApplication()?.ActivePresentation;
            if (presentation == null) return false;
            float width = presentation.PageSetup.SlideWidth;
            float height = presentation.PageSetup.SlideHeight;
            return ProcessTopLevelShapesReverse(manager, selectedOnly, "CleanerFeature.RemoveInvisibleOrOffSlide", shape =>
            {
                if (IsProtectedShape(shape)) return 0;
                bool invisible = shape.Visible == Office.MsoTriState.msoFalse;
                bool offSlide = shape.Left + shape.Width <= 0 || shape.Top + shape.Height <= 0 ||
                                shape.Left >= width || shape.Top >= height;
                if (!invisible && !offSlide) return 0;
                shape.Delete();
                return 1;
            }, "Removed {0} invisible or fully off-slide object(s); protected smart objects were preserved.");
        }

        public static bool RemoveDocumentInformation(PowerPointManager manager)
        {
            try
            {
                dynamic presentation = manager?.GetApplication()?.ActivePresentation;
                if (presentation == null) return false;
                presentation.RemoveDocumentInformation(99); // PpRemoveDocInfoType.ppRDIAll
                ShowCleanerResult("Removed document properties and personal information.");
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.RemoveDocumentInformation");
                return false;
            }
        }

        public static bool RemoveUnusedLayouts(PowerPointManager manager)
        {
            try
            {
                var presentation = manager?.GetApplication()?.ActivePresentation;
                if (presentation == null) return false;
                var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (Slide slide in presentation.Slides)
                    used.Add((slide.Design?.Name ?? string.Empty) + "\n" + (slide.CustomLayout?.Name ?? string.Empty));

                int removed = 0;
                foreach (Design design in presentation.Designs)
                {
                    var layouts = design.SlideMaster.CustomLayouts;
                    for (int index = layouts.Count; index >= 1; index--)
                    {
                        if (layouts.Count <= 1) break;
                        var layout = layouts[index];
                        if (used.Contains((design.Name ?? string.Empty) + "\n" + (layout.Name ?? string.Empty))) continue;
                        layout.Delete();
                        removed++;
                    }
                }
                ShowCleanerResult($"Removed {removed} unused custom layout(s).");
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.RemoveUnusedLayouts");
                return false;
            }
        }

        public static bool RemoveNonOpeCustomXml(PowerPointManager manager)
        {
            try
            {
                var presentation = manager?.GetApplication()?.ActivePresentation;
                if (presentation == null) return false;
                int removed = 0;
                for (int index = presentation.CustomXMLParts.Count; index >= 1; index--)
                {
                    var part = presentation.CustomXMLParts[index];
                    string xml = part.XML ?? string.Empty;
                    if (xml.IndexOf("urn:oPenEfficiency:", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        xml.IndexOf("think-cell", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        xml.IndexOf("EE4P", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                    part.Delete();
                    removed++;
                }
                ShowCleanerResult($"Removed {removed} unprotected custom XML part(s).");
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CleanerFeature.RemoveNonOpeCustomXml");
                return false;
            }
        }

        public static bool MakeInvisibleShapesVisible(PowerPointManager manager, bool selectedOnly)
        {
            return ProcessShapes(manager, selectedOnly, "CleanerFeature.MakeInvisibleShapesVisible", shape =>
            {
                if (IsProtectedShape(shape) || shape.Visible != Office.MsoTriState.msoFalse) return 0;
                shape.Visible = Office.MsoTriState.msoTrue;
                return 1;
            }, "Made {0} invisible shape(s) visible.");
        }

        public static bool SanitizeNumbers(PowerPointManager manager, bool selectedOnly)
        {
            var random = new Random();
            return ProcessShapes(manager, selectedOnly, "CleanerFeature.SanitizeNumbers", shape =>
            {
                if (IsProtectedShape(shape)) return 0;
                int changed = 0;
                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    for (int row = 1; row <= shape.Table.Rows.Count; row++)
                        for (int column = 1; column <= shape.Table.Columns.Count; column++)
                            changed += SanitizeTextRange(shape.Table.Cell(row, column).Shape.TextFrame.TextRange, random);
                }
                else if (shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                         shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    changed += SanitizeTextRange(shape.TextFrame.TextRange, random);
                }
                return changed > 0 ? 1 : 0;
            }, "Sanitized visible numbers in {0} shape(s).");
        }

        private static int SanitizeTextRange(TextRange range, Random random)
        {
            int changed = 0;
            string text = range?.Text ?? string.Empty;
            for (int index = 1; index <= text.Length; index++)
            {
                if (!char.IsDigit(text[index - 1])) continue;
                range.Characters(index, 1).Text = random.Next(0, 10).ToString();
                changed++;
            }
            return changed;
        }

        private static bool IsProtectedShape(Shape shape)
        {
            return ShapeMetadataService.HasOpeTags(shape) || ShapeMetadataService.HasThirdPartyTags(shape);
        }

        private static bool ProcessShapes(
            PowerPointManager manager,
            bool selectedOnly,
            string context,
            Func<Shape, int> process,
            string resultFormat)
        {
            try
            {
                int changed = 0;
                foreach (Slide slide in GetTargetSlides(manager, selectedOnly))
                    for (int index = 1; index <= slide.Shapes.Count; index++)
                        changed += ProcessShapeRecursive(slide.Shapes[index], process);
                ShowCleanerResult(string.Format(resultFormat, changed));
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, context);
                return false;
            }
        }

        private static int ProcessShapeRecursive(Shape shape, Func<Shape, int> process)
        {
            int changed = process(shape);
            if (shape.Type != Office.MsoShapeType.msoGroup) return changed;
            for (int index = 1; index <= shape.GroupItems.Count; index++)
                changed += ProcessShapeRecursive(shape.GroupItems[index], process);
            return changed;
        }

        private static bool ProcessTopLevelShapesReverse(
            PowerPointManager manager,
            bool selectedOnly,
            string context,
            Func<Shape, int> process,
            string resultFormat)
        {
            try
            {
                int changed = 0;
                foreach (Slide slide in GetTargetSlides(manager, selectedOnly))
                    for (int index = slide.Shapes.Count; index >= 1; index--)
                        changed += process(slide.Shapes[index]);
                ShowCleanerResult(string.Format(resultFormat, changed));
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, context);
                return false;
            }
        }

        private static void ShowCleanerResult(string message)
        {
            System.Windows.Forms.MessageBox.Show(
                message,
                "Presentation Cleaner",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
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
                            catch { /* A stale shape is skipped while cleaning the remaining slide content. */ }
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
            string tempFolder = null;
            string outputPath = null;
            Presentation newPres = null;
            bool completed = false;
            try
            {
                var app = manager.GetApplication();
                var pres = app.ActivePresentation;
                var slides = GetTargetSlides(manager, selectedOnly);
                var slideList = new List<Slide>();
                foreach (Slide s in slides) slideList.Add(s);
                if (slideList.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "No slides are available in the selected scope.",
                        "Convert Slides to Images",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }

                tempFolder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "oPenEfficiency_SlideImages_" + Guid.NewGuid().ToString("N"));
                System.IO.Directory.CreateDirectory(tempFolder);
                newPres = app.Presentations.Add(Office.MsoTriState.msoFalse);
                newPres.PageSetup.SlideWidth = pres.PageSetup.SlideWidth;
                newPres.PageSetup.SlideHeight = pres.PageSetup.SlideHeight;
                int processedCount = 0;

                foreach (Slide slide in slideList)
                {
                    int slideIndex = slide.SlideNumber;
                    string imagePath = System.IO.Path.Combine(tempFolder, $"slide_{processedCount + 1:D4}_{slideIndex}.png");
                    slide.Export(imagePath, "PNG", (int)pres.PageSetup.SlideWidth * 3, (int)pres.PageSetup.SlideHeight * 3);

                    var newSlide = newPres.Slides.Add(newPres.Slides.Count + 1, PpSlideLayout.ppLayoutBlank);
                    newSlide.Shapes.AddPicture(
                        imagePath,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue,
                        0,
                        0,
                        pres.PageSetup.SlideWidth,
                        pres.PageSetup.SlideHeight);
                    processedCount++;
                }

                string originalDir = string.IsNullOrWhiteSpace(pres.Path)
                    ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    : pres.Path;
                string originalName = System.IO.Path.GetFileNameWithoutExtension(pres.Name);
                if (string.IsNullOrWhiteSpace(originalName)) originalName = "Presentation";
                outputPath = System.IO.Path.Combine(originalDir, $"{originalName}_UNEDITABLE.pptx");
                if (System.IO.File.Exists(outputPath))
                {
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    outputPath = System.IO.Path.Combine(originalDir, $"{originalName}_UNEDITABLE_{timestamp}.pptx");
                }

                newPres.SaveAs(outputPath);
                completed = true;

                System.Windows.Forms.MessageBox.Show(
                    $"Created uneditable presentation with {processedCount} slide(s) converted to images.\n\n" +
                    $"Saved to:\n{outputPath}\n\n" +
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
            finally
            {
                if (!completed && newPres != null)
                {
                    try { newPres.Close(); }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "CleanerFeature.ConvertSlidesToImages.CloseFailedPresentation"); }
                }

                if (!completed && !string.IsNullOrWhiteSpace(outputPath) && System.IO.File.Exists(outputPath))
                {
                    try { System.IO.File.Delete(outputPath); }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "CleanerFeature.ConvertSlidesToImages.DeletePartialOutput"); }
                }

                if (!string.IsNullOrWhiteSpace(tempFolder) && System.IO.Directory.Exists(tempFolder))
                {
                    try { System.IO.Directory.Delete(tempFolder, true); }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "CleanerFeature.ConvertSlidesToImages.Cleanup"); }
                }
            }
        }
    }
}
