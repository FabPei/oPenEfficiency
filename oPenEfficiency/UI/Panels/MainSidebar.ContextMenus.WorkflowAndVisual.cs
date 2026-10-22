using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using oPenEfficiency.Features;
using oPenEfficiency.Features.Charts;
using oPenEfficiency.Features.Tables;
using oPenEfficiency.Features.Text;
using oPenEfficiency.Features.Utilities;
using oPenEfficiency.Features.Visuals;
using oPenEfficiency.Models;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Commands;
using oPenEfficiency.UI.Dialogs;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency
{
    public partial class MainSidebar
    {
        private void AttachWorkflowAndVisualContextMenus(ButtonBase btn)
        {
            if (btn.Name == "BtnDeleteText")
            {
                if (btn.ContextMenu.Items.Count > 0)
                    btn.ContextMenu.Items.Add(new Separator());

                var deleteAll = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 0, "Delete All Text (Keep Shapes)")
                };
                deleteAll.Click += (s, e) => DeleteTextFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(deleteAll);

                var deleteHighlighted = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 1, "Delete Highlighted Text")
                };
                deleteHighlighted.Click += (s, e) => DeleteTextFeature.DeleteHighlightedText(GetManager());
                btn.ContextMenu.Items.Add(deleteHighlighted);

                var deleteParagraphs = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 2, "Delete Selected Paragraphs")
                };
                deleteParagraphs.Click += (s, e) => DeleteTextFeature.DeleteSelectedParagraphs(GetManager());
                btn.ContextMenu.Items.Add(deleteParagraphs);

                btn.ContextMenu.Opened += (s, e) =>
                {
                    bool hasHighlightedText = DeleteTextFeature.HasHighlightedText(GetManager());
                    deleteHighlighted.IsEnabled = hasHighlightedText;
                    deleteParagraphs.IsEnabled = hasHighlightedText;
                };

                btn.ContextMenu.Items.Add(new Separator());
                var replaceText = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 3, "Replace Selected Text...")
                };
                replaceText.Click += (s, e) =>
                    ToggleFloatingWindow(
                        "BtnApplyTextTool",
                        () => new UI.ApplyTextWindow(GetManager()),
                        null);
                btn.ContextMenu.Items.Add(replaceText);
            }
            else if (btn.Name == "BtnTextFormatPainter")
            {
                if (btn.ContextMenu.Items.Count > 0)
                    btn.ContextMenu.Items.Add(new Separator());

                var typographyOnly = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 0, "Apply Typography Only")
                };
                typographyOnly.Click += (s, e) =>
                    TextFormatPainterFeature.ExecuteTypographyOnly(GetManager());
                btn.ContextMenu.Items.Add(typographyOnly);

                var paragraphsAndBulletsOnly = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 1, "Apply Paragraphs & Bullets Only")
                };
                paragraphsAndBulletsOnly.Click += (s, e) =>
                    TextFormatPainterFeature.ExecuteParagraphsAndBulletsOnly(GetManager());
                btn.ContextMenu.Items.Add(paragraphsAndBulletsOnly);

                var preserveEmphasis = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 2, "Apply While Preserving Emphasis")
                };
                preserveEmphasis.Click += (s, e) =>
                    TextFormatPainterFeature.ExecutePreservingEmphasis(GetManager());
                btn.ContextMenu.Items.Add(preserveEmphasis);
            }
            else if (btn.Name == "BtnSwapTextPlain")
            {
                if (btn.ContextMenu.Items.Count > 0)
                    btn.ContextMenu.Items.Add(new Separator());

                var shrinkOverflow = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 0, "Swap Text & Shrink on Overflow")
                };
                shrinkOverflow.Click += (s, e) => SwapTextPlainFeature.ExecuteSwapAndShrink(GetManager());
                btn.ContextMenu.Items.Add(shrinkOverflow);

                var cycleText = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 1, "Cycle Text Through Selection")
                };
                cycleText.Click += (s, e) => SwapTextPlainFeature.CycleTextThroughSelection(GetManager());
                btn.ContextMenu.Items.Add(cycleText);

                var swapParagraphs = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 2, "Swap Selected Paragraphs Only...")
                };
                swapParagraphs.Click += (s, e) => SwapTextPlainFeature.SwapSelectedParagraphs(GetManager());
                btn.ContextMenu.Items.Add(swapParagraphs);
            }
            else if (btn.Name == "BtnTrafficLight")
            {
                if (btn.ContextMenu.Items.Count > 0)
                    btn.ContextMenu.Items.Add(new Separator());

                var trafficLightStates = new[]
                {
                    (GetCatalogVariantLabel(btn.Name, 0, "Set / Insert Red"), TrafficLightFeature.TrafficLightState.Red),
                    (GetCatalogVariantLabel(btn.Name, 1, "Set / Insert Yellow"), TrafficLightFeature.TrafficLightState.Yellow),
                    (GetCatalogVariantLabel(btn.Name, 2, "Set / Insert Green"), TrafficLightFeature.TrafficLightState.Green),
                    (GetCatalogVariantLabel(btn.Name, 3, "Set / Insert Off"), TrafficLightFeature.TrafficLightState.Off)
                };

                foreach (var stateOption in trafficLightStates)
                {
                    var item = new MenuItem { Header = stateOption.Item1 };
                    var state = stateOption.Item2;
                    item.Click += (s, e) => TrafficLightFeature.ExecuteState(GetManager(), state);
                    btn.ContextMenu.Items.Add(item);
                }
            }
            else if (btn.Name == "BtnSwapPositions")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var anchors = new[] {
                    (GetCatalogVariantLabel(btn.Name, 0, "Top Left"), oPenEfficiency.Core.Services.AnchorSwapPlacementPolicy.SwapAnchor.TopLeft),
                    (GetCatalogVariantLabel(btn.Name, 1, "Top Center"), oPenEfficiency.Core.Services.AnchorSwapPlacementPolicy.SwapAnchor.Top),
                    (GetCatalogVariantLabel(btn.Name, 2, "Top Right"), oPenEfficiency.Core.Services.AnchorSwapPlacementPolicy.SwapAnchor.TopRight),
                    (GetCatalogVariantLabel(btn.Name, 3, "Left Center"), oPenEfficiency.Core.Services.AnchorSwapPlacementPolicy.SwapAnchor.Left),
                    (GetCatalogVariantLabel(btn.Name, 4, "Center (Default)"), oPenEfficiency.Core.Services.AnchorSwapPlacementPolicy.SwapAnchor.Center),
                    (GetCatalogVariantLabel(btn.Name, 5, "Right Center"), oPenEfficiency.Core.Services.AnchorSwapPlacementPolicy.SwapAnchor.Right),
                    (GetCatalogVariantLabel(btn.Name, 6, "Bottom Left"), oPenEfficiency.Core.Services.AnchorSwapPlacementPolicy.SwapAnchor.BottomLeft),
                    (GetCatalogVariantLabel(btn.Name, 7, "Bottom Center"), oPenEfficiency.Core.Services.AnchorSwapPlacementPolicy.SwapAnchor.Bottom),
                    (GetCatalogVariantLabel(btn.Name, 8, "Bottom Right"), oPenEfficiency.Core.Services.AnchorSwapPlacementPolicy.SwapAnchor.BottomRight)
                };

                foreach (var anchorInfo in anchors)
                {
                    var mi = new MenuItem { Header = anchorInfo.Item1 };
                    var anchorVal = anchorInfo.Item2;
                    mi.Click += (s, e) => SwapPositionsFeature.Execute(GetManager(), anchorVal);
                    btn.ContextMenu.Items.Add(mi);
                }
            }
            else if (btn.Name == "BtnStoryline")
            {
                if (btn.ContextMenu.Items.Count > 0)
                    btn.ContextMenu.Items.Add(new Separator());

                var openSection = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Open Current Section") };
                openSection.Click += (s, e) => StorylineFeature.Open(GetManager(), true, false);
                btn.ContextMenu.Items.Add(openSection);

                var refresh = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Refresh Storyline") };
                refresh.Click += (s, e) => StorylineFeature.RefreshOpenWindow();
                btn.ContextMenu.Items.Add(refresh);

                btn.ContextMenu.Items.Add(new Separator());
                var check = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Check Headline Quality") };
                check.Click += (s, e) => StorylineFeature.Open(GetManager(), false, true);
                btn.ContextMenu.Items.Add(check);

                var missing = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Show Missing Titles") };
                missing.Click += (s, e) => StorylineFeature.OpenMissingTitles(GetManager());
                btn.ContextMenu.Items.Add(missing);

                btn.ContextMenu.Items.Add(new Separator());
                var copyNumbered = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 4, "Copy Numbered Storyline") };
                copyNumbered.Click += (s, e) =>
                {
                    var titles = StorylineFeature.ExtractStoryline(GetManager());
                    System.Windows.Clipboard.SetText(StorylineFeature.BuildOutline(titles, true, false, false));
                };
                btn.ContextMenu.Items.Add(copyNumbered);

                var copyHeadlines = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 5, "Copy Headlines Only") };
                copyHeadlines.Click += (s, e) =>
                {
                    var titles = StorylineFeature.ExtractStoryline(GetManager());
                    System.Windows.Clipboard.SetText(StorylineFeature.BuildOutline(titles, false, false, false));
                };
                btn.ContextMenu.Items.Add(copyHeadlines);

                var copySections = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 6, "Copy Grouped by Sections") };
                copySections.Click += (s, e) =>
                {
                    var titles = StorylineFeature.ExtractStoryline(GetManager());
                    System.Windows.Clipboard.SetText(StorylineFeature.BuildOutline(titles, true, true, false));
                };
                btn.ContextMenu.Items.Add(copySections);

                var copyIssues = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 7, "Copy Issues Report") };
                copyIssues.Click += (s, e) =>
                {
                    var titles = StorylineFeature.ExtractStoryline(GetManager());
                    System.Windows.Clipboard.SetText(StorylineFeature.BuildOutline(titles, true, true, true));
                };
                btn.ContextMenu.Items.Add(copyIssues);
            }
            else if (btn.Name == "BtnAgendaWizard")
            {
                if (btn.ContextMenu.Items.Count > 0)
                    btn.ContextMenu.Items.Add(new Separator());

                var updateAgenda = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 0, "Update Existing Agenda")
                };
                updateAgenda.Click += (s, e) =>
                    AgendaWizardFeature.Execute(GetManager(), AgendaWizardFeature.LaunchAction.Open);
                btn.ContextMenu.Items.Add(updateAgenda);

                var buildStructure = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 1, "Build Deck Structure")
                };
                buildStructure.Click += (s, e) =>
                    AgendaWizardFeature.Execute(GetManager(), AgendaWizardFeature.LaunchAction.PrepareDeckStructure);
                btn.ContextMenu.Items.Add(buildStructure);

                var extractFromDeck = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 2, "Extract Agenda from Deck")
                };
                extractFromDeck.Click += (s, e) =>
                    AgendaWizardFeature.Execute(GetManager(), AgendaWizardFeature.LaunchAction.ExtractFromDeck);
                btn.ContextMenu.Items.Add(extractFromDeck);

                var validate = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 3, "Validate Agenda & Sections")
                };
                validate.Click += (s, e) =>
                    AgendaWizardFeature.Execute(GetManager(), AgendaWizardFeature.LaunchAction.Validate);
                btn.ContextMenu.Items.Add(validate);

                btn.ContextMenu.Items.Add(new Separator());
                var openDesigner = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 4, "Open Agenda Layout Designer")
                };
                openDesigner.Click += (s, e) => SaveAgendaLayoutFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(openDesigner);
            }
            else if (btn.Name == "BtnStyleCheck")
            {
                if (btn.ContextMenu.Items.Count > 0)
                    btn.ContextMenu.Items.Add(new Separator());

                var currentSlide = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 0, "Check Current Slide")
                };
                currentSlide.Click += (s, e) =>
                    StyleCheckFeature.ExecuteDialog(
                        GetManager(),
                        null,
                        StyleCheckScope.CurrentSlide);
                btn.ContextMenu.Items.Add(currentSlide);

                var selectedSlides = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 1, "Check Selected Slides")
                };
                selectedSlides.Click += (s, e) =>
                    StyleCheckFeature.ExecuteDialog(
                        GetManager(),
                        null,
                        StyleCheckScope.SelectedSlides);
                btn.ContextMenu.Items.Add(selectedSlides);

                var entirePresentation = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 2, "Check Entire Presentation")
                };
                entirePresentation.Click += (s, e) =>
                    StyleCheckFeature.ExecuteDialog(
                        GetManager(),
                        null,
                        StyleCheckScope.EntirePresentation);
                btn.ContextMenu.Items.Add(entirePresentation);
            }
            else if (btn.Name == "BtnPasteXY")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var clear = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Clear Copied XY Coordinates") };
                clear.Click += (s, e) => PasteCoordinatesFeature.ClearCopiedGeometry();
                btn.ContextMenu.Items.Add(clear);
            }
            else if (btn.Name == "BtnSlideGuidelines")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var mi1 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Three Columns") };
                mi1.Click += (s, e) => SlideGuidelinesFeature.Execute(GetManager(), SlideGuidelinesFeature.GuidelineLayout.ThreeColumn);
                btn.ContextMenu.Items.Add(mi1);

                var mi2 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Four Columns") };
                mi2.Click += (s, e) => SlideGuidelinesFeature.Execute(GetManager(), SlideGuidelinesFeature.GuidelineLayout.FourColumn);
                btn.ContextMenu.Items.Add(mi2);

                var mi3 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Five Columns") };
                mi3.Click += (s, e) => SlideGuidelinesFeature.Execute(GetManager(), SlideGuidelinesFeature.GuidelineLayout.FiveColumn);
                btn.ContextMenu.Items.Add(mi3);

                var mi4 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "1:2 Columns") };
                mi4.Click += (s, e) => SlideGuidelinesFeature.Execute(GetManager(), SlideGuidelinesFeature.GuidelineLayout.OneTwoColumn);
                btn.ContextMenu.Items.Add(mi4);

                btn.ContextMenu.Items.Add(new Separator());
                var miThirds = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 4, "Rule of Thirds") };
                miThirds.Click += (s, e) => SlideGuidelinesFeature.Execute(GetManager(), SlideGuidelinesFeature.GuidelineLayout.RuleOfThirds);
                btn.ContextMenu.Items.Add(miThirds);

                btn.ContextMenu.Items.Add(new Separator());
                var miDefault = new MenuItem { Header = "2-Column Layout" };
                miDefault.Click += (s, e) => SlideGuidelinesFeature.Execute(GetManager(), SlideGuidelinesFeature.GuidelineLayout.Default);
                btn.ContextMenu.Items.Add(miDefault);

                btn.ContextMenu.Items.Add(new Separator());
                var miRemove = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 5, "Remove Guidelines") };
                miRemove.Click += (s, e) => SlideGuidelinesFeature.RemoveGuidelines(GetManager());
                btn.ContextMenu.Items.Add(miRemove);
            }
            else if (btn.Name == "BtnCleaner")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var miSlides = new MenuItem { Header = "Remove Hidden Slides" };
                miSlides.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveHiddenSlides);
                btn.ContextMenu.Items.Add(miSlides);

                var miAnim = new MenuItem { Header = "Remove Animations" };
                var miAnimAll = new MenuItem { Header = "From All Slides" };
                miAnimAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveAnimationsAll);
                var miAnimSel = new MenuItem { Header = "From Selected Slides" };
                miAnimSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveAnimationsSelected);
                miAnim.Items.Add(miAnimAll);
                miAnim.Items.Add(miAnimSel);
                btn.ContextMenu.Items.Add(miAnim);

                var miTrans = new MenuItem { Header = "Remove Transitions" };
                var miTransAll = new MenuItem { Header = "From All Slides" };
                miTransAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveTransitionsAll);
                var miTransSel = new MenuItem { Header = "From Selected Slides" };
                miTransSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveTransitionsSelected);
                miTrans.Items.Add(miTransAll);
                miTrans.Items.Add(miTransSel);
                btn.ContextMenu.Items.Add(miTrans);

                var miNotes = new MenuItem { Header = "Remove Speaker Notes" };
                var miNotesAll = new MenuItem { Header = "From All Slides" };
                miNotesAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveSpeakerNotesAll);
                var miNotesSel = new MenuItem { Header = "From Selected Slides" };
                miNotesSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveSpeakerNotesSelected);
                miNotes.Items.Add(miNotesAll);
                miNotes.Items.Add(miNotesSel);
                btn.ContextMenu.Items.Add(miNotes);

                var miDevNotes = new MenuItem { Header = "Remove Developer Notes" };
                var miDevNotesAll = new MenuItem { Header = "From All Slides" };
                miDevNotesAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveDeveloperNotesAll);
                var miDevNotesSel = new MenuItem { Header = "From Selected Slides" };
                miDevNotesSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveDeveloperNotesSelected);
                miDevNotes.Items.Add(miDevNotesAll);
                miDevNotes.Items.Add(miDevNotesSel);
                btn.ContextMenu.Items.Add(miDevNotes);

                var miComm = new MenuItem { Header = "Remove Comments" };
                var miCommAll = new MenuItem { Header = "From All Slides" };
                miCommAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveCommentsAll);
                var miCommSel = new MenuItem { Header = "From Selected Slides" };
                miCommSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveCommentsSelected);
                miComm.Items.Add(miCommAll);
                miComm.Items.Add(miCommSel);
                btn.ContextMenu.Items.Add(miComm);

                var miAccessibility = new MenuItem { Header = "Remove Alternative Text" };
                var miAccessibilityAll = new MenuItem { Header = "From All Slides" };
                miAccessibilityAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveAlternativeTextAll);
                var miAccessibilitySel = new MenuItem { Header = "From Selected Slides" };
                miAccessibilitySel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveAlternativeTextSelected);
                miAccessibility.Items.Add(miAccessibilityAll);
                miAccessibility.Items.Add(miAccessibilitySel);
                btn.ContextMenu.Items.Add(miAccessibility);

                var miUnlock = new MenuItem { Header = "Unlock Shapes" };
                var miUnlockAll = new MenuItem { Header = "On All Slides" };
                miUnlockAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.UnlockShapesAll);
                var miUnlockSel = new MenuItem { Header = "On Selected Slides" };
                miUnlockSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.UnlockShapesSelected);
                miUnlock.Items.Add(miUnlockAll);
                miUnlock.Items.Add(miUnlockSel);
                btn.ContextMenu.Items.Add(miUnlock);

                var miInk = new MenuItem { Header = "Remove Ink" };
                var miInkAll = new MenuItem { Header = "From All Slides" };
                miInkAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveInkAll);
                var miInkSel = new MenuItem { Header = "From Selected Slides" };
                miInkSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveInkSelected);
                miInk.Items.Add(miInkAll);
                miInk.Items.Add(miInkSel);
                btn.ContextMenu.Items.Add(miInk);

                var miLinks = new MenuItem { Header = "Break External Data Links" };
                var miLinksAll = new MenuItem { Header = "On All Slides" };
                miLinksAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.BreakDataLinksAll);
                var miLinksSel = new MenuItem { Header = "On Selected Slides" };
                miLinksSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.BreakDataLinksSelected);
                miLinks.Items.Add(miLinksAll);
                miLinks.Items.Add(miLinksSel);
                btn.ContextMenu.Items.Add(miLinks);

                var miInvisible = new MenuItem { Header = "Invisible & Off-Slide Objects" };
                var miRemoveInvisibleAll = new MenuItem { Header = "Remove from All Slides" };
                miRemoveInvisibleAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveInvisibleOrOffSlideAll);
                var miRemoveInvisibleSel = new MenuItem { Header = "Remove from Selected Slides" };
                miRemoveInvisibleSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveInvisibleOrOffSlideSelected);
                var miShowInvisibleAll = new MenuItem { Header = "Make Visible on All Slides" };
                miShowInvisibleAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.MakeInvisibleShapesVisibleAll);
                var miShowInvisibleSel = new MenuItem { Header = "Make Visible on Selected Slides" };
                miShowInvisibleSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.MakeInvisibleShapesVisibleSelected);
                miInvisible.Items.Add(miRemoveInvisibleAll);
                miInvisible.Items.Add(miRemoveInvisibleSel);
                miInvisible.Items.Add(new Separator());
                miInvisible.Items.Add(miShowInvisibleAll);
                miInvisible.Items.Add(miShowInvisibleSel);
                btn.ContextMenu.Items.Add(miInvisible);

                var miSanitizeNumbers = new MenuItem { Header = "Sanitize Visible Numbers" };
                var miSanitizeNumbersAll = new MenuItem { Header = "On All Slides" };
                miSanitizeNumbersAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.SanitizeNumbersAll);
                var miSanitizeNumbersSel = new MenuItem { Header = "On Selected Slides" };
                miSanitizeNumbersSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.SanitizeNumbersSelected);
                miSanitizeNumbers.Items.Add(miSanitizeNumbersAll);
                miSanitizeNumbers.Items.Add(miSanitizeNumbersSel);
                btn.ContextMenu.Items.Add(miSanitizeNumbers);

                var miDocumentInfo = new MenuItem { Header = "Remove Document Information" };
                miDocumentInfo.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveDocumentInformation);
                btn.ContextMenu.Items.Add(miDocumentInfo);

                var miCustomXml = new MenuItem { Header = "Remove Non-oPE Custom XML" };
                miCustomXml.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveNonOpeCustomXml);
                btn.ContextMenu.Items.Add(miCustomXml);

                var miMasters = new MenuItem { Header = "Remove Unused Masters" };
                miMasters.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveUnusedMasters);
                btn.ContextMenu.Items.Add(miMasters);

                var miLayouts = new MenuItem { Header = "Remove Unused Layouts" };
                miLayouts.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveUnusedLayouts);
                btn.ContextMenu.Items.Add(miLayouts);

                var miSticky = new MenuItem { Header = "Delete All Sticky Notes" };
                miSticky.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.DeleteStickyNotes);
                btn.ContextMenu.Items.Add(miSticky);

                var miAnon = new MenuItem { Header = "Anonymize Selected Slide Text" };
                miAnon.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.AnonymizeSelected);
                btn.ContextMenu.Items.Add(miAnon);

                var miFoot = new MenuItem { Header = "Remove Unedited Footers" };
                var miFootAll = new MenuItem { Header = "From All Slides" };
                miFootAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveUneditedFootersAll);
                var miFootSel = new MenuItem { Header = "From Selected Slides" };
                miFootSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveUneditedFootersSelected);
                miFoot.Items.Add(miFootAll);
                miFoot.Items.Add(miFootSel);
                btn.ContextMenu.Items.Add(miFoot);

                var miDoubleSpace = new MenuItem { Header = "Remove Double Spaces" };
                var miDoubleSpaceAll = new MenuItem { Header = "From All Slides" };
                miDoubleSpaceAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveDoubleSpacesAll);
                var miDoubleSpaceSel = new MenuItem { Header = "From Selected Slides" };
                miDoubleSpaceSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.RemoveDoubleSpacesSelected);
                miDoubleSpace.Items.Add(miDoubleSpaceAll);
                miDoubleSpace.Items.Add(miDoubleSpaceSel);
                btn.ContextMenu.Items.Add(miDoubleSpace);

                btn.ContextMenu.Items.Add(new Separator());
                var miConvert = new MenuItem { Header = "Convert Slides to Images" };
                var miConvertAll = new MenuItem { Header = "All Slides" };
                miConvertAll.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.ConvertSlidesToImagesAll);
                var miConvertSel = new MenuItem { Header = "Selected Slides" };
                miConvertSel.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.ConvertSlidesToImagesSelected);
                miConvert.Items.Add(miConvertAll);
                miConvert.Items.Add(miConvertSel);
                btn.ContextMenu.Items.Add(miConvert);

                btn.ContextMenu.Items.Add(new Separator());
                var miSuper = new MenuItem { Header = "Detect Superscript Numbers Without Footer" };
                miSuper.Click += (s, e) => CleanerFeature.Execute(GetManager(), oPenEfficiency.Core.Services.CleanerActionPolicy.CleanerAction.DetectSuperscriptWithoutFooter);
                btn.ContextMenu.Items.Add(miSuper);
            }
            else if (btn.Name == "BtnArrangeCircle")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var shapes = new[] {
                    ("Circle", "circle"),
                    (GetCatalogVariantLabel(btn.Name, 0, "Triangle"), "triangle"),
                    (GetCatalogVariantLabel(btn.Name, 1, "Square"), "square"),
                    (GetCatalogVariantLabel(btn.Name, 2, "Star"), "star"),
                    (GetCatalogVariantLabel(btn.Name, 3, "Octagon"), "octagon"),
                    (GetCatalogVariantLabel(btn.Name, 4, "Pentagon"), "pentagon")
                };
                foreach (var shape in shapes) {
                    var mi = new MenuItem { Header = shape.Item1 };
                    mi.Click += (s, e) => ArrangeInShapeFeature.Execute(GetManager(), shape.Item2);
                    btn.ContextMenu.Items.Add(mi);
                }
                btn.ContextMenu.Items.Add(new Separator());
                var miPro = new MenuItem { Header = "Pro..." };
                miPro.Click += (s, e) => ArrangeProFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(miPro);
            }
            else if (btn.Name == "BtnShapeLock")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var miT = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Toggle Lock All Shapes") };
                miT.Click += (s, e) => ShapeLockingFeature.SetGlobalShapeLock(GetManager(), null);
                btn.ContextMenu.Items.Add(miT);
                var miL = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Lock All Shapes") };
                miL.Click += (s, e) => ShapeLockingFeature.SetGlobalShapeLock(GetManager(), true);
                btn.ContextMenu.Items.Add(miL);
                var miU = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Unlock All Shapes") };
                miU.Click += (s, e) => ShapeLockingFeature.SetGlobalShapeLock(GetManager(), false);
                btn.ContextMenu.Items.Add(miU);
            }
            else if (btn.Name == "BtnShowHidden")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var showTagged = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Show openEfficiency-Hidden Shapes") };
                showTagged.Click += (s, e) => VisibilityFeature.ShowAllHiddenShapes(GetManager());
                btn.ContextMenu.Items.Add(showTagged);
                btn.ContextMenu.Items.Add(new Separator());
                var revealAll = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Reveal All Hidden Shapes on Slide...") };
                revealAll.Click += (s, e) =>
                {
                    if (System.Windows.MessageBox.Show(
                        "Reveal every hidden shape on the current slide, including shapes hidden outside openEfficiency?",
                        "Reveal All Hidden Shapes", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                    {
                        VisibilityFeature.ShowAllHiddenShapes(GetManager(), true);
                    }
                };
                btn.ContextMenu.Items.Add(revealAll);
            }
            else if (btn.Name == "BtnLockDimensions")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var miT = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Lock/Unlock Aspect Ratio (All Slides)") };
                miT.Click += (s, e) => DimensionLockFeature.SetGlobalDimensionLock(GetManager(), null);
                btn.ContextMenu.Items.Add(miT);
                var miL = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Lock Aspect Ratio (All Slides)") };
                miL.Click += (s, e) => DimensionLockFeature.SetGlobalDimensionLock(GetManager(), true);
                btn.ContextMenu.Items.Add(miL);
                var miU = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Unlock Aspect Ratio (All Slides)") };
                miU.Click += (s, e) => DimensionLockFeature.SetGlobalDimensionLock(GetManager(), false);
                btn.ContextMenu.Items.Add(miU);
            }
            else if (btn.Name == "BtnAddSticker")
            {
                if (btn.ContextMenu.Items.Count > 0)
                    btn.ContextMenu.Items.Add(new Separator());

                var draft = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Insert Draft Sticker") };
                draft.Click += (s, e) => AddStickerFeature.ExecutePreset(GetManager(), "DRAFT", 0x0B9EF5);
                btn.ContextMenu.Items.Add(draft);
                var review = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Insert Review Sticker") };
                review.Click += (s, e) => AddStickerFeature.ExecutePreset(GetManager(), "REVIEW", 0xF6823B);
                btn.ContextMenu.Items.Add(review);
                var approved = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Insert Approved Sticker") };
                approved.Click += (s, e) => AddStickerFeature.ExecutePreset(GetManager(), "APPROVED", 0x5EC522);
                btn.ContextMenu.Items.Add(approved);
                var confidential = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Insert Confidential Sticker") };
                confidential.Click += (s, e) => AddStickerFeature.ExecutePreset(GetManager(), "CONFIDENTIAL", 0x4444EF);
                btn.ContextMenu.Items.Add(confidential);
            }
            else if (btn.Name == "BtnIllustrativeSticker")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var presets = new MenuItem { Header = "Preset labels" };
                foreach (var label in IllustrativeStickerFeature.GetPresetLabels())
                {
                    string capturedLabel = label;
                    AddStatusStickerMenuItem(presets, capturedLabel,
                        () => IllustrativeStickerFeature.ApplyPreset(GetManager(), capturedLabel, StatusStickerScope.CurrentSlide));
                }
                btn.ContextMenu.Items.Add(presets);

                AddStatusStickerMenuItem(btn.ContextMenu, "Custom label...", () =>
                {
                    string text = PromptForStatusStickerText();
                    if (!string.IsNullOrWhiteSpace(text))
                        IllustrativeStickerFeature.ApplyPreset(GetManager(), text, StatusStickerScope.CurrentSlide);
                });

                var workflow = new MenuItem { Header = "Workflow status" };
                foreach (string status in new[] { "Draft", "Work in progress", "For discussion", "For decision", "Client-ready", "Updated" })
                {
                    string capturedStatus = status;
                    AddStatusStickerMenuItem(workflow, capturedStatus,
                        () => IllustrativeStickerFeature.ApplyDimension(GetManager(), workStatus: capturedStatus));
                }
                AddStatusStickerMenuItem(workflow, "Clear workflow status",
                    () => IllustrativeStickerFeature.ApplyDimension(GetManager(), workStatus: ""));
                btn.ContextMenu.Items.Add(workflow);

                var confidentiality = new MenuItem { Header = "Confidentiality" };
                AddStatusStickerMenuItem(confidentiality, "Confidential",
                    () => IllustrativeStickerFeature.ApplyDimension(GetManager(), confidentiality: "Confidential"));
                AddStatusStickerMenuItem(confidentiality, "Strictly confidential",
                    () => IllustrativeStickerFeature.ApplyDimension(GetManager(), confidentiality: "Strictly confidential"));
                AddStatusStickerMenuItem(confidentiality, "Clear confidentiality",
                    () => IllustrativeStickerFeature.ApplyDimension(GetManager(), confidentiality: ""));
                btn.ContextMenu.Items.Add(confidentiality);

                var audience = new MenuItem { Header = "Audience" };
                AddStatusStickerMenuItem(audience, "Internal",
                    () => IllustrativeStickerFeature.ApplyDimension(GetManager(), audience: "Internal"));
                AddStatusStickerMenuItem(audience, "Client",
                    () => IllustrativeStickerFeature.ApplyDimension(GetManager(), audience: "Client"));
                AddStatusStickerMenuItem(audience, "Clear audience",
                    () => IllustrativeStickerFeature.ApplyDimension(GetManager(), audience: ""));
                btn.ContextMenu.Items.Add(audience);

                btn.ContextMenu.Items.Add(new Separator());
                AddStatusStickerMenuItem(btn.ContextMenu, "Update sticker on current slide",
                    () => IllustrativeStickerFeature.RefreshCurrent(GetManager()));

                var apply = new MenuItem { Header = "Apply current sticker to" };
                AddStatusStickerMenuItem(apply, "Selected slides",
                    () => IllustrativeStickerFeature.CopyCurrentToScope(GetManager(), StatusStickerScope.SelectedSlides));
                AddStatusStickerMenuItem(apply, "Entire presentation",
                    () => IllustrativeStickerFeature.CopyCurrentToScope(GetManager(), StatusStickerScope.EntirePresentation));
                btn.ContextMenu.Items.Add(apply);

                var remove = new MenuItem { Header = "Remove status stickers" };
                AddStatusStickerMenuItem(remove, "Current slide",
                    () => IllustrativeStickerFeature.Remove(GetManager(), StatusStickerScope.CurrentSlide));
                AddStatusStickerMenuItem(remove, "Selected slides",
                    () => IllustrativeStickerFeature.Remove(GetManager(), StatusStickerScope.SelectedSlides));
                AddStatusStickerMenuItem(remove, "Entire presentation",
                    () => IllustrativeStickerFeature.Remove(GetManager(), StatusStickerScope.EntirePresentation));
                btn.ContextMenu.Items.Add(remove);

                var position = new MenuItem { Header = "Position" };
                AddStatusStickerMenuItem(position, "Top left",
                    () => IllustrativeStickerFeature.SetPosition(GetManager(), StatusStickerPosition.TopLeft));
                AddStatusStickerMenuItem(position, "Top right",
                    () => IllustrativeStickerFeature.SetPosition(GetManager(), StatusStickerPosition.TopRight));
                AddStatusStickerMenuItem(position, "Bottom left",
                    () => IllustrativeStickerFeature.SetPosition(GetManager(), StatusStickerPosition.BottomLeft));
                AddStatusStickerMenuItem(position, "Bottom right",
                    () => IllustrativeStickerFeature.SetPosition(GetManager(), StatusStickerPosition.BottomRight));
                btn.ContextMenu.Items.Add(position);

                AddStatusStickerMenuItem(btn.ContextMenu, "Use presentation Accent 1 color",
                    () => IllustrativeStickerFeature.UseThemeColor(GetManager()));
                AddStatusStickerMenuItem(btn.ContextMenu, "Toggle current date",
                    () => IllustrativeStickerFeature.ToggleDate(GetManager()));
                AddStatusStickerMenuItem(btn.ContextMenu, "Toggle author name",
                    () => IllustrativeStickerFeature.ToggleAuthor(GetManager()));

                btn.ContextMenu.Items.Add(new Separator());
                AddStatusStickerMenuItem(btn.ContextMenu, "Manage presets and defaults...", OpenStatusStickerSettings);
            }
            else if (btn.Name == "BtnHarveyBall")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var levels = new[] { 0f, 0.25f, 0.5f, 0.75f, 1f };
                foreach (var l in levels) {
                    var mi = new MenuItem { Header = $"Set {l * 100}%" };
                    mi.Click += (s, e) => HarveyBallFeature.Execute(GetManager(), l);
                    btn.ContextMenu.Items.Add(mi);
                }
            }
            else if (btn.Name == "BtnCheckbox")
            {
                if (btn.ContextMenu.Items.Count > 0)
                    btn.ContextMenu.Items.Add(new Separator());

                var labels = new[]
                {
                    GetCatalogVariantLabel(btn.Name, 0, "Framed Tick"),
                    GetCatalogVariantLabel(btn.Name, 1, "Framed Cross"),
                    GetCatalogVariantLabel(btn.Name, 2, "Framed Empty"),
                    GetCatalogVariantLabel(btn.Name, 3, "Unframed Tick"),
                    GetCatalogVariantLabel(btn.Name, 4, "Unframed Cross"),
                    GetCatalogVariantLabel(btn.Name, 5, "Unframed Empty")
                };
                AddCheckboxPresetMenu(btn, "Framed", labels, 0, true);
                AddCheckboxPresetMenu(btn, "Without frame", labels, 3, false);
            }
            else if (btn.Name == "BtnStarRating")
            {
                if (btn.ContextMenu.Items.Count > 0)
                    btn.ContextMenu.Items.Add(new Separator());

                var rating1 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Set / Insert 1 of 5") };
                rating1.Click += (s, e) => StarRatingFeature.Execute(GetManager(), 5, 1d, 0, false, 0x07C1FF);
                btn.ContextMenu.Items.Add(rating1);
                var rating2 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Set / Insert 2 of 5") };
                rating2.Click += (s, e) => StarRatingFeature.Execute(GetManager(), 5, 2d, 0, false, 0x07C1FF);
                btn.ContextMenu.Items.Add(rating2);
                var rating3 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Set / Insert 3 of 5") };
                rating3.Click += (s, e) => StarRatingFeature.Execute(GetManager(), 5, 3d, 0, false, 0x07C1FF);
                btn.ContextMenu.Items.Add(rating3);
                var rating4 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Set / Insert 4 of 5") };
                rating4.Click += (s, e) => StarRatingFeature.Execute(GetManager(), 5, 4d, 0, false, 0x07C1FF);
                btn.ContextMenu.Items.Add(rating4);
                var rating5 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 4, "Set / Insert 5 of 5") };
                rating5.Click += (s, e) => StarRatingFeature.Execute(GetManager(), 5, 5d, 0, false, 0x07C1FF);
                btn.ContextMenu.Items.Add(rating5);
            }
            else if (btn.Name == "BtnNumeration")
            {
                if (btn.ContextMenu.Items.Count > 0)
                    btn.ContextMenu.Items.Add(new Separator());

                var number1 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Insert Number 1") };
                number1.Click += (s, e) => InsertNumerationPreset(1);
                btn.ContextMenu.Items.Add(number1);
                var number2 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Insert Number 2") };
                number2.Click += (s, e) => InsertNumerationPreset(2);
                btn.ContextMenu.Items.Add(number2);
                var number3 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Insert Number 3") };
                number3.Click += (s, e) => InsertNumerationPreset(3);
                btn.ContextMenu.Items.Add(number3);
                var number4 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Insert Number 4") };
                number4.Click += (s, e) => InsertNumerationPreset(4);
                btn.ContextMenu.Items.Add(number4);
                var number5 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 4, "Insert Number 5") };
                number5.Click += (s, e) => InsertNumerationPreset(5);
                btn.ContextMenu.Items.Add(number5);
            }
            else if (btn.Name == "BtnProgressSeries")
            {
                if (btn.ContextMenu.Items.Count > 0)
                    btn.ContextMenu.Items.Add(new Separator());

                var series3 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "3-Step Number Series") };
                series3.Click += (s, e) => InsertProgressSeriesPreset(3, 2, "Numbers", "H");
                btn.ContextMenu.Items.Add(series3);
                var series5 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "5-Step Number Series") };
                series5.Click += (s, e) => InsertProgressSeriesPreset(5, 3, "Numbers", "H");
                btn.ContextMenu.Items.Add(series5);
                var vertical5 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "5-Step Vertical Series") };
                vertical5.Click += (s, e) => InsertProgressSeriesPreset(5, 3, "Numbers", "V");
                btn.ContextMenu.Items.Add(vertical5);
                var bullets5 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "5-Step Bullet Series") };
                bullets5.Click += (s, e) => InsertProgressSeriesPreset(5, 3, "Bullets", "H");
                btn.ContextMenu.Items.Add(bullets5);
            }
            else if (btn.Name == "BtnThermometer")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var levels = new[] { 0f, 0.25f, 0.5f, 0.75f, 1f };
                foreach (var l in levels) {
                    var mi = new MenuItem { Header = $"Set {l * 100}%" };
                    mi.Click += (s, e) => ThermometerFeature.Execute(GetManager(), l);
                    btn.ContextMenu.Items.Add(mi);
                }
            }
            else if (btn.Name == "BtnThemeColor")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var fill = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Open Fill Theme Colors") };
                fill.Click += (s, e) => ThemeColorPickerFeature.OpenForTarget(GetManager(), PickColorFeature.ApplyTarget.Fill);
                btn.ContextMenu.Items.Add(fill);
                var line = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Open Line Theme Colors") };
                line.Click += (s, e) => ThemeColorPickerFeature.OpenForTarget(GetManager(), PickColorFeature.ApplyTarget.Line);
                btn.ContextMenu.Items.Add(line);
                var text = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Open Text Theme Colors") };
                text.Click += (s, e) => ThemeColorPickerFeature.OpenForTarget(GetManager(), PickColorFeature.ApplyTarget.Text);
                btn.ContextMenu.Items.Add(text);
                btn.ContextMenu.Items.Add(new Separator());
                var docked = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Show / Hide Docked Palette") };
                docked.Click += (s, e) => _host?.ToggleThemeColorPane();
                btn.ContextMenu.Items.Add(docked);
            }
            else if (btn.Name == "BtnSpecialChars")
            {
                btn.ContextMenu.Items.Clear();

                var symbolsItem = new MenuItem { Header = "Symbols" };
                foreach (var category in Models.SpecialCharacterCatalog.Symbols)
                {
                    var subItem = new MenuItem { Header = category.Header };
                    subItem.Click += (s, e) => OpenSpecialCharactersWindow(category.Header);
                    symbolsItem.Items.Add(subItem);
                }
                btn.ContextMenu.Items.Add(symbolsItem);

                var languagesItem = new MenuItem { Header = "Languages" };
                foreach (var category in Models.SpecialCharacterCatalog.Languages)
                {
                    var subItem = new MenuItem { Header = category.Header };
                    subItem.Click += (s, e) => OpenSpecialCharactersWindow(category.Header);
                    languagesItem.Items.Add(subItem);
                }
                btn.ContextMenu.Items.Add(languagesItem);
            }
            else if (btn.Name == "BtnSlideNotes")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var mi1 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Delete Notes on This Slide") };
                mi1.Click += (s, e) => oPenEfficiency.Features.Utilities.SlideNotesFeature.DeleteNotesOnCurrentSlide(GetManager());
                btn.ContextMenu.Items.Add(mi1);

                var mi2 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Delete All Notes in Presentation") };
                mi2.Click += (s, e) => oPenEfficiency.Features.Utilities.SlideNotesFeature.DeleteAllNotes(GetManager());
                btn.ContextMenu.Items.Add(mi2);

                var mi3 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Show or Hide Review Note on Current Slide") };
                mi3.Click += (s, e) => oPenEfficiency.Features.Utilities.SlideNotesFeature.ToggleOverlay(GetManager());
                btn.ContextMenu.Items.Add(mi3);

                var mi4 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Open Review Notes Overview") };
                mi4.Click += (s, e) => oPenEfficiency.Features.Utilities.SlideNotesFeature.OpenOverview(GetManager());
                btn.ContextMenu.Items.Add(mi4);
            }
            else if (btn.Name == "BtnManageStickyNotes" || btn.Name == "BtnStickyNoteManager")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var mi1 = new MenuItem { Header = "Move notes off this slide" };
                mi1.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.MoveOff, StickyNoteManagerFeature.Scope.ThisSlide);
                btn.ContextMenu.Items.Add(mi1);

                var mi2 = new MenuItem { Header = "Move notes back on this slide" };
                mi2.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.MoveOn, StickyNoteManagerFeature.Scope.ThisSlide);
                btn.ContextMenu.Items.Add(mi2);

                var mi3 = new MenuItem { Header = "Delete notes on this slide" };
                mi3.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.Delete, StickyNoteManagerFeature.Scope.ThisSlide);
                btn.ContextMenu.Items.Add(mi3);

                var mi4 = new MenuItem { Header = "Copy comments on this slide to notes" };
                mi4.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.ConvertComments, StickyNoteManagerFeature.Scope.ThisSlide);
                btn.ContextMenu.Items.Add(mi4);

                btn.ContextMenu.Items.Add(new Separator());

                var mi5 = new MenuItem { Header = "Move notes off selected slides" };
                mi5.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.MoveOff, StickyNoteManagerFeature.Scope.SelectedSlides);
                btn.ContextMenu.Items.Add(mi5);

                var mi6 = new MenuItem { Header = "Move notes back on selected slides" };
                mi6.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.MoveOn, StickyNoteManagerFeature.Scope.SelectedSlides);
                btn.ContextMenu.Items.Add(mi6);

                var mi7 = new MenuItem { Header = "Delete notes on selected slides" };
                mi7.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.Delete, StickyNoteManagerFeature.Scope.SelectedSlides);
                btn.ContextMenu.Items.Add(mi7);

                var mi8 = new MenuItem { Header = "Copy comments on selected slides to notes" };
                mi8.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.ConvertComments, StickyNoteManagerFeature.Scope.SelectedSlides);
                btn.ContextMenu.Items.Add(mi8);

                btn.ContextMenu.Items.Add(new Separator());

                var mi9 = new MenuItem { Header = "Mark selected shapes as sticky notes" };
                mi9.Click += (s, e) => StickyNoteManagerFeature.Execute(GetManager(), StickyNoteManagerFeature.ActionType.PrefixName, StickyNoteManagerFeature.Scope.ThisSlide);
                btn.ContextMenu.Items.Add(mi9);
            }
        }
    }
}
