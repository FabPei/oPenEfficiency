using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;
using System.Linq;
using System.Runtime.InteropServices;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Service for generating agenda slides based on saved layout configurations.
    /// </summary>
    public class AgendaGenerationService
    {
        private const string ThinkCellPresentationProtectionTag = "THINKCELLPRESENTATIONDONOTDELETE";
        private const string ThinkCellShapeProtectionTag = "THINKCELLSHAPEDONOTDELETE";
        private readonly PowerPointManager _manager;
        private readonly AgendaLayoutConfig _layout;
        private readonly int _targetSlideId; // Store ID for stability

        public AgendaGenerationService(PowerPointManager manager, Slide targetSlide, AgendaLayoutConfig layout)
        {
            _manager = manager;
            _layout = layout;
            _targetSlideId = targetSlide.SlideID;
        }

        public static void EnsureMutationAllowed(
            Presentation presentation,
            bool synchronizeExisting,
            bool createSections)
        {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));

            try
            {
                bool presentationProtected =
                    !string.IsNullOrEmpty(presentation.Tags[ThinkCellPresentationProtectionTag]);
                string protectionError = AgendaWorkflowPolicy.GetMutationProtectionError(
                    presentationProtected,
                    false,
                    synchronizeExisting,
                    createSections);
                if (protectionError != null) throw new InvalidOperationException(protectionError);

                if (!synchronizeExisting && !createSections) return;

                bool protectedSlideContentFound = false;
                for (int slideIndex = 1; slideIndex <= presentation.Slides.Count; slideIndex++)
                {
                    Slide slide = presentation.Slides[slideIndex];
                    if (IsThinkCellProtected(slide))
                    {
                        protectedSlideContentFound = true;
                        break;
                    }
                }

                protectionError = AgendaWorkflowPolicy.GetMutationProtectionError(
                    false,
                    protectedSlideContentFound,
                    synchronizeExisting,
                    createSections);
                if (protectionError != null) throw new InvalidOperationException(protectionError);
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(
                    "Agenda Creator could not verify think-cell protection before changing the presentation.",
                    ex);
            }
            catch (InvalidComObjectException ex)
            {
                throw new InvalidOperationException(
                    "Agenda Creator could not verify think-cell protection because a PowerPoint object is no longer available.",
                    ex);
            }
        }

        private static bool IsThinkCellProtected(Slide slide)
        {
            if (slide == null) return false;
            if (!string.IsNullOrEmpty(slide.Tags[ThinkCellShapeProtectionTag])) return true;

            for (int shapeIndex = 1; shapeIndex <= slide.Shapes.Count; shapeIndex++)
            {
                if (!string.IsNullOrEmpty(slide.Shapes[shapeIndex].Tags[ThinkCellShapeProtectionTag]))
                    return true;
            }

            return false;
        }

        public void GenerateFullAgenda(
            List<AgendaItemData> items,
            bool createSections,
            bool createDividerSlides,
            bool highlightCurrent = false,
            bool synchronizeExisting = false,
            bool createStarterSlides = false,
            string agendaTitle = null,
            string agendaSubtitle = null)
        {
            if (items == null || items.Count == 0) return;

            var app = _manager.GetApplication();
            var pres = app.ActivePresentation;
            Slide targetSlide = GetSafeSlide(pres, _targetSlideId);
            if (targetSlide == null) return;

            Dictionary<string, List<int>> existingContent = null;
            if (synchronizeExisting)
            {
                existingContent = GetContentSlideMap(pres);
                var oldTopics = GetExistingAgendaTopics(pres);
                ClearExistingAgendaSections(pres, oldTopics);
                ClearExistingDividerSlides(pres, _targetSlideId);
            }

            // Store active layout info
            try
            {
                pres.Tags.Add("oPE_AgendaLayout", _layout.LayoutName);
                targetSlide.Tags.Add("oPE_AgendaLayout", _layout.LayoutName);
                targetSlide.Tags.Add("oPE_AgendaActive", "True");
            }
            catch (System.Exception ex)
            {
                throw new System.InvalidOperationException("Agenda metadata could not be stored.", ex);
            }

            // 0. Cleanup
            ClearExistingAgendaShapes(targetSlide);
            ApplySlideFrame(targetSlide, agendaTitle, agendaSubtitle);
            
            // Start inserting after the target slide
            int currentTargetIndex = GetSafeSlideIndex(targetSlide);
            int insertSlideIndex = currentTargetIndex + 1;
            
            // Map item index to its divider slide ID
            var dividerSlideIds = new Dictionary<int, int>();
            
            // 1. Create Divider Slides and Sections
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                if (AgendaWorkflowPolicy.ShouldCreateChapterInfrastructure(
                    item.Level,
                    item.IsBreak,
                    item.ShouldGenerate))
                {
                    Slide dividerSlide = null;
                    if (createDividerSlides)
                    {
                        try {
                            int actualInsertIndex = Math.Min(insertSlideIndex, pres.Slides.Count + 1);
                            dividerSlide = pres.Slides.Add(actualInsertIndex, targetSlide.Layout);
                            dividerSlide.Tags.Add("EE4P_AGENDA_DIVIDER", "True");
                            dividerSlide.Tags.Add("oPE_AgendaSlide", "True");
                            dividerSlide.Tags.Add("oPE_AgendaLayout", _layout.LayoutName);
                            dividerSlide.Tags.Add("oPE_AgendaTopic", item.Topic);
                            dividerSlide.Tags.Add("oPE_AgendaItemId", item.Id);
                            
                            // Setup title: Always "Agenda"
                            SetupSlideTitle(dividerSlide, "Agenda");

                            // Clean other placeholders
                            SafeCleanPlaceholders(dividerSlide);

                            dividerSlideIds[i] = dividerSlide.SlideID;
                            item.PageNumber = dividerSlide.SlideIndex.ToString();
                        }
                        catch (System.Exception ex)
                        {
                            throw new System.InvalidOperationException($"Divider slide {i + 1} could not be created.", ex);
                        }
                    }
                    
                    if (createSections)
                    {
                        try {
                            // Since we cleared sections above, we just create them fresh at the right position.
                            int targetIndexForSection = dividerSlide != null ? GetSafeSlideIndex(dividerSlide) : insertSlideIndex;
                            pres.SectionProperties.AddBeforeSlide(Math.Min(targetIndexForSection, pres.Slides.Count + 1), item.Topic);
                        }
                        catch (System.Exception ex)
                        {
                            throw new System.InvalidOperationException($"Section '{item.Topic}' could not be created.", ex);
                        }
                    }
                    
                    if (createDividerSlides) insertSlideIndex++;
                }
            }

            var newDividerIds = dividerSlideIds
                .OrderBy(pair => pair.Key)
                .Select(pair => pair.Value)
                .ToList();

            if (synchronizeExisting && existingContent != null)
                RestoreContentSlidePositions(pres, newDividerIds, existingContent, items);

            if (createStarterSlides && createDividerSlides)
                CreateMissingStarterSlides(pres, dividerSlideIds, existingContent, items);

            // 2. Generate on main target slide
            var dividerSlides = new Dictionary<int, Slide>();
            foreach(var kvp in dividerSlideIds) {
                try { dividerSlides[kvp.Key] = pres.Slides.FindBySlideID(kvp.Value); }
                catch (Exception ex) { ReportOptionalAgendaFailure(ex, "AgendaGenerationService.Generate.DividerLookup"); }
            }
            GenerateAgendaItems(items, -1, dividerSlides, highlightCurrent);

            // 3. Generate on each divider slide
            if (createDividerSlides)
            {
                foreach (var kvp in dividerSlides)
                {
                    int itemIndex = kvp.Key;
                    Slide slide = kvp.Value;
                    if (slide == null) continue;

                    var dividerGenerationService = new AgendaGenerationService(_manager, slide, _layout);
                    dividerGenerationService.GenerateAgendaItems(items, itemIndex, dividerSlides, highlightCurrent);
                }
            }
        }

        #region Neighborhood Logic

        private List<string> GetExistingAgendaTopics(Presentation pres)
        {
            var topics = new List<string>();
            if (pres == null) return topics;

            for (int i = 1; i <= pres.Slides.Count; i++)
            {
                try
                {
                    string topic = pres.Slides[i].Tags["oPE_AgendaTopic"];
                    if (!string.IsNullOrWhiteSpace(topic))
                        topics.Add(topic);
                }
                catch (Exception ex) { ReportOptionalAgendaFailure(ex, "AgendaGenerationService.GetExistingTopics.ReadTag"); }
            }

            return topics;
        }

        /// <summary>
        /// Deletes all PowerPoint sections that contain at least one oPE agenda slide,
        /// or whose name matches an old agenda topic.
        /// </summary>
        private void ClearExistingAgendaSections(Presentation pres, List<string> oldTopics)
        {
            if (pres == null) return;
            try
            {
                dynamic dynSections = pres.SectionProperties;
                // Iterate backwards through sections to safe deletion
                for (int i = pres.SectionProperties.Count; i >= 1; i--)
                {
                    string sectionName = pres.SectionProperties.Name(i);
                    
                    // Case 1: Name matches an old topic
                    bool isAgendaSection = oldTopics.Contains(sectionName, StringComparer.OrdinalIgnoreCase);
                    
                    // Case 2: Contains slides with agenda tags
                    if (!isAgendaSection)
                    {
                        int firstSlide = dynSections.FirstSlideIndex(i);
                        int slideCount = dynSections.SlidesCount(i);
                        
                        for (int s = 0; s < slideCount; s++)
                        {
                            try {
                                var slide = pres.Slides[firstSlide + s];
                                if (!string.IsNullOrEmpty(slide.Tags["oPE_AgendaSlide"]) || 
                                    !string.IsNullOrEmpty(slide.Tags["EE4P_AGENDA_DIVIDER"]) ||
                                    !string.IsNullOrEmpty(slide.Tags["oPE_AgendaTopic"]))
                                {
                                    isAgendaSection = true;
                                    break;
                                }
                            }
                            catch (Exception ex) { ReportOptionalAgendaFailure(ex, "AgendaGenerationService.RemoveAgendaSections.ReadSlideTag"); }
                        }
                    }

                    if (isAgendaSection)
                    {
                        // Delete the section (moves slides to the previous section)
                        pres.SectionProperties.Delete(i, false);
                    }
                }
            }
            catch (Exception ex) { ReportOptionalAgendaFailure(ex, "AgendaGenerationService.RemoveAgendaSections"); }
        }

        /// <summary>
        /// Scans the presentation and builds a map of: Topic -> List of content SlideIDs following that topic's divider.
        /// </summary>
        private Dictionary<string, List<int>> GetContentSlideMap(Presentation pres)
        {
            var map = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
            string currentAgendaKey = null;

            for (int i = 1; i <= pres.Slides.Count; i++)
            {
                var slide = pres.Slides[i];
                string topicTag = slide.Tags["oPE_AgendaTopic"];
                string itemIdTag = slide.Tags["oPE_AgendaItemId"];
                bool isDivider = !string.IsNullOrEmpty(slide.Tags["EE4P_AGENDA_DIVIDER"]) || !string.IsNullOrEmpty(slide.Tags["oPE_AgendaSlide"]);

                if (isDivider && !string.IsNullOrEmpty(topicTag))
                {
                    currentAgendaKey = GetAgendaKey(itemIdTag, topicTag);
                    if (!map.ContainsKey(currentAgendaKey)) map[currentAgendaKey] = new List<int>();
                }
                else if (!string.IsNullOrEmpty(slide.Tags["oPE_BackupSlide"]) ||
                         !string.IsNullOrEmpty(slide.Tags["EE4P_BACKUP_SLIDE"]))
                {
                    currentAgendaKey = null;
                }
                else if (currentAgendaKey != null && !isDivider)
                {
                    map[currentAgendaKey].Add(slide.SlideID);
                }
            }
            return map;
        }

        /// <summary>
        /// Moves content slides back to their positions following the new dividers.
        /// </summary>
        private void RestoreContentSlidePositions(Presentation pres, List<int> newDividerIds, Dictionary<string, List<int>> contentMap, List<AgendaItemData> items)
        {
            foreach (var agendaKey in contentMap.Keys)
            {
                Slide newDivider = null;
                foreach (int id in newDividerIds)
                {
                    try {
                        var s = pres.Slides.FindBySlideID(id);
                        string dividerKey = GetAgendaKey(
                            s.Tags["oPE_AgendaItemId"],
                            s.Tags["oPE_AgendaTopic"]);
                        bool legacyTopicMatch = agendaKey.StartsWith("topic:", StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(
                                "topic:" + (s.Tags["oPE_AgendaTopic"] ?? string.Empty).Trim(),
                                agendaKey,
                                StringComparison.OrdinalIgnoreCase);
                        if (string.Equals(dividerKey, agendaKey, StringComparison.OrdinalIgnoreCase) ||
                            legacyTopicMatch)
                        {
                            newDivider = s;
                            break;
                        }
                    }
                    catch (Exception ex) { ReportOptionalAgendaFailure(ex, "AgendaGenerationService.ReorderContent.FindDivider"); }
                }

                if (newDivider != null)
                {
                    var slideIdsToMove = contentMap[agendaKey];
                    int targetPos = newDivider.SlideIndex;

                    foreach (int slideId in slideIdsToMove)
                    {
                        try {
                            var slideToMove = pres.Slides.FindBySlideID(slideId);
                            if (slideToMove != null)
                            {
                                targetPos++;
                                slideToMove.MoveTo(targetPos);
                            }
                        }
                        catch (Exception ex) { ReportOptionalAgendaFailure(ex, "AgendaGenerationService.ReorderContent.MoveSlide"); }
                    }
                }
            }
        }

        private static string GetAgendaKey(string itemId, string topic)
        {
            return !string.IsNullOrWhiteSpace(itemId)
                ? "id:" + itemId.Trim()
                : "topic:" + (topic ?? string.Empty).Trim();
        }

        private void CreateMissingStarterSlides(
            Presentation pres,
            Dictionary<int, int> dividerSlideIds,
            Dictionary<string, List<int>> existingContent,
            List<AgendaItemData> items)
        {
            foreach (var pair in dividerSlideIds.OrderByDescending(entry => entry.Key))
            {
                var item = items[pair.Key];
                if (item.Level != 0 || item.IsBreak) continue;

                bool hasExistingContent = existingContent != null &&
                    ((existingContent.TryGetValue(GetAgendaKey(item.Id, item.Topic), out var slideIds) &&
                      slideIds.Count > 0) ||
                     (existingContent.TryGetValue(GetAgendaKey(null, item.Topic), out slideIds) &&
                      slideIds.Count > 0));
                if (hasExistingContent) continue;

                try
                {
                    var divider = pres.Slides.FindBySlideID(pair.Value);
                    var starter = pres.Slides.Add(
                        Math.Min(divider.SlideIndex + 1, pres.Slides.Count + 1),
                        PpSlideLayout.ppLayoutText);
                    starter.Tags.Add("oPE_AgendaStarterSlide", "True");
                    starter.Tags.Add("oPE_AgendaItemId", item.Id);
                    starter.Tags.Add("oPE_AgendaTopic", item.Topic);
                    if (starter.Shapes.HasTitle == Office.MsoTriState.msoTrue)
                        starter.Shapes.Title.TextFrame.TextRange.Text = item.Topic;
                }
                catch (System.Exception ex)
                {
                    throw new System.InvalidOperationException(
                        $"A starter slide for '{item.Topic}' could not be created.",
                        ex);
                }
            }
        }

        #endregion

        private Slide GetSafeSlide(Presentation pres, int id)
        {
            try { return pres.Slides.FindBySlideID(id); } catch { return null; }
        }

        private int GetSafeSlideIndex(Slide slide)
        {
            try { return slide.SlideIndex; }
            catch {
                try {
                    int id = slide.SlideID;
                    var pres = (Presentation)slide.Parent;
                    return pres.Slides.FindBySlideID(id).SlideIndex;
                }
                catch { return 1; }
            }
        }

        private void SetupSlideTitle(Slide slide, string title)
        {
            try {
                if (_layout.HeadlineShape != null)
                {
                    Shape titleShape = null;
                    if (slide.Shapes.HasTitle == Office.MsoTriState.msoTrue) {
                        titleShape = slide.Shapes.Title;
                    } else {
                        titleShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 
                            _layout.HeadlineShape.Left, _layout.HeadlineShape.Top, 
                            _layout.HeadlineShape.Width, _layout.HeadlineShape.Height);
                    }

                    titleShape.TextFrame.TextRange.Text = title;
                    titleShape.Left = _layout.HeadlineShape.Left;
                    titleShape.Top = _layout.HeadlineShape.Top;
                    titleShape.Width = _layout.HeadlineShape.Width;
                    titleShape.Height = _layout.HeadlineShape.Height;
                    ApplyShapeFormat(titleShape, _layout.HeadlineShape);
                    titleShape.Tags.Add("oPE_AgendaRole", "Headline");
                }
                else if (slide.Shapes.HasTitle == Office.MsoTriState.msoTrue) {
                    slide.Shapes.Title.TextFrame.TextRange.Text = title;
                }
            }
            catch (Exception ex) { ReportOptionalAgendaFailure(ex, "AgendaGenerationService.SetSlideTitle"); }
        }

        private void ApplySlideFrame(Slide slide, string title, string subtitle)
        {
            if (_layout.SlideShape != null)
            {
                if (_layout.SlideShape.HasFill)
                {
                    slide.FollowMasterBackground = Office.MsoTriState.msoFalse;
                    slide.Background.Fill.Solid();
                    slide.Background.Fill.ForeColor.RGB = _layout.SlideShape.FillColorRgb;
                }
                else
                {
                    slide.FollowMasterBackground = Office.MsoTriState.msoTrue;
                }
            }

            SetupSlideTitle(slide, string.IsNullOrWhiteSpace(title) ? "Agenda" : title.Trim());
            if (_layout.SubHeadlineShape != null && !string.IsNullOrWhiteSpace(subtitle))
                CreateShape(_layout.SubHeadlineShape, subtitle.Trim(), "SubHeadline");
        }

        private void SafeCleanPlaceholders(Slide slide)
        {
            try {
                for (int sIdx = slide.Shapes.Count; sIdx >= 1; sIdx--)
                {
                    try { 
                        var s = slide.Shapes[sIdx];
                        if (s.Type == Office.MsoShapeType.msoPlaceholder)
                        {
                            if (s.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderTitle || 
                                s.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderCenterTitle)
                                continue;
                        }
                        if (s.Tags["oPE_AgendaRole"] == "Headline") continue;
                        s.Delete(); 
                    }
                    catch (Exception ex) { ReportOptionalAgendaFailure(ex, "AgendaGenerationService.SafeCleanPlaceholders.DeleteShape"); }
                }
            }
            catch (Exception ex) { ReportOptionalAgendaFailure(ex, "AgendaGenerationService.SafeCleanPlaceholders"); }
        }

        private void ClearExistingAgendaShapes(Slide slide)
        {
            if (slide == null) return;
            try {
                for (int i = slide.Shapes.Count; i >= 1; i--)
                {
                    try {
                        var shape = slide.Shapes[i];
                        if (!string.IsNullOrEmpty(shape.Tags["EE4P_AGENDAWIZARD"]) || !string.IsNullOrEmpty(shape.Tags["oPE_AgendaRole"]))
                            shape.Delete();
                    }
                    catch (Exception ex) { ReportOptionalAgendaFailure(ex, "AgendaGenerationService.ClearExistingAgendaShapes.DeleteShape"); }
                }
            }
            catch (Exception ex) { ReportOptionalAgendaFailure(ex, "AgendaGenerationService.ClearExistingAgendaShapes"); }
        }

        private void ClearExistingDividerSlides(Presentation pres, int protectId)
        {
            if (pres == null) return;
            try {
                for (int i = pres.Slides.Count; i >= 1; i--)
                {
                    try {
                        var slide = pres.Slides[i];
                        if (slide.SlideID == protectId) continue;
                        if (!string.IsNullOrEmpty(slide.Tags["EE4P_AGENDA_DIVIDER"]))
                            slide.Delete();
                    }
                    catch (Exception ex) { ReportOptionalAgendaFailure(ex, "AgendaGenerationService.ClearExistingDividerSlides.DeleteSlide"); }
                }
            }
            catch (Exception ex) { ReportOptionalAgendaFailure(ex, "AgendaGenerationService.ClearExistingDividerSlides"); }
        }

        public void GenerateAgendaItems(List<AgendaItemData> items, int activeItemIndex = -1, Dictionary<int, Slide> dividerSlides = null, bool highlightCurrent = false)
        {
            if (items == null || items.Count == 0) return;
            if (_layout.ItemTextShape == null) throw new InvalidOperationException("ItemTextShape is not defined.");

            float itemSpacing = _layout.ItemSpacing;
            float anchorX = _layout.ItemTextShape.Left;
            float anchorY = _layout.ItemTextShape.Top;
            int[] levelCounters = new int[20];

            bool isHorizontal = _layout.ListOrientation == "Horizontal";

            if (isHorizontal)
                GenerateHorizontalAgendaItems(items, levelCounters, anchorX, anchorY, itemSpacing, activeItemIndex, dividerSlides, highlightCurrent);
            else
                GenerateVerticalAgendaItems(items, levelCounters, anchorY, itemSpacing, activeItemIndex, dividerSlides, highlightCurrent);
        }

        private void GenerateVerticalAgendaItems(List<AgendaItemData> items, int[] levelCounters, float anchorY, float itemSpacing, int activeItemIndex = -1, Dictionary<int, Slide> dividerSlides = null, bool highlightCurrent = false)
        {
            float currentY = anchorY;
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                levelCounters[item.Level]++;
                for (int l = item.Level + 1; l < levelCounters.Length; l++) levelCounters[l] = 0;
                string displayNumber = (item.Level == 0) ? levelCounters[0].ToString() : string.Join(".", levelCounters.Take(item.Level + 1).Where(c => c > 0));

                if (item.Level == 0)
                    currentY = GenerateMainItem(item, displayNumber, currentY, null, activeItemIndex, i, dividerSlides, highlightCurrent);
                else
                    currentY = GenerateSubItem(item, displayNumber, currentY, null, activeItemIndex, i, highlightCurrent);

                currentY += itemSpacing;
            }
        }

        private void GenerateHorizontalAgendaItems(List<AgendaItemData> items, int[] levelCounters, float anchorX, float anchorY, float itemSpacing, int activeItemIndex = -1, Dictionary<int, Slide> dividerSlides = null, bool highlightCurrent = false)
        {
            var app = _manager.GetApplication();
            var pres = app.ActivePresentation;
            float rightMargin = pres.PageSetup.SlideWidth - 20;
            float currentRowX = anchorX;
            float currentRowY = anchorY;
            float maxRowHeight = 0;

            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                levelCounters[item.Level]++;
                for (int l = item.Level + 1; l < levelCounters.Length; l++) levelCounters[l] = 0;
                string displayNumber = (item.Level == 0) ? levelCounters[0].ToString() : string.Join(".", levelCounters.Take(item.Level + 1).Where(c => c > 0));

                if (item.Level == 0)
                {
                    float itemWidth = GetItemTotalWidth(item);
                    if (i > 0 && currentRowX + itemWidth > rightMargin)
                    {
                        currentRowX = anchorX;
                        currentRowY += maxRowHeight + itemSpacing;
                        maxRowHeight = 0;
                    }

                    float itemBottomY = GenerateMainItem(item, displayNumber, currentRowY, currentRowX, activeItemIndex, i, dividerSlides, highlightCurrent);
                    maxRowHeight = Math.Max(maxRowHeight, itemBottomY - currentRowY);
                    currentRowX += itemWidth + itemSpacing;
                }
                else
                {
                    GenerateSubItem(item, displayNumber, currentRowY + maxRowHeight, null, activeItemIndex, i, highlightCurrent);
                }
            }
        }

        private float GetItemTotalWidth(AgendaItemData item)
        {
            float totalWidth = _layout.ItemTextShape.Width;
            if (_layout.ItemNumberingShape != null) totalWidth += _layout.ItemNumberingShape.Width + 10;
            if (_layout.ItemTimeShape != null && !string.IsNullOrEmpty(item.TimeSlot)) totalWidth += _layout.ItemTimeShape.Width + 10;
            return totalWidth;
        }

        private float GenerateMainItem(AgendaItemData item, string displayNumber, float startY, float? itemX, int activeIndex, int currentIndex, Dictionary<int, Slide> dividerSlides, bool highlightCurrent)
        {
            Shape textShape = null;
            float bottomY = startY;
            bool isActive = highlightCurrent && (activeIndex == currentIndex);
            Slide linkTarget = (dividerSlides != null && dividerSlides.ContainsKey(currentIndex)) ? dividerSlides[currentIndex] : null;

            string topicText = (_layout.ItemNumberingShape == null) ? $"{displayNumber}. {item.Topic}" : item.Topic;

            if (_layout.ItemTextShape != null)
            {
                textShape = CreateShape(_layout.ItemTextShape, topicText, "Topic", isActive, linkTarget, item.Id);
                textShape.Top = startY;
                if (itemX.HasValue) textShape.Left = itemX.Value;
                bottomY = Math.Max(bottomY, textShape.Top + textShape.Height);
            }

            if (_layout.ItemNumberingShape != null)
            {
                var numShape = CreateShape(_layout.ItemNumberingShape, displayNumber, "ItemNo", isActive, linkTarget, item.Id);
                numShape.Top = startY + (_layout.AlignInfoShapesWithItem ? 0 : (_layout.ItemNumberingShape.Top - _layout.ItemTextShape.Top));
                if (textShape != null) numShape.Left = textShape.Left + (_layout.ItemNumberingShape.Left - _layout.ItemTextShape.Left);
                bottomY = Math.Max(bottomY, numShape.Top + numShape.Height);
            }

            if (_layout.ItemTimeShape != null && !string.IsNullOrEmpty(item.TimeSlot))
            {
                var timeShape = CreateShape(_layout.ItemTimeShape, item.TimeSlot, "TimeSlot", isActive, null, item.Id);
                PositionInfoShape(timeShape, textShape, "Time");
                bottomY = Math.Max(bottomY, timeShape.Top + timeShape.Height);
            }

            if (_layout.ItemPageNumberShape != null && !string.IsNullOrEmpty(item.PageNumber))
            {
                var pageShape = CreateShape(_layout.ItemPageNumberShape, item.PageNumber, "PageNumber", isActive, linkTarget, item.Id);
                PositionInfoShape(pageShape, textShape, "PageNumber");
                bottomY = Math.Max(bottomY, pageShape.Top + pageShape.Height);
            }

            if (_layout.ItemDurationShape != null && item.ShowMinutes)
            {
                var durationShape = CreateShape(
                    _layout.ItemDurationShape,
                    $"{item.Duration} min",
                    "Duration",
                    isActive,
                    null,
                    item.Id);
                PositionInfoShape(durationShape, textShape, "Duration");
                bottomY = Math.Max(bottomY, durationShape.Top + durationShape.Height);
            }

            if (_layout.ItemResponsibleShape != null &&
                item.ShowResponsible &&
                !string.IsNullOrWhiteSpace(item.Responsible))
            {
                var responsibleShape = CreateShape(
                    _layout.ItemResponsibleShape,
                    item.Responsible,
                    "Responsible",
                    isActive,
                    null,
                    item.Id);
                PositionInfoShape(responsibleShape, textShape, "Responsible");
                bottomY = Math.Max(bottomY, responsibleShape.Top + responsibleShape.Height);
            }

            if (_layout.ItemBackgroundShape != null)
            {
                 var bgShape = CreateShape(_layout.ItemBackgroundShape, "", "Background", isActive, null, item.Id);
                 bgShape.Top = startY;
                 if (itemX.HasValue) bgShape.Left = itemX.Value + (_layout.ItemBackgroundShape.Left - _layout.ItemTextShape.Left);
                 else bgShape.Left = textShape != null ? textShape.Left + (_layout.ItemBackgroundShape.Left - _layout.ItemTextShape.Left) : _layout.ItemBackgroundShape.Left;
                 bgShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
            }

            return bottomY;
        }

        private float GenerateSubItem(AgendaItemData item, string displayNumber, float startY, float? itemX, int activeIndex, int currentIndex, bool highlightCurrent)
        {
            Shape textShape = null;
            float bottomY = startY;
            bool isActive = highlightCurrent && (activeIndex == currentIndex);

            ShapeFormatConfig textConfig = _layout.SubItemTextShape ?? _layout.ItemTextShape;
            string topicText = (_layout.SubItemNumberingShape == null && _layout.ItemNumberingShape == null) ? $"{displayNumber}. {item.Topic}" : item.Topic;

            if (textConfig != null)
            {
                textShape = CreateShape(textConfig, topicText, "Topic", isActive, null, item.Id);
                textShape.Top = startY;
                if (itemX.HasValue) textShape.Left = itemX.Value;
                else textShape.Left = textConfig.Left;
                bottomY = Math.Max(bottomY, textShape.Top + textShape.Height);
            }

            if (textConfig != null)
            {
                ShapeFormatConfig numConfig = _layout.SubItemNumberingShape ?? _layout.ItemNumberingShape;
                if (numConfig != null)
                {
                    var numShape = CreateShape(numConfig, displayNumber, "ItemNo", isActive, null, item.Id);
                    numShape.Top = startY + (_layout.AlignInfoShapesWithItem ? 0 : (numConfig.Top - textConfig.Top));
                    if (textShape != null) numShape.Left = textShape.Left + (numConfig.Left - textConfig.Left);
                    bottomY = Math.Max(bottomY, numShape.Top + numShape.Height);
                }
            }

            if (_layout.SubItemBackgroundShape != null || _layout.ItemBackgroundShape != null)
            {
                 var bgConfig = _layout.SubItemBackgroundShape ?? _layout.ItemBackgroundShape;
                 var bgShape = CreateShape(bgConfig, "", "Background", isActive, null, item.Id);
                 bgShape.Top = startY;
                 bgShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
            }

            ShapeFormatConfig durationConfig = _layout.SubItemDurationShape ?? _layout.ItemDurationShape;
            if (durationConfig != null && item.ShowMinutes)
            {
                var durationShape = CreateShape(
                    durationConfig,
                    $"{item.Duration} min",
                    "Duration",
                    isActive,
                    null,
                    item.Id);
                PositionInfoShape(durationShape, textShape, "Duration", true);
                bottomY = Math.Max(bottomY, durationShape.Top + durationShape.Height);
            }

            ShapeFormatConfig responsibleConfig = _layout.SubItemResponsibleShape ?? _layout.ItemResponsibleShape;
            if (responsibleConfig != null &&
                item.ShowResponsible &&
                !string.IsNullOrWhiteSpace(item.Responsible))
            {
                var responsibleShape = CreateShape(
                    responsibleConfig,
                    item.Responsible,
                    "Responsible",
                    isActive,
                    null,
                    item.Id);
                PositionInfoShape(responsibleShape, textShape, "Responsible", true);
                bottomY = Math.Max(bottomY, responsibleShape.Top + responsibleShape.Height);
            }

            return bottomY;
        }

        private void PositionInfoShape(Shape infoShape, Shape textShape, string infoType, bool isSubItem = false)
        {
            if (textShape == null) return;
            ShapeFormatConfig textConfig = isSubItem ? (_layout.SubItemTextShape ?? _layout.ItemTextShape) : _layout.ItemTextShape;
            ShapeFormatConfig infoConfig = isSubItem ? (GetSubItemInfoConfig(infoType)) : GetItemInfoConfig(infoType);
            if (infoConfig == null) return;

            bool isHorizontal = _layout.ListOrientation == "Horizontal";
            bool align = _layout.AlignInfoShapesWithItem;

            if (isHorizontal) {
                infoShape.Left = align ? textShape.Left : textShape.Left + (infoConfig.Left - textConfig.Left);
                infoShape.Top = textShape.Top + (infoConfig.Top - textConfig.Top);
            } else {
                infoShape.Top = align ? textShape.Top : textShape.Top + (infoConfig.Top - textConfig.Top);
                infoShape.Left = textShape.Left + (infoConfig.Left - textConfig.Left);
            }
        }

        private ShapeFormatConfig GetItemInfoConfig(string type) {
            switch(type) {
                case "Time": return _layout.ItemTimeShape;
                case "Duration": return _layout.ItemDurationShape;
                case "Responsible": return _layout.ItemResponsibleShape;
                case "PageNumber": return _layout.ItemPageNumberShape;
                default: return null;
            }
        }

        private ShapeFormatConfig GetSubItemInfoConfig(string type) {
            switch(type) {
                case "Time": return _layout.SubItemTimeShape ?? _layout.ItemTimeShape;
                case "Duration": return _layout.SubItemDurationShape ?? _layout.ItemDurationShape;
                case "Responsible": return _layout.SubItemResponsibleShape ?? _layout.ItemResponsibleShape;
                case "PageNumber": return _layout.SubItemPageNumberShape ?? _layout.ItemPageNumberShape;
                default: return null;
            }
        }

        private void ApplyShapeFormat(Shape shape, ShapeFormatConfig config)
        {
            if (config.HasFill) {
                shape.Fill.Visible = Office.MsoTriState.msoTrue;
                shape.Fill.ForeColor.RGB = config.FillColorRgb;
            } else shape.Fill.Visible = Office.MsoTriState.msoFalse;

            if (config.HasLine) {
                shape.Line.Visible = Office.MsoTriState.msoTrue;
                shape.Line.ForeColor.RGB = config.LineColorRgb;
                shape.Line.Weight = config.LineWeight;
            } else shape.Line.Visible = Office.MsoTriState.msoFalse;

            if (shape.HasTextFrame == Office.MsoTriState.msoTrue) {
                var tr = shape.TextFrame.TextRange;
                tr.Font.Name = !string.IsNullOrEmpty(config.FontName) ? config.FontName : "Segoe UI";
                tr.Font.Size = config.FontSize;
                tr.Font.Color.RGB = config.FontColorRgb;
                if (config.FontBold) tr.Font.Bold = Office.MsoTriState.msoTrue;
                if (config.FontItalic) tr.Font.Italic = Office.MsoTriState.msoTrue;
                int alignment = config.TextAlignment;
                if (alignment < 1 || alignment > 4) alignment = 1;
                tr.ParagraphFormat.Alignment = (PpParagraphAlignment)alignment;
                shape.TextFrame.MarginLeft = config.MarginLeft;
                shape.TextFrame.MarginRight = config.MarginRight;
                shape.TextFrame.MarginTop = config.MarginTop;
                shape.TextFrame.MarginBottom = config.MarginBottom;
            }
        }

        private Shape CreateShape(
            ShapeFormatConfig config,
            string text,
            string role = null,
            bool isActive = false,
            Slide linkToSlide = null,
            string agendaItemId = null)
        {
            var app = _manager.GetApplication();
            var pres = app.ActivePresentation;
            Slide targetSlide = pres.Slides.FindBySlideID(_targetSlideId);
            Shape shape = targetSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, config.Left, config.Top, config.Width, config.Height);
            
            if (config.HasFill || (isActive && _layout.HighlightFillColor)) {
                shape.Fill.Visible = Office.MsoTriState.msoTrue;
                shape.Fill.ForeColor.RGB = (isActive && _layout.HighlightFillColor) ? _layout.HighlightFillColorRgb : config.FillColorRgb;
            } else shape.Fill.Visible = Office.MsoTriState.msoFalse;

            if (config.HasLine) {
                shape.Line.Visible = Office.MsoTriState.msoTrue;
                shape.Line.ForeColor.RGB = config.LineColorRgb;
                shape.Line.Weight = config.LineWeight;
            } else shape.Line.Visible = Office.MsoTriState.msoFalse;

            if (!string.IsNullOrEmpty(text)) {
                shape.TextFrame.TextRange.Text = text;
                var tr = shape.TextFrame.TextRange;
                tr.Font.Name = !string.IsNullOrEmpty(config.FontName) ? config.FontName : "Segoe UI";
                tr.Font.Size = config.FontSize;
                tr.Font.Color.RGB = (isActive && _layout.HighlightFontColor) ? _layout.HighlightFontColorRgb : config.FontColorRgb;
                if (isActive) {
                    if (_layout.HighlightBold) tr.Font.Bold = Office.MsoTriState.msoTrue;
                    if (_layout.HighlightItalic) tr.Font.Italic = Office.MsoTriState.msoTrue;
                    if (_layout.HighlightUnderline) tr.Font.Underline = Office.MsoTriState.msoTrue;
                } else {
                    if (config.FontBold) tr.Font.Bold = Office.MsoTriState.msoTrue;
                    if (config.FontItalic) tr.Font.Italic = Office.MsoTriState.msoTrue;
                }
                if (linkToSlide != null) {
                    try {
                        shape.ActionSettings[PpMouseActivation.ppMouseClick].Action = PpActionType.ppActionHyperlink;
                        shape.ActionSettings[PpMouseActivation.ppMouseClick].Hyperlink.SubAddress = linkToSlide.SlideID + "," + linkToSlide.SlideIndex + ",";
                    }
                    catch (Exception ex) { ReportOptionalAgendaFailure(ex, "AgendaGenerationService.AddTextShape.Hyperlink"); }
                }
                int align = config.TextAlignment; if (align < 1 || align > 4) align = 1;
                tr.ParagraphFormat.Alignment = (PpParagraphAlignment)align;
            }

            shape.TextFrame.MarginLeft = config.MarginLeft;
            shape.TextFrame.MarginRight = config.MarginRight;
            shape.TextFrame.MarginTop = config.MarginTop;
            shape.TextFrame.MarginBottom = config.MarginBottom;
            int orient = config.Orientation; if (orient < 1 || orient > 6) orient = 1;
            shape.TextFrame.Orientation = (Office.MsoTextOrientation)orient;

            if (!string.IsNullOrEmpty(role)) {
                string guid = string.IsNullOrWhiteSpace(agendaItemId)
                    ? Guid.NewGuid().ToString("D")
                    : agendaItemId;
                shape.Name = $"ee4p_item_{guid}_shape";
                shape.Tags.Add("oPE_AgendaRole", role);
                shape.Tags.Add("EE4P_AGENDAWIZARD", role);
                shape.Tags.Add("oPE_AgendaItemId", guid);
            }
            return shape;
        }

        private static void ReportOptionalAgendaFailure(Exception exception, string context)
        {
            if (!ExceptionPolicy.IsExpectedComFailure(exception))
            {
                ExceptionLogger.Log(exception, context);
            }
            // Missing or stale slide/shape COM state must not abort generation of other agenda items.
        }
    }

    public class AgendaItemData
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("D");
        public int Level { get; set; }
        public string Topic { get; set; }
        public int Duration { get; set; }
        public string TimeSlot { get; set; }
        public string PageNumber { get; set; }
        public bool IsBreak { get; set; }
        public string Responsible { get; set; }
        public bool ShouldGenerate { get; set; } = true;
        public bool ShowMinutes { get; set; } = true;
        public bool ShowResponsible { get; set; } = true;
    }
}
