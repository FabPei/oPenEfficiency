using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Models;
using System.Linq;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Service for generating agenda slides based on saved layout configurations.
    /// </summary>
    public class AgendaGenerationService
    {
        private readonly PowerPointManager _manager;
        private readonly AgendaLayoutConfig _layout;
        private readonly int _targetSlideId; // Store ID for stability

        public AgendaGenerationService(PowerPointManager manager, Slide targetSlide, AgendaLayoutConfig layout)
        {
            _manager = manager;
            _layout = layout;
            _targetSlideId = targetSlide.SlideID;
        }

        public void GenerateFullAgenda(List<AgendaItemData> items, bool createSections, bool createDividerSlides, bool highlightCurrent = false)
        {
            if (items == null || items.Count == 0) return;

            var app = _manager.GetApplication();
            var pres = app.ActivePresentation;
            Slide targetSlide = GetSafeSlide(pres, _targetSlideId);
            if (targetSlide == null) return;

            // --- NEIGHBORHOOD LOGIC: Capture current content slide mapping ---
            var contentMap = GetContentSlideMap(pres);
            var oldTopics = contentMap.Keys.ToList();

            // Store active layout info
            try
            {
                pres.Tags.Add("oPE_AgendaLayout", _layout.LayoutName);
                targetSlide.Tags.Add("oPE_AgendaLayout", _layout.LayoutName);
                targetSlide.Tags.Add("oPE_AgendaActive", "True");
            }
            catch { }

            // 0. Cleanup
            ClearExistingAgendaShapes(targetSlide);
            
            // CRITICAL: Clear sections BEFORE deleting the divider slides, 
            // so we can still identify sections via slide tags.
            if (createSections)
            {
                ClearExistingAgendaSections(pres, oldTopics);
            }

            if (createDividerSlides)
            {
                ClearExistingDividerSlides(pres, _targetSlideId);
            }

            // Start inserting after the target slide
            int currentTargetIndex = GetSafeSlideIndex(targetSlide);
            int insertSlideIndex = currentTargetIndex + 1;
            
            // Map item index to its divider slide ID
            var dividerSlideIds = new Dictionary<int, int>();
            
            // 1. Create Divider Slides and Sections
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                if (item.ShouldGenerate)
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
                            dividerSlide.Tags.Add("oPE_AgendaTopic", item.Topic); // CRITICAL: Store topic name in tag
                            
                            // Setup title: Always "Agenda"
                            SetupSlideTitle(dividerSlide, "Agenda");

                            // Clean other placeholders
                            SafeCleanPlaceholders(dividerSlide);

                            dividerSlideIds[i] = dividerSlide.SlideID;
                            item.PageNumber = dividerSlide.SlideIndex.ToString();
                        } catch { }
                    }
                    
                    if (createSections)
                    {
                        try {
                            // Since we cleared sections above, we just create them fresh at the right position.
                            int targetIndexForSection = dividerSlide != null ? GetSafeSlideIndex(dividerSlide) : insertSlideIndex;
                            pres.SectionProperties.AddBeforeSlide(Math.Min(targetIndexForSection, pres.Slides.Count + 1), item.Topic);
                        } catch { }
                    }
                    
                    if (createDividerSlides) insertSlideIndex++;
                }
            }

            // --- NEIGHBORHOOD LOGIC: Restore content slide positions ---
            if (createDividerSlides)
            {
                RestoreContentSlidePositions(pres, dividerSlideIds.Values.ToList(), contentMap, items);
            }

            // 2. Generate on main target slide
            var dividerSlides = new Dictionary<int, Slide>();
            foreach(var kvp in dividerSlideIds) {
                try { dividerSlides[kvp.Key] = pres.Slides.FindBySlideID(kvp.Value); } catch {}
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
                            } catch { }
                        }
                    }

                    if (isAgendaSection)
                    {
                        // Delete the section (moves slides to the previous section)
                        pres.SectionProperties.Delete(i, false);
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Scans the presentation and builds a map of: Topic -> List of content SlideIDs following that topic's divider.
        /// </summary>
        private Dictionary<string, List<int>> GetContentSlideMap(Presentation pres)
        {
            var map = new Dictionary<string, List<int>>();
            string currentTopic = null;

            for (int i = 1; i <= pres.Slides.Count; i++)
            {
                var slide = pres.Slides[i];
                string topicTag = slide.Tags["oPE_AgendaTopic"];
                bool isDivider = !string.IsNullOrEmpty(slide.Tags["EE4P_AGENDA_DIVIDER"]) || !string.IsNullOrEmpty(slide.Tags["oPE_AgendaSlide"]);

                if (isDivider && !string.IsNullOrEmpty(topicTag))
                {
                    currentTopic = topicTag;
                    if (!map.ContainsKey(currentTopic)) map[currentTopic] = new List<int>();
                }
                else if (currentTopic != null && !isDivider)
                {
                    // This is a content slide belonging to currentTopic
                    map[currentTopic].Add(slide.SlideID);
                }
            }
            return map;
        }

        /// <summary>
        /// Moves content slides back to their positions following the new dividers.
        /// </summary>
        private void RestoreContentSlidePositions(Presentation pres, List<int> newDividerIds, Dictionary<string, List<int>> contentMap, List<AgendaItemData> items)
        {
            foreach (var topic in contentMap.Keys)
            {
                // Find the new divider for this topic
                Slide newDivider = null;
                foreach (int id in newDividerIds)
                {
                    try {
                        var s = pres.Slides.FindBySlideID(id);
                        if (s.Tags["oPE_AgendaTopic"] == topic)
                        {
                            newDivider = s;
                            break;
                        }
                    } catch { }
                }

                if (newDivider != null)
                {
                    var slideIdsToMove = contentMap[topic];
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
                        } catch { }
                    }
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
                    ApplyShapeFormat(titleShape, _layout.HeadlineShape);
                    titleShape.Tags.Add("oPE_AgendaRole", "Headline");
                }
                else if (slide.Shapes.HasTitle == Office.MsoTriState.msoTrue) {
                    slide.Shapes.Title.TextFrame.TextRange.Text = title;
                }
            } catch { }
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
                    } catch { }
                }
            } catch {}
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
                    } catch { }
                }
            } catch {}
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
                    } catch { }
                }
            } catch {}
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
                textShape = CreateShape(_layout.ItemTextShape, topicText, "Topic", isActive, linkTarget);
                textShape.Top = startY;
                if (itemX.HasValue) textShape.Left = itemX.Value;
                bottomY = Math.Max(bottomY, textShape.Top + textShape.Height);
            }

            if (_layout.ItemNumberingShape != null)
            {
                var numShape = CreateShape(_layout.ItemNumberingShape, displayNumber, "ItemNo", isActive, linkTarget);
                numShape.Top = startY + (_layout.AlignInfoShapesWithItem ? 0 : (_layout.ItemNumberingShape.Top - _layout.ItemTextShape.Top));
                if (textShape != null) numShape.Left = textShape.Left + (_layout.ItemNumberingShape.Left - _layout.ItemTextShape.Left);
                bottomY = Math.Max(bottomY, numShape.Top + numShape.Height);
            }

            if (_layout.ItemTimeShape != null && !string.IsNullOrEmpty(item.TimeSlot))
            {
                var timeShape = CreateShape(_layout.ItemTimeShape, item.TimeSlot, "TimeSlot", isActive);
                PositionInfoShape(timeShape, textShape, "Time");
                bottomY = Math.Max(bottomY, timeShape.Top + timeShape.Height);
            }

            if (_layout.ItemPageNumberShape != null && !string.IsNullOrEmpty(item.PageNumber))
            {
                var pageShape = CreateShape(_layout.ItemPageNumberShape, item.PageNumber, "PageNumber", isActive, linkTarget);
                PositionInfoShape(pageShape, textShape, "PageNumber");
                bottomY = Math.Max(bottomY, pageShape.Top + pageShape.Height);
            }

            if (_layout.ItemBackgroundShape != null)
            {
                 var bgShape = CreateShape(_layout.ItemBackgroundShape, "", "Background", isActive);
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
                textShape = CreateShape(textConfig, topicText, "Topic", isActive);
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
                    var numShape = CreateShape(numConfig, displayNumber, "ItemNo", isActive);
                    numShape.Top = startY + (_layout.AlignInfoShapesWithItem ? 0 : (numConfig.Top - textConfig.Top));
                    if (textShape != null) numShape.Left = textShape.Left + (numConfig.Left - textConfig.Left);
                    bottomY = Math.Max(bottomY, numShape.Top + numShape.Height);
                }
            }

            if (_layout.SubItemBackgroundShape != null || _layout.ItemBackgroundShape != null)
            {
                 var bgConfig = _layout.SubItemBackgroundShape ?? _layout.ItemBackgroundShape;
                 var bgShape = CreateShape(bgConfig, "", "Background", isActive);
                 bgShape.Top = startY;
                 bgShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
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
            }
        }

        private Shape CreateShape(ShapeFormatConfig config, string text, string role = null, bool isActive = false, Slide linkToSlide = null)
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
                    } catch { }
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
                string guid = Guid.NewGuid().ToString("D");
                shape.Name = $"ee4p_item_{guid}_shape";
                shape.Tags.Add("oPE_AgendaRole", role);
                shape.Tags.Add("EE4P_AGENDAWIZARD", role);
            }
            return shape;
        }
    }

    public class AgendaItemData
    {
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
