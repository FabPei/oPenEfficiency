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
        private void AttachLayoutContextMenus(ButtonBase btn)
        {
            if (btn.Name == "BtnAlignLeft" || btn.Name == "BtnAlignRight" || btn.Name == "BtnAlignTop" || btn.Name == "BtnAlignBottom")
            {
                btn.ContextMenu.Items.Add(new Separator());

                string direction;
                Action alignWithinSelection;
                Action alignToSlide;
                Action alignToFirst;
                Action alignToLast;
                switch (btn.Name)
                {
                    case "BtnAlignLeft":
                        direction = "left";
                        alignWithinSelection = () => AlignToFirstLeftFeature.ExecuteWithinSelection(GetManager());
                        alignToSlide = () => AlignToFirstLeftFeature.ExecuteToSlide(GetManager());
                        alignToFirst = () => AlignToFirstLeftFeature.ExecuteToFirstSelected(GetManager());
                        alignToLast = () => AlignToFirstLeftFeature.ExecuteToLastSelected(GetManager());
                        break;
                    case "BtnAlignRight":
                        direction = "right";
                        alignWithinSelection = () => AlignToFirstRightFeature.ExecuteWithinSelection(GetManager());
                        alignToSlide = () => AlignToFirstRightFeature.ExecuteToSlide(GetManager());
                        alignToFirst = () => AlignToFirstRightFeature.ExecuteToFirstSelected(GetManager());
                        alignToLast = () => AlignToFirstRightFeature.ExecuteToLastSelected(GetManager());
                        break;
                    case "BtnAlignTop":
                        direction = "top";
                        alignWithinSelection = () => AlignToFirstTopFeature.ExecuteWithinSelection(GetManager());
                        alignToSlide = () => AlignToFirstTopFeature.ExecuteToSlide(GetManager());
                        alignToFirst = () => AlignToFirstTopFeature.ExecuteToFirstSelected(GetManager());
                        alignToLast = () => AlignToFirstTopFeature.ExecuteToLastSelected(GetManager());
                        break;
                    default:
                        direction = "bottom";
                        alignWithinSelection = () => AlignToFirstBottomFeature.ExecuteWithinSelection(GetManager());
                        alignToSlide = () => AlignToFirstBottomFeature.ExecuteToSlide(GetManager());
                        alignToFirst = () => AlignToFirstBottomFeature.ExecuteToFirstSelected(GetManager());
                        alignToLast = () => AlignToFirstBottomFeature.ExecuteToLastSelected(GetManager());
                        break;
                }

                var withinSelection = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Align Within Selection") };
                withinSelection.Click += (s, e) => alignWithinSelection();
                btn.ContextMenu.Items.Add(withinSelection);
                var toSlide = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Align to Slide") };
                toSlide.Click += (s, e) => alignToSlide();
                btn.ContextMenu.Items.Add(toSlide);
                var toFirst = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Align to First Selected") };
                toFirst.Click += (s, e) => alignToFirst();
                btn.ContextMenu.Items.Add(toFirst);
                var toLast = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Align to Last Selected") };
                toLast.Click += (s, e) => alignToLast();
                btn.ContextMenu.Items.Add(toLast);
                btn.ContextMenu.Items.Add(new Separator());
                var miCollision = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 4, "Move Until Next Object") };
                var capturedDir = direction;
                miCollision.Click += (s, e) =>
                {
                    Features.AlignmentHelpers.ExecuteMoveUntilCollision(GetManager(), capturedDir);
                };
                btn.ContextMenu.Items.Add(miCollision);
            }
            else if (btn.Name == "BtnAlignCenter")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var withinSelection = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Align Centers Within Selection") };
                withinSelection.Click += (s, e) => AlignToFirstCenterHorizontalFeature.ExecuteWithinSelection(GetManager());
                btn.ContextMenu.Items.Add(withinSelection);
                var toSlide = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Align Centers to Slide") };
                toSlide.Click += (s, e) => AlignToFirstCenterHorizontalFeature.ExecuteToSlide(GetManager());
                btn.ContextMenu.Items.Add(toSlide);
                var nearestGuide = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Snap Centers to Nearest Guide") };
                nearestGuide.Click += (s, e) => AlignToFirstCenterHorizontalFeature.ExecuteToNearestGuide(GetManager());
                btn.ContextMenu.Items.Add(nearestGuide);
                var toAnchor = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Align to Configured Anchor") };
                toAnchor.Click += (s, e) => AlignToFirstCenterHorizontalFeature.ExecuteToAnchor(GetManager());
                btn.ContextMenu.Items.Add(toAnchor);
            }
            else if (btn.Name == "BtnAlignCenterVertical")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var withinSelection = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Align Middles Within Selection") };
                withinSelection.Click += (s, e) => AlignToFirstCenterVerticalFeature.ExecuteWithinSelection(GetManager());
                btn.ContextMenu.Items.Add(withinSelection);
                var toSlide = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Align Middles to Slide") };
                toSlide.Click += (s, e) => AlignToFirstCenterVerticalFeature.ExecuteToSlide(GetManager());
                btn.ContextMenu.Items.Add(toSlide);
                var nearestGuide = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Snap Middles to Nearest Guide") };
                nearestGuide.Click += (s, e) => AlignToFirstCenterVerticalFeature.ExecuteToNearestGuide(GetManager());
                btn.ContextMenu.Items.Add(nearestGuide);
                var toAnchor = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Align to Configured Anchor") };
                toAnchor.Click += (s, e) => AlignToFirstCenterVerticalFeature.ExecuteToAnchor(GetManager());
                btn.ContextMenu.Items.Add(toAnchor);
            }
            else if (btn.Name == "BtnArrangeGrid")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var two = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Arrange in 2 Columns") };
                two.Click += (s, e) => ArrangeGridFeature.ExecuteColumns(GetManager(), 2);
                btn.ContextMenu.Items.Add(two);
                var three = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Arrange in 3 Columns") };
                three.Click += (s, e) => ArrangeGridFeature.ExecuteColumns(GetManager(), 3);
                btn.ContextMenu.Items.Add(three);
                var four = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Arrange in 4 Columns") };
                four.Click += (s, e) => ArrangeGridFeature.ExecuteColumns(GetManager(), 4);
                btn.ContextMenu.Items.Add(four);
                var five = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Arrange in 5 Columns") };
                five.Click += (s, e) => ArrangeGridFeature.ExecuteColumns(GetManager(), 5);
                btn.ContextMenu.Items.Add(five);
            }
            else if (btn.Name == "BtnFitFormToText")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var zeroMargins = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Fit with Zero Margins") };
                zeroMargins.Click += (s, e) => FitFormToTextFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(zeroMargins);
                var keepMargins = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Fit and Keep Existing Margins") };
                keepMargins.Click += (s, e) => FitFormToTextFeature.Execute(GetManager(), false);
                btn.ContextMenu.Items.Add(keepMargins);
            }
            else if (btn.Name == "BtnIncreaseFontSize")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var one = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Increase 1 Step") };
                one.Click += (s, e) => IncreaseFontSizeFeature.ExecuteSteps(GetManager(), 1);
                btn.ContextMenu.Items.Add(one);
                var two = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Increase 2 Steps") };
                two.Click += (s, e) => IncreaseFontSizeFeature.ExecuteSteps(GetManager(), 2);
                btn.ContextMenu.Items.Add(two);
                var five = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Increase 5 Steps") };
                five.Click += (s, e) => IncreaseFontSizeFeature.ExecuteSteps(GetManager(), 5);
                btn.ContextMenu.Items.Add(five);
            }
            else if (btn.Name == "BtnDecreaseFontSize")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var one = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Decrease 1 Step") };
                one.Click += (s, e) => DecreaseFontSizeFeature.ExecuteSteps(GetManager(), 1);
                btn.ContextMenu.Items.Add(one);
                var two = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Decrease 2 Steps") };
                two.Click += (s, e) => DecreaseFontSizeFeature.ExecuteSteps(GetManager(), 2);
                btn.ContextMenu.Items.Add(two);
                var five = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Decrease 5 Steps") };
                five.Click += (s, e) => DecreaseFontSizeFeature.ExecuteSteps(GetManager(), 5);
                btn.ContextMenu.Items.Add(five);
            }
            else if (btn.Name == "BtnSmartCorners")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var rounded = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Make Corners Rounded") };
                rounded.Click += (s, e) => SmartCornersFeature.MakeRounded(GetManager());
                btn.ContextMenu.Items.Add(rounded);
                var square = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Make Corners Square") };
                square.Click += (s, e) => SmartCornersFeature.MakeSquare(GetManager());
                btn.ContextMenu.Items.Add(square);
                var toggle = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Toggle Rounded Corners") };
                toggle.Click += (s, e) => SmartCornersFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(toggle);
            }
            else if (btn.Name == "BtnNewSlide")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var gallery = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Open New Slide Gallery") };
                gallery.Click += (s, e) => NewSlideFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(gallery);
                var blank = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Insert Blank Slide") };
                blank.Click += (s, e) => NewSlideFeature.InsertBlankAfterCurrent(GetManager());
                btn.ContextMenu.Items.Add(blank);
                var duplicate = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Duplicate Current Slide") };
                duplicate.Click += (s, e) => NewSlideFeature.DuplicateCurrent(GetManager());
                btn.ContextMenu.Items.Add(duplicate);
            }
            else if (btn.Name == "BtnTableColumnInsertion")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var left = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Insert Column Left") };
                left.Click += (s, e) => TableColumnInsertionFeature.Execute(GetManager(), true, false);
                btn.ContextMenu.Items.Add(left);
                var right = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Insert Column Right") };
                right.Click += (s, e) => TableColumnInsertionFeature.Execute(GetManager(), false, false);
                btn.ContextMenu.Items.Add(right);
                var leftPreserve = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Insert Left and Preserve Width") };
                leftPreserve.Click += (s, e) => TableColumnInsertionFeature.Execute(GetManager(), true, true);
                btn.ContextMenu.Items.Add(leftPreserve);
                var rightPreserve = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Insert Right and Preserve Width") };
                rightPreserve.Click += (s, e) => TableColumnInsertionFeature.Execute(GetManager(), false, true);
                btn.ContextMenu.Items.Add(rightPreserve);
                var duplicateLeft = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 4, "Duplicate Left Column Format") };
                duplicateLeft.Click += (s, e) => TableColumnInsertionFeature.DuplicateColumnFormat(GetManager(), true);
                btn.ContextMenu.Items.Add(duplicateLeft);
                var duplicateRight = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 5, "Duplicate Right Column Format") };
                duplicateRight.Click += (s, e) => TableColumnInsertionFeature.DuplicateColumnFormat(GetManager(), false);
                btn.ContextMenu.Items.Add(duplicateRight);
            }
            else if (btn.Name == "BtnTableRowInsertion")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var above = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Insert Row Above") };
                above.Click += (s, e) => TableRowInsertionFeature.Execute(GetManager(), true, false);
                btn.ContextMenu.Items.Add(above);
                var below = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Insert Row Below") };
                below.Click += (s, e) => TableRowInsertionFeature.Execute(GetManager(), false, false);
                btn.ContextMenu.Items.Add(below);
                var abovePreserve = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Insert Above and Preserve Height") };
                abovePreserve.Click += (s, e) => TableRowInsertionFeature.Execute(GetManager(), true, true);
                btn.ContextMenu.Items.Add(abovePreserve);
                var belowPreserve = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Insert Below and Preserve Height") };
                belowPreserve.Click += (s, e) => TableRowInsertionFeature.Execute(GetManager(), false, true);
                btn.ContextMenu.Items.Add(belowPreserve);
                var duplicateAbove = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 4, "Duplicate Above Row Format") };
                duplicateAbove.Click += (s, e) => TableRowInsertionFeature.DuplicateRowFormat(GetManager(), true);
                btn.ContextMenu.Items.Add(duplicateAbove);
                var duplicateBelow = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 5, "Duplicate Below Row Format") };
                duplicateBelow.Click += (s, e) => TableRowInsertionFeature.DuplicateRowFormat(GetManager(), false);
                btn.ContextMenu.Items.Add(duplicateBelow);
            }
            else if (btn.Name == "BtnTableSortAZ")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var az = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Sort Selected Cells A–Z") };
                az.Click += (s, e) => TableSortFeature.Execute(GetManager(), TableSortFeature.SortType.AZ_Selection);
                btn.ContextMenu.Items.Add(az);
                var za = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Sort Selected Cells Z–A") };
                za.Click += (s, e) => TableSortFeature.Execute(GetManager(), TableSortFeature.SortType.ZA_Selection);
                btn.ContextMenu.Items.Add(za);
                var rowsAz = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Sort Whole Row Text A–Z") };
                rowsAz.Click += (s, e) => TableSortFeature.Execute(GetManager(), TableSortFeature.SortType.AZ_WholeRow);
                btn.ContextMenu.Items.Add(rowsAz);
                var rowsZa = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Sort Whole Row Text Z–A") };
                rowsZa.Click += (s, e) => TableSortFeature.Execute(GetManager(), TableSortFeature.SortType.ZA_WholeRow);
                btn.ContextMenu.Items.Add(rowsZa);
            }
            else if (btn.Name == "BtnSlideSize")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var analyze = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Analyze Slide Media Sizes") };
                analyze.Click += (s, e) => SlideSizeFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(analyze);
                var remove = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Remove Analysis Labels") };
                remove.Click += (s, e) => SlideSizeFeature.RemoveAnalysisLabels(GetManager());
                btn.ContextMenu.Items.Add(remove);
            }
            else if (btn.Name == "BtnInsertText")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var lineBreak = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Insert Line Break") };
                lineBreak.Click += (s, e) => InsertTextAtCursorFeature.Execute(GetManager(), "\r");
                btn.ContextMenu.Items.Add(lineBreak);
                var bullet = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Insert Bullet") };
                bullet.Click += (s, e) => InsertTextAtCursorFeature.Execute(GetManager(), "• ");
                btn.ContextMenu.Items.Add(bullet);
                var enDash = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Insert En Dash") };
                enDash.Click += (s, e) => InsertTextAtCursorFeature.Execute(GetManager(), "–");
                btn.ContextMenu.Items.Add(enDash);
                var emDash = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Insert Em Dash") };
                emDash.Click += (s, e) => InsertTextAtCursorFeature.Execute(GetManager(), "—");
                btn.ContextMenu.Items.Add(emDash);
            }
            else if (btn.Name == "BtnObjectConnector")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var smart = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Connect Nearest Sides (Smart)") };
                smart.Click += (s, e) => ObjectConnectorFeature.Execute(GetManager(), autoDetect: true, smart: true);
                btn.ContextMenu.Items.Add(smart);
                var rightLeft = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Connect Right to Left") };
                rightLeft.Click += (s, e) => ObjectConnectorFeature.Execute(GetManager(), "Right", "Left");
                btn.ContextMenu.Items.Add(rightLeft);
                var leftRight = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Connect Left to Right") };
                leftRight.Click += (s, e) => ObjectConnectorFeature.Execute(GetManager(), "Left", "Right");
                btn.ContextMenu.Items.Add(leftRight);
                var bottomTop = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Connect Bottom to Top") };
                bottomTop.Click += (s, e) => ObjectConnectorFeature.Execute(GetManager(), "Bottom", "Top");
                btn.ContextMenu.Items.Add(bottomTop);
                var topBottom = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 4, "Connect Top to Bottom") };
                topBottom.Click += (s, e) => ObjectConnectorFeature.Execute(GetManager(), "Top", "Bottom");
                btn.ContextMenu.Items.Add(topBottom);
            }
            else if (btn.Name == "BtnTagInspector")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var slide = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Inspect Current Slide Tags") };
                slide.Click += (s, e) => TagInspectorFeature.Execute(GetManager(), TagInspectorWindow.InspectorScope.CurrentSlide);
                btn.ContextMenu.Items.Add(slide);
                var shapes = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Inspect Selected Shape Tags") };
                shapes.Click += (s, e) => TagInspectorFeature.Execute(GetManager(), TagInspectorWindow.InspectorScope.SelectedShapes);
                btn.ContextMenu.Items.Add(shapes);
                var presentation = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Inspect Presentation Tags") };
                presentation.Click += (s, e) => TagInspectorFeature.Execute(GetManager(), TagInspectorWindow.InspectorScope.Presentation);
                btn.ContextMenu.Items.Add(presentation);
            }
            else if (btn.Name == "BtnFormatNumbers")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var us = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Format Numbers: US (1,234.56)") };
                us.Click += (s, e) => FormatNumbersFeature.Execute(GetManager(), ",", ".");
                btn.ContextMenu.Items.Add(us);
                var german = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Format Numbers: German (1.234,56)") };
                german.Click += (s, e) => FormatNumbersFeature.Execute(GetManager(), ".", ",");
                btn.ContextMenu.Items.Add(german);
                var french = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Format Numbers: French (1 234,56)") };
                french.Click += (s, e) => FormatNumbersFeature.Execute(GetManager(), " ", ",");
                btn.ContextMenu.Items.Add(french);
                var swiss = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Format Numbers: Swiss (1'234.56)") };
                swiss.Click += (s, e) => FormatNumbersFeature.Execute(GetManager(), "'", ".");
                btn.ContextMenu.Items.Add(swiss);
            }
            else if (btn.Name == "BtnSplitShape")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var oneByTwo = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Split 1 × 2") };
                oneByTwo.Click += (s, e) => SplitShapeToGridFeature.ExecutePreset(GetManager(), 1, 2);
                btn.ContextMenu.Items.Add(oneByTwo);
                var twoByOne = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Split 2 × 1") };
                twoByOne.Click += (s, e) => SplitShapeToGridFeature.ExecutePreset(GetManager(), 2, 1);
                btn.ContextMenu.Items.Add(twoByOne);
                var twoByTwo = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Split 2 × 2") };
                twoByTwo.Click += (s, e) => SplitShapeToGridFeature.ExecutePreset(GetManager(), 2, 2);
                btn.ContextMenu.Items.Add(twoByTwo);
                var threeByTwo = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Split 3 × 2") };
                threeByTwo.Click += (s, e) => SplitShapeToGridFeature.ExecutePreset(GetManager(), 3, 2);
                btn.ContextMenu.Items.Add(threeByTwo);
                var threeByThree = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 4, "Split 3 × 3") };
                threeByThree.Click += (s, e) => SplitShapeToGridFeature.ExecutePreset(GetManager(), 3, 3);
                btn.ContextMenu.Items.Add(threeByThree);
            }
            else if (btn.Name == "BtnSetMargin")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var all = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Paste All Text Margins") };
                all.Click += (s, e) => Features.Visuals.SetMarginFeature.Execute(GetManager(), true, true, false);
                btn.ContextMenu.Items.Add(all);
                var horizontal = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Paste Left and Right Text Margins") };
                horizontal.Click += (s, e) => Features.Visuals.SetMarginFeature.Execute(GetManager(), true, false, false);
                btn.ContextMenu.Items.Add(horizontal);
                var vertical = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Paste Top and Bottom Text Margins") };
                vertical.Click += (s, e) => Features.Visuals.SetMarginFeature.Execute(GetManager(), false, true, false);
                btn.ContextMenu.Items.Add(vertical);
                btn.ContextMenu.Items.Add(new Separator());
                var reset = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Reset Text Margins to Zero") };
                reset.Click += (s, e) => Features.Visuals.SetMarginFeature.Execute(GetManager(), true, true, true);
                btn.ContextMenu.Items.Add(reset);
            }
            else if (btn.Name == "BtnLearnMargin")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var learn = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Copy Text Margins from Selection") };
                learn.Click += (s, e) => Features.Visuals.LearnMarginFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(learn);
                var clear = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Clear Copied Text Margins") };
                clear.Click += (s, e) => Features.Visuals.LearnMarginFeature.ClearLearnedMargins();
                btn.ContextMenu.Items.Add(clear);
            }
            else if (btn.Name == "BtnTransparentColor")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var exact = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Pick Color: Exact Match (0%)") };
                exact.Click += (s, e) => TransparentColorFeature.Execute(GetManager(), 0);
                btn.ContextMenu.Items.Add(exact);
                var balanced = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Pick Color: Balanced Tolerance (10%)") };
                balanced.Click += (s, e) => TransparentColorFeature.Execute(GetManager(), 10);
                btn.ContextMenu.Items.Add(balanced);
                var broad = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Pick Color: Broad Tolerance (25%)") };
                broad.Click += (s, e) => TransparentColorFeature.Execute(GetManager(), 25);
                btn.ContextMenu.Items.Add(broad);
            }
            else if (btn.Name == "BtnStickyNote")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var theme = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Insert Theme Sticky Note") };
                theme.Click += (s, e) => StickyNoteFeature.ExecuteStyle(GetManager(), StickyNoteFeature.StickyNoteStyle.Theme);
                btn.ContextMenu.Items.Add(theme);
                var consulting = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Insert Consulting Sticky Note") };
                consulting.Click += (s, e) => StickyNoteFeature.ExecuteStyle(GetManager(), StickyNoteFeature.StickyNoteStyle.Consulting);
                btn.ContextMenu.Items.Add(consulting);
                var compact = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Insert Compact Banner Note") };
                compact.Click += (s, e) => StickyNoteFeature.ExecuteStyle(GetManager(), StickyNoteFeature.StickyNoteStyle.CompactBanner);
                btn.ContextMenu.Items.Add(compact);
            }
            else if (btn.Name == "BtnSpellCheck")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var shapes = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Choose Language for Selected Shapes") };
                shapes.Click += (s, e) => SpellCheckLanguageFeature.Execute(GetManager(), SpellCheckScope.SelectedShapes);
                btn.ContextMenu.Items.Add(shapes);
                var slides = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Choose Language for Selected Slides") };
                slides.Click += (s, e) => SpellCheckLanguageFeature.Execute(GetManager(), SpellCheckScope.SelectedSlides);
                btn.ContextMenu.Items.Add(slides);
                var presentation = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Choose Language for Entire Presentation") };
                presentation.Click += (s, e) => SpellCheckLanguageFeature.Execute(GetManager(), SpellCheckScope.Presentation);
                btn.ContextMenu.Items.Add(presentation);
            }
            else if (btn.Name == "BtnOptimizeFreeForm")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var light = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Light Cleanup (2 pt)") };
                light.Click += (s, e) => Features.Shapes.OptimizeFreeFormFeature.Execute(GetManager(), 2f);
                btn.ContextMenu.Items.Add(light);
                var standard = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Standard Cleanup (5 pt)") };
                standard.Click += (s, e) => Features.Shapes.OptimizeFreeFormFeature.Execute(GetManager(), 5f);
                btn.ContextMenu.Items.Add(standard);
                var aggressive = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Aggressive Cleanup (10 pt)") };
                aggressive.Click += (s, e) => Features.Shapes.OptimizeFreeFormFeature.Execute(GetManager(), 10f);
                btn.ContextMenu.Items.Add(aggressive);
            }
            else if (btn.Name == "BtnRefreshLinkedTable")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var refresh = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Refresh Linked Table") };
                refresh.Click += (s, e) => RefreshLinkedTableFeature.ExecuteAction(GetManager(), RefreshLinkedTableFeature.LinkedTableAction.Refresh);
                btn.ContextMenu.Items.Add(refresh);
                var open = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Open Source in Excel") };
                open.Click += (s, e) => RefreshLinkedTableFeature.ExecuteAction(GetManager(), RefreshLinkedTableFeature.LinkedTableAction.OpenSource);
                btn.ContextMenu.Items.Add(open);
                var range = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Change Linked Range...") };
                range.Click += (s, e) => RefreshLinkedTableFeature.ExecuteAction(GetManager(), RefreshLinkedTableFeature.LinkedTableAction.ChangeRange);
                btn.ContextMenu.Items.Add(range);
                var relink = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Relink Source...") };
                relink.Click += (s, e) => RefreshLinkedTableFeature.ExecuteAction(GetManager(), RefreshLinkedTableFeature.LinkedTableAction.RelinkSource);
                btn.ContextMenu.Items.Add(relink);
                var autoResize = new MenuItem
                {
                    Header = GetCatalogVariantLabel(btn.Name, 4, "Toggle Auto-Resize"),
                    IsCheckable = true,
                    IsChecked = RefreshLinkedTableFeature.IsAutoResizeEnabled(GetManager())
                };
                autoResize.Click += (s, e) => RefreshLinkedTableFeature.ExecuteAction(GetManager(), RefreshLinkedTableFeature.LinkedTableAction.ToggleAutoResize);
                btn.ContextMenu.Items.Add(autoResize);
            }
            else if (btn.Name == "BtnSettings")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var general = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Open General Settings") };
                general.Click += (s, e) => OpenSettings(0);
                btn.ContextMenu.Items.Add(general);
                var layout = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Open Layout Settings") };
                layout.Click += (s, e) => OpenSettings(1);
                btn.ContextMenu.Items.Add(layout);
                var appearance = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Open Appearance Settings") };
                appearance.Click += (s, e) => OpenSettings(2);
                btn.ContextMenu.Items.Add(appearance);
                var notes = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Open Stickers & Notes Settings") };
                notes.Click += (s, e) => OpenSettings(3);
                btn.ContextMenu.Items.Add(notes);
                var shortcuts = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 4, "Open Shortcut Settings") };
                shortcuts.Click += (s, e) => OpenSettings(4);
                btn.ContextMenu.Items.Add(shortcuts);
                var backup = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 5, "Open Backup & Config Settings") };
                backup.Click += (s, e) => OpenSettings(5);
                btn.ContextMenu.Items.Add(backup);
                var about = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 6, "Open About") };
                about.Click += (s, e) => OpenSettings(6);
                btn.ContextMenu.Items.Add(about);
            }
            else if (btn.Name == "BtnExploreFeatures")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var search = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Search All Features") };
                search.Click += (s, e) => _host?.OpenFeatureExplorer(null, true);
                btn.ContextMenu.Items.Add(search);
                var align = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Explore Align & Size") };
                align.Click += (s, e) => _host?.OpenFeatureExplorer("Align & Size");
                btn.ContextMenu.Items.Add(align);
                var text = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Explore Text & Format") };
                text.Click += (s, e) => _host?.OpenFeatureExplorer("Text & Format");
                btn.ContextMenu.Items.Add(text);
                var data = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Explore Data & Automation") };
                data.Click += (s, e) => _host?.OpenFeatureExplorer("Data & Automation");
                btn.ContextMenu.Items.Add(data);
                var labs = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 4, "Explore Labs") };
                labs.Click += (s, e) => _host?.OpenFeatureExplorer("Labs");
                btn.ContextMenu.Items.Add(labs);
            }
            else if (btn.Name == "BtnApplyTextTool")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var editor = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Open Text Editor") };
                editor.Click += (s, e) => OpenApplyTextEditor();
                btn.ContextMenu.Items.Add(editor);
                var first = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Replace Text from First Selected") };
                first.Click += (s, e) => ApplyTextFromMaster(false);
                btn.ContextMenu.Items.Add(first);
                var last = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Replace Text from Last Selected") };
                last.Click += (s, e) => ApplyTextFromMaster(true);
                btn.ContextMenu.Items.Add(last);
            }
            else if (btn.Name == "BtnReplaceFont")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var current = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Replace Formats on Current Slide") };
                current.Click += (s, e) => BulkFormatReplacerFeature.Execute(GetManager(), BulkFormatReplacerFeature.ReplacementScope.CurrentSlide);
                btn.ContextMenu.Items.Add(current);
                var selected = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Replace Formats on Selected Slides") };
                selected.Click += (s, e) => BulkFormatReplacerFeature.Execute(GetManager(), BulkFormatReplacerFeature.ReplacementScope.SelectedSlides);
                btn.ContextMenu.Items.Add(selected);
                var entire = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Replace Formats in Entire Presentation") };
                entire.Click += (s, e) => BulkFormatReplacerFeature.Execute(GetManager(), BulkFormatReplacerFeature.ReplacementScope.EntirePresentation);
                btn.ContextMenu.Items.Add(entire);
            }
            else if (btn.Name == "BtnIconBrowser")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var search = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Search Icons") };
                search.Click += (s, e) => _host?.OpenIconBrowser(IconBrowserStartView.Search);
                btn.ContextMenu.Items.Add(search);
                var favorites = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Browse Favorite Icons") };
                favorites.Click += (s, e) => _host?.OpenIconBrowser(IconBrowserStartView.Favorites);
                btn.ContextMenu.Items.Add(favorites);
                var svg = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Insert Icons as SVG") };
                svg.Click += (s, e) => _host?.OpenIconBrowser(IconBrowserStartView.Svg);
                btn.ContextMenu.Items.Add(svg);
                var shapes = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Insert Icons as PowerPoint Shapes") };
                shapes.Click += (s, e) => _host?.OpenIconBrowser(IconBrowserStartView.PowerPointShape);
                btn.ContextMenu.Items.Add(shapes);
            }
            else if (btn.Name == "BtnWinnerPicker")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var one = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Pick 1 Winner") };
                one.Click += (s, e) => LotteryFeature.PickWinners(GetManager(), 1);
                btn.ContextMenu.Items.Add(one);
                var two = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Pick 2 Winners") };
                two.Click += (s, e) => LotteryFeature.PickWinners(GetManager(), 2);
                btn.ContextMenu.Items.Add(two);
                var three = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Pick 3 Winners") };
                three.Click += (s, e) => LotteryFeature.PickWinners(GetManager(), 3);
                btn.ContextMenu.Items.Add(three);
                btn.ContextMenu.Items.Add(new Separator());
                var clear = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Clear Winner Markers") };
                clear.Click += (s, e) => LotteryFeature.ClearWinnerMarkers();
                btn.ContextMenu.Items.Add(clear);
            }
            else if (btn.Name == "BtnChess")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var easy = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "New Easy Game") };
                easy.Click += (s, e) => ChessFeature.StartNewGame(GetManager(), 1);
                btn.ContextMenu.Items.Add(easy);
                var medium = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "New Medium Game") };
                medium.Click += (s, e) => ChessFeature.StartNewGame(GetManager(), 2);
                btn.ContextMenu.Items.Add(medium);
                var hard = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "New Hard Game") };
                hard.Click += (s, e) => ChessFeature.StartNewGame(GetManager(), 3);
                btn.ContextMenu.Items.Add(hard);
                var close = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Close Chess") };
                close.Click += (s, e) => ChessFeature.CloseGame();
                btn.ContextMenu.Items.Add(close);
            }
            else if (btn.Name == "Btn2048")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var open = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Open / Continue 2048") };
                open.Click += (s, e) => Game2048Feature.Execute(GetManager());
                btn.ContextMenu.Items.Add(open);
                var restart = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Start New 2048 Game") };
                restart.Click += (s, e) => Game2048Feature.StartNewGame(GetManager());
                btn.ContextMenu.Items.Add(restart);
                var close = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Close 2048") };
                close.Click += (s, e) => Game2048Feature.CloseGame();
                btn.ContextMenu.Items.Add(close);
            }
            else if (btn.Name == "BtnMinesweeper")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var open = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Open / Continue Minesweeper") };
                open.Click += (s, e) => MinesweeperFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(open);
                var restart = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Start New Minesweeper Game") };
                restart.Click += (s, e) => MinesweeperFeature.StartNewGame(GetManager());
                btn.ContextMenu.Items.Add(restart);
                var close = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Close Minesweeper") };
                close.Click += (s, e) => MinesweeperFeature.CloseGame();
                btn.ContextMenu.Items.Add(close);
            }
            else if (btn.Name == "BtnArrangeInShape")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var circle = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Arrange Shapes in Circle") };
                circle.Click += (s, e) => ArrangeInShapeFeature.Execute(GetManager(), "circle");
                btn.ContextMenu.Items.Add(circle);
                var triangle = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Arrange in Triangle") };
                triangle.Click += (s, e) => ArrangeInShapeFeature.Execute(GetManager(), "triangle");
                btn.ContextMenu.Items.Add(triangle);
                var square = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Arrange in Square") };
                square.Click += (s, e) => ArrangeInShapeFeature.Execute(GetManager(), "square");
                btn.ContextMenu.Items.Add(square);
                var star = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Arrange in Star") };
                star.Click += (s, e) => ArrangeInShapeFeature.Execute(GetManager(), "star");
                btn.ContextMenu.Items.Add(star);
                var pentagon = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 4, "Arrange in Pentagon") };
                pentagon.Click += (s, e) => ArrangeInShapeFeature.Execute(GetManager(), "pentagon");
                btn.ContextMenu.Items.Add(pentagon);
                var heart = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 5, "Arrange in Heart") };
                heart.Click += (s, e) => ArrangeInShapeFeature.Execute(GetManager(), "heart");
                btn.ContextMenu.Items.Add(heart);
            }
            else if (btn.Name == "BtnPositionPainter")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var position = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Copy/Paste Position Only") };
                position.Click += (s, e) => Features.Visuals.PositionPainterFeature.Execute(GetManager(), true, true, false, false);
                btn.ContextMenu.Items.Add(position);
                var size = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Copy/Paste Size Only") };
                size.Click += (s, e) => Features.Visuals.PositionPainterFeature.Execute(GetManager(), false, false, true, true);
                btn.ContextMenu.Items.Add(size);
                var both = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Copy/Paste Position and Size") };
                both.Click += (s, e) => Features.Visuals.PositionPainterFeature.Execute(GetManager(), true, true, true, true);
                btn.ContextMenu.Items.Add(both);
            }
            else if (btn.Name == "BtnTableHeatmap")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var rows = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Configure Row-Wise Heatmap") };
                rows.Click += (s, e) => TableHeatmapFeature.Execute(GetManager(), TableHeatmapScope.Row);
                btn.ContextMenu.Items.Add(rows);
                var columns = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Configure Column-Wise Heatmap") };
                columns.Click += (s, e) => TableHeatmapFeature.Execute(GetManager(), TableHeatmapScope.Column);
                btn.ContextMenu.Items.Add(columns);
                var table = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Configure Entire-Table Heatmap") };
                table.Click += (s, e) => TableHeatmapFeature.Execute(GetManager(), TableHeatmapScope.EntireTable);
                btn.ContextMenu.Items.Add(table);
            }
            else if (btn.Name == "BtnMatchSize")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var width = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Match Width") };
                width.Click += (s, e) => MatchSizeFeature.Execute(GetManager(), true, false);
                btn.ContextMenu.Items.Add(width);
                var height = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Match Height") };
                height.Click += (s, e) => MatchSizeFeature.Execute(GetManager(), false, true);
                btn.ContextMenu.Items.Add(height);
                var both = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Match Width and Height") };
                both.Click += (s, e) => MatchSizeFeature.Execute(GetManager(), true, true);
                btn.ContextMenu.Items.Add(both);
            }
            else if (btn.Name == "BtnMatchWidth")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var width = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Match Width") };
                width.Click += (s, e) => MatchSizeFeature.Execute(GetManager(), true, false);
                btn.ContextMenu.Items.Add(width);
                var height = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Match Height") };
                height.Click += (s, e) => MatchSizeFeature.Execute(GetManager(), false, true);
                btn.ContextMenu.Items.Add(height);
                var both = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Match Width and Height") };
                both.Click += (s, e) => MatchSizeFeature.Execute(GetManager(), true, true);
                btn.ContextMenu.Items.Add(both);
            }
            else if (btn.Name == "BtnMatchHeight")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var width = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Match Width") };
                width.Click += (s, e) => MatchSizeFeature.Execute(GetManager(), true, false);
                btn.ContextMenu.Items.Add(width);
                var height = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Match Height") };
                height.Click += (s, e) => MatchSizeFeature.Execute(GetManager(), false, true);
                btn.ContextMenu.Items.Add(height);
                var both = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Match Width and Height") };
                both.Click += (s, e) => MatchSizeFeature.Execute(GetManager(), true, true);
                btn.ContextMenu.Items.Add(both);
                btn.ContextMenu.Items.Add(new Separator());
                var matchAndSwap = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Match Height and Swap Positions") };
                matchAndSwap.Click += (s, e) => MatchHeightToFirstFeature.ExecuteAndSwapPositions(GetManager());
                btn.ContextMenu.Items.Add(matchAndSwap);
                btn.ContextMenu.Opened += (s, e) =>
                {
                    var selection = GetManager()?.GetSelectedShapes();
                    matchAndSwap.IsEnabled = selection != null && selection.Count == 2;
                };
            }
            else if (btn.Name == "BtnCopyXY")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var clear = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Clear Copied XY Coordinates") };
                clear.Click += (s, e) => CopyCoordinatesFeature.ClearCopiedGeometry();
                btn.ContextMenu.Items.Add(clear);
            }
            else if (btn.Name == "BtnRectifyRotation")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var rotation0 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Set Rotation to 0°") };
                rotation0.Click += (s, e) => RectifyRotationFeature.ExecuteSpecific(GetManager(), 0f);
                btn.ContextMenu.Items.Add(rotation0);
                var rotation90 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Set Rotation to 90°") };
                rotation90.Click += (s, e) => RectifyRotationFeature.ExecuteSpecific(GetManager(), 90f);
                btn.ContextMenu.Items.Add(rotation90);
                var rotation180 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Set Rotation to 180°") };
                rotation180.Click += (s, e) => RectifyRotationFeature.ExecuteSpecific(GetManager(), 180f);
                btn.ContextMenu.Items.Add(rotation180);
                var rotation270 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Set Rotation to 270°") };
                rotation270.Click += (s, e) => RectifyRotationFeature.ExecuteSpecific(GetManager(), 270f);
                btn.ContextMenu.Items.Add(rotation270);
            }
            else if (btn.Name == "BtnFlipHorizontal")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var horizontal = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Flip Horizontal") };
                horizontal.Click += (s, e) => oPenEfficiency.Features.Visuals.FlipHorizontalFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(horizontal);
                var vertical = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Flip Vertical") };
                vertical.Click += (s, e) => oPenEfficiency.Features.Visuals.FlipVerticalFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(vertical);
            }
            else if (btn.Name == "BtnFlipVertical")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var horizontal = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Flip Horizontal") };
                horizontal.Click += (s, e) => oPenEfficiency.Features.Visuals.FlipHorizontalFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(horizontal);
                var vertical = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Flip Vertical") };
                vertical.Click += (s, e) => oPenEfficiency.Features.Visuals.FlipVerticalFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(vertical);
            }
            else if (btn.Name == "BtnMatchAngles")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var first = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Use First Selected as Geometry Anchor") };
                first.Click += (s, e) => oPenEfficiency.Features.Shapes.MatchAnglesFeature.ExecuteUsingFirstSelected(GetManager());
                btn.ContextMenu.Items.Add(first);
                var last = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Use Last Selected as Geometry Anchor") };
                last.Click += (s, e) => oPenEfficiency.Features.Shapes.MatchAnglesFeature.ExecuteUsingLastSelected(GetManager());
                btn.ContextMenu.Items.Add(last);
            }
            else if (btn.Name == "BtnAlignShapeAdjustments")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var first = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Use First Selected as Geometry Anchor") };
                first.Click += (s, e) => oPenEfficiency.Features.Shapes.AlignShapeAdjustmentsFeature.ExecuteUsingFirstSelected(GetManager());
                btn.ContextMenu.Items.Add(first);
                var last = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Use Last Selected as Geometry Anchor") };
                last.Click += (s, e) => oPenEfficiency.Features.Shapes.AlignShapeAdjustmentsFeature.ExecuteUsingLastSelected(GetManager());
                btn.ContextMenu.Items.Add(last);
            }
            else if (btn.Name == "BtnCreateMotionPath")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var forward = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "First Shape → Second Shape") };
                forward.Click += (s, e) => oPenEfficiency.Features.Shapes.CreateMotionPathFeature.Execute(GetManager(), oPenEfficiency.Core.Services.MotionPathPolicy.MotionPathMode.FirstToSecond);
                btn.ContextMenu.Items.Add(forward);
                var reverse = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Second Shape → First Shape") };
                reverse.Click += (s, e) => oPenEfficiency.Features.Shapes.CreateMotionPathFeature.Execute(GetManager(), oPenEfficiency.Core.Services.MotionPathPolicy.MotionPathMode.SecondToFirst);
                btn.ContextMenu.Items.Add(reverse);
                var both = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Create Both Directions") };
                both.Click += (s, e) => oPenEfficiency.Features.Shapes.CreateMotionPathFeature.Execute(GetManager(), oPenEfficiency.Core.Services.MotionPathPolicy.MotionPathMode.BothDirections);
                btn.ContextMenu.Items.Add(both);
            }
            else if (btn.Name == "BtnHideSelected")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var hide = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Hide Selected Shapes") };
                hide.Click += (s, e) => VisibilityFeature.HideSelectedShapes(GetManager());
                btn.ContextMenu.Items.Add(hide);
                var showTagged = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Show openEfficiency-Hidden Shapes") };
                showTagged.Click += (s, e) => VisibilityFeature.ShowAllHiddenShapes(GetManager());
                btn.ContextMenu.Items.Add(showTagged);
                btn.ContextMenu.Items.Add(new Separator());
                var revealAll = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Reveal All Hidden Shapes on Slide...") };
                revealAll.Click += (s, e) =>
                {
                    if (System.Windows.MessageBox.Show(
                        "Reveal every hidden shape on the current slide, including shapes hidden outside openEfficiency?",
                        "Reveal All Hidden Shapes", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                        VisibilityFeature.ShowAllHiddenShapes(GetManager(), true);
                };
                btn.ContextMenu.Items.Add(revealAll);
            }
            else if (btn.Name == "BtnHideMasterObjects")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var hide = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Hide Master Objects") };
                hide.Click += (s, e) => HideMasterObjectsFeature.Execute(GetManager(), HideMasterObjectsFeature.MasterVisibilityMode.Hide);
                btn.ContextMenu.Items.Add(hide);
                var show = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Show Master Objects") };
                show.Click += (s, e) => HideMasterObjectsFeature.Execute(GetManager(), HideMasterObjectsFeature.MasterVisibilityMode.Show);
                btn.ContextMenu.Items.Add(show);
                var toggle = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Toggle Master Objects") };
                toggle.Click += (s, e) => HideMasterObjectsFeature.Execute(GetManager(), HideMasterObjectsFeature.MasterVisibilityMode.Toggle);
                btn.ContextMenu.Items.Add(toggle);
            }
            else if (btn.Name == "BtnDockLeft")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var left = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Stack Left") };
                left.Click += (s, e) => DockToFirstLeftFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(left);
                var right = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Stack Right") };
                right.Click += (s, e) => DockToFirstRightFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(right);
                var top = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Stack Above") };
                top.Click += (s, e) => DockToFirstTopFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(top);
                var bottom = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Stack Below") };
                bottom.Click += (s, e) => DockToFirstBottomFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(bottom);
            }
            else if (btn.Name == "BtnDockRight")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var left = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Stack Left") };
                left.Click += (s, e) => DockToFirstLeftFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(left);
                var right = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Stack Right") };
                right.Click += (s, e) => DockToFirstRightFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(right);
                var top = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Stack Above") };
                top.Click += (s, e) => DockToFirstTopFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(top);
                var bottom = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Stack Below") };
                bottom.Click += (s, e) => DockToFirstBottomFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(bottom);
            }
            else if (btn.Name == "BtnDockTop")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var left = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Stack Left") };
                left.Click += (s, e) => DockToFirstLeftFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(left);
                var right = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Stack Right") };
                right.Click += (s, e) => DockToFirstRightFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(right);
                var top = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Stack Above") };
                top.Click += (s, e) => DockToFirstTopFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(top);
                var bottom = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Stack Below") };
                bottom.Click += (s, e) => DockToFirstBottomFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(bottom);
            }
            else if (btn.Name == "BtnDockBottom")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var left = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Stack Left") };
                left.Click += (s, e) => DockToFirstLeftFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(left);
                var right = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Stack Right") };
                right.Click += (s, e) => DockToFirstRightFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(right);
                var top = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Stack Above") };
                top.Click += (s, e) => DockToFirstTopFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(top);
                var bottom = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Stack Below") };
                bottom.Click += (s, e) => DockToFirstBottomFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(bottom);
            }
            else if (btn.Name == "BtnIncreaseHorizontalSpacing")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var increaseHorizontal = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Increase Horizontal Spacing") };
                increaseHorizontal.Click += (s, e) => AdjustHorizontalSpacingFeature.Execute(GetManager(), GetSpacingIncrement());
                btn.ContextMenu.Items.Add(increaseHorizontal);
                var decreaseHorizontal = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Decrease Horizontal Spacing") };
                decreaseHorizontal.Click += (s, e) => AdjustHorizontalSpacingFeature.Execute(GetManager(), -GetSpacingIncrement());
                btn.ContextMenu.Items.Add(decreaseHorizontal);
                var increaseVertical = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Increase Vertical Spacing") };
                increaseVertical.Click += (s, e) => AdjustVerticalSpacingFeature.Execute(GetManager(), GetSpacingIncrement());
                btn.ContextMenu.Items.Add(increaseVertical);
                var decreaseVertical = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Decrease Vertical Spacing") };
                decreaseVertical.Click += (s, e) => AdjustVerticalSpacingFeature.Execute(GetManager(), -GetSpacingIncrement());
                btn.ContextMenu.Items.Add(decreaseVertical);
            }
            else if (btn.Name == "BtnDecreaseHorizontalSpacing")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var increaseHorizontal = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Increase Horizontal Spacing") };
                increaseHorizontal.Click += (s, e) => AdjustHorizontalSpacingFeature.Execute(GetManager(), GetSpacingIncrement());
                btn.ContextMenu.Items.Add(increaseHorizontal);
                var decreaseHorizontal = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Decrease Horizontal Spacing") };
                decreaseHorizontal.Click += (s, e) => AdjustHorizontalSpacingFeature.Execute(GetManager(), -GetSpacingIncrement());
                btn.ContextMenu.Items.Add(decreaseHorizontal);
                var increaseVertical = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Increase Vertical Spacing") };
                increaseVertical.Click += (s, e) => AdjustVerticalSpacingFeature.Execute(GetManager(), GetSpacingIncrement());
                btn.ContextMenu.Items.Add(increaseVertical);
                var decreaseVertical = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Decrease Vertical Spacing") };
                decreaseVertical.Click += (s, e) => AdjustVerticalSpacingFeature.Execute(GetManager(), -GetSpacingIncrement());
                btn.ContextMenu.Items.Add(decreaseVertical);
            }
            else if (btn.Name == "BtnIncreaseVerticalSpacing")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var increaseHorizontal = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Increase Horizontal Spacing") };
                increaseHorizontal.Click += (s, e) => AdjustHorizontalSpacingFeature.Execute(GetManager(), GetSpacingIncrement());
                btn.ContextMenu.Items.Add(increaseHorizontal);
                var decreaseHorizontal = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Decrease Horizontal Spacing") };
                decreaseHorizontal.Click += (s, e) => AdjustHorizontalSpacingFeature.Execute(GetManager(), -GetSpacingIncrement());
                btn.ContextMenu.Items.Add(decreaseHorizontal);
                var increaseVertical = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Increase Vertical Spacing") };
                increaseVertical.Click += (s, e) => AdjustVerticalSpacingFeature.Execute(GetManager(), GetSpacingIncrement());
                btn.ContextMenu.Items.Add(increaseVertical);
                var decreaseVertical = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Decrease Vertical Spacing") };
                decreaseVertical.Click += (s, e) => AdjustVerticalSpacingFeature.Execute(GetManager(), -GetSpacingIncrement());
                btn.ContextMenu.Items.Add(decreaseVertical);
            }
            else if (btn.Name == "BtnDecreaseVerticalSpacing")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var increaseHorizontal = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Increase Horizontal Spacing") };
                increaseHorizontal.Click += (s, e) => AdjustHorizontalSpacingFeature.Execute(GetManager(), GetSpacingIncrement());
                btn.ContextMenu.Items.Add(increaseHorizontal);
                var decreaseHorizontal = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Decrease Horizontal Spacing") };
                decreaseHorizontal.Click += (s, e) => AdjustHorizontalSpacingFeature.Execute(GetManager(), -GetSpacingIncrement());
                btn.ContextMenu.Items.Add(decreaseHorizontal);
                var increaseVertical = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Increase Vertical Spacing") };
                increaseVertical.Click += (s, e) => AdjustVerticalSpacingFeature.Execute(GetManager(), GetSpacingIncrement());
                btn.ContextMenu.Items.Add(increaseVertical);
                var decreaseVertical = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Decrease Vertical Spacing") };
                decreaseVertical.Click += (s, e) => AdjustVerticalSpacingFeature.Execute(GetManager(), -GetSpacingIncrement());
                btn.ContextMenu.Items.Add(decreaseVertical);
            }
            else if (btn.Name == "BtnStretchLeft")
            {
                AddStretchContextMenu(btn, new[] { GetCatalogVariantLabel(btn.Name, 0, "Stretch Left"), GetCatalogVariantLabel(btn.Name, 1, "Stretch Right"), GetCatalogVariantLabel(btn.Name, 2, "Stretch Top"), GetCatalogVariantLabel(btn.Name, 3, "Stretch Bottom"), GetCatalogVariantLabel(btn.Name, 4, "Move Left Edge"), GetCatalogVariantLabel(btn.Name, 5, "Move Right Edge"), GetCatalogVariantLabel(btn.Name, 6, "Move Top Edge"), GetCatalogVariantLabel(btn.Name, 7, "Move Bottom Edge") });
            }
            else if (btn.Name == "BtnStretchRight")
            {
                AddStretchContextMenu(btn, new[] { GetCatalogVariantLabel(btn.Name, 0, "Stretch Left"), GetCatalogVariantLabel(btn.Name, 1, "Stretch Right"), GetCatalogVariantLabel(btn.Name, 2, "Stretch Top"), GetCatalogVariantLabel(btn.Name, 3, "Stretch Bottom"), GetCatalogVariantLabel(btn.Name, 4, "Move Left Edge"), GetCatalogVariantLabel(btn.Name, 5, "Move Right Edge"), GetCatalogVariantLabel(btn.Name, 6, "Move Top Edge"), GetCatalogVariantLabel(btn.Name, 7, "Move Bottom Edge") });
            }
            else if (btn.Name == "BtnStretchTop")
            {
                AddStretchContextMenu(btn, new[] { GetCatalogVariantLabel(btn.Name, 0, "Stretch Left"), GetCatalogVariantLabel(btn.Name, 1, "Stretch Right"), GetCatalogVariantLabel(btn.Name, 2, "Stretch Top"), GetCatalogVariantLabel(btn.Name, 3, "Stretch Bottom"), GetCatalogVariantLabel(btn.Name, 4, "Move Left Edge"), GetCatalogVariantLabel(btn.Name, 5, "Move Right Edge"), GetCatalogVariantLabel(btn.Name, 6, "Move Top Edge"), GetCatalogVariantLabel(btn.Name, 7, "Move Bottom Edge") });
            }
            else if (btn.Name == "BtnStretchBottom")
            {
                AddStretchContextMenu(btn, new[] { GetCatalogVariantLabel(btn.Name, 0, "Stretch Left"), GetCatalogVariantLabel(btn.Name, 1, "Stretch Right"), GetCatalogVariantLabel(btn.Name, 2, "Stretch Top"), GetCatalogVariantLabel(btn.Name, 3, "Stretch Bottom"), GetCatalogVariantLabel(btn.Name, 4, "Move Left Edge"), GetCatalogVariantLabel(btn.Name, 5, "Move Right Edge"), GetCatalogVariantLabel(btn.Name, 6, "Move Top Edge"), GetCatalogVariantLabel(btn.Name, 7, "Move Bottom Edge") });
            }
            else if (btn.Name == "BtnStretchLeftEdge")
            {
                AddStretchContextMenu(btn, new[] { GetCatalogVariantLabel(btn.Name, 0, "Stretch Left"), GetCatalogVariantLabel(btn.Name, 1, "Stretch Right"), GetCatalogVariantLabel(btn.Name, 2, "Stretch Top"), GetCatalogVariantLabel(btn.Name, 3, "Stretch Bottom"), GetCatalogVariantLabel(btn.Name, 4, "Move Left Edge"), GetCatalogVariantLabel(btn.Name, 5, "Move Right Edge"), GetCatalogVariantLabel(btn.Name, 6, "Move Top Edge"), GetCatalogVariantLabel(btn.Name, 7, "Move Bottom Edge") });
            }
            else if (btn.Name == "BtnStretchRightEdge")
            {
                AddStretchContextMenu(btn, new[] { GetCatalogVariantLabel(btn.Name, 0, "Stretch Left"), GetCatalogVariantLabel(btn.Name, 1, "Stretch Right"), GetCatalogVariantLabel(btn.Name, 2, "Stretch Top"), GetCatalogVariantLabel(btn.Name, 3, "Stretch Bottom"), GetCatalogVariantLabel(btn.Name, 4, "Move Left Edge"), GetCatalogVariantLabel(btn.Name, 5, "Move Right Edge"), GetCatalogVariantLabel(btn.Name, 6, "Move Top Edge"), GetCatalogVariantLabel(btn.Name, 7, "Move Bottom Edge") });
            }
            else if (btn.Name == "BtnStretchTopEdge")
            {
                AddStretchContextMenu(btn, new[] { GetCatalogVariantLabel(btn.Name, 0, "Stretch Left"), GetCatalogVariantLabel(btn.Name, 1, "Stretch Right"), GetCatalogVariantLabel(btn.Name, 2, "Stretch Top"), GetCatalogVariantLabel(btn.Name, 3, "Stretch Bottom"), GetCatalogVariantLabel(btn.Name, 4, "Move Left Edge"), GetCatalogVariantLabel(btn.Name, 5, "Move Right Edge"), GetCatalogVariantLabel(btn.Name, 6, "Move Top Edge"), GetCatalogVariantLabel(btn.Name, 7, "Move Bottom Edge") });
            }
            else if (btn.Name == "BtnStretchBottomEdge")
            {
                AddStretchContextMenu(btn, new[] { GetCatalogVariantLabel(btn.Name, 0, "Stretch Left"), GetCatalogVariantLabel(btn.Name, 1, "Stretch Right"), GetCatalogVariantLabel(btn.Name, 2, "Stretch Top"), GetCatalogVariantLabel(btn.Name, 3, "Stretch Bottom"), GetCatalogVariantLabel(btn.Name, 4, "Move Left Edge"), GetCatalogVariantLabel(btn.Name, 5, "Move Right Edge"), GetCatalogVariantLabel(btn.Name, 6, "Move Top Edge"), GetCatalogVariantLabel(btn.Name, 7, "Move Bottom Edge") });
            }
        }
    }
}
