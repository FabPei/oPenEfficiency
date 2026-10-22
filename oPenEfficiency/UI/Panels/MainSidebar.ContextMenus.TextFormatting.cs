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
        private void AttachTextFormattingContextMenus(ButtonBase btn)
        {
            if (btn.Name == "BtnSplitByParagraphs")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var removeEmpty = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Split and Remove Empty Paragraphs") };
                removeEmpty.Click += (s, e) => SplitByParagraphsFeature.Execute(GetManager(), false);
                btn.ContextMenu.Items.Add(removeEmpty);
                var keepEmpty = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Split and Keep Empty Paragraphs") };
                keepEmpty.Click += (s, e) => SplitByParagraphsFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(keepEmpty);
                btn.ContextMenu.Items.Add(new Separator());
                var mergeText = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Merge Selected Text Boxes") };
                mergeText.Click += (s, e) => SplitByParagraphsFeature.MergeSelectedTextBoxes(GetManager());
                btn.ContextMenu.Items.Add(mergeText);
            }
            else if (btn.Name == "BtnTableBranding")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var miManage = new MenuItem { Header = "Manage Table Styles..." };
                miManage.Click += (s, e) => TableBrandingFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(miManage);
            }
            else if (btn.Name == "BtnTableTranspose")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var replace = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Replace Original in Place") };
                replace.Click += (s, e) => TableTransposeFeature.Execute(GetManager(), TableTransposeFeature.TransposeMode.ReplaceOriginal);
                btn.ContextMenu.Items.Add(replace);
                var copyRight = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Create Copy to the Right") };
                copyRight.Click += (s, e) => TableTransposeFeature.Execute(GetManager(), TableTransposeFeature.TransposeMode.CopyRight);
                btn.ContextMenu.Items.Add(copyRight);
                var copyBelow = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Create Copy Below") };
                copyBelow.Click += (s, e) => TableTransposeFeature.Execute(GetManager(), TableTransposeFeature.TransposeMode.CopyBelow);
                btn.ContextMenu.Items.Add(copyBelow);
            }
            else if (btn.Name == "BtnConvertTableToShapes")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var ask = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Ask After Conversion") };
                ask.Click += (s, e) => ConvertTableToShapesFeature.Execute(GetManager(), ConvertTableToShapesFeature.SourceDisposition.AskAfterConversion);
                btn.ContextMenu.Items.Add(ask);
                var keep = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Convert and Keep Original") };
                keep.Click += (s, e) => ConvertTableToShapesFeature.Execute(GetManager(), ConvertTableToShapesFeature.SourceDisposition.KeepOriginal);
                btn.ContextMenu.Items.Add(keep);
                btn.ContextMenu.Items.Add(new Separator());
                var replace = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Convert and Replace Original") };
                replace.Click += (s, e) => ConvertTableToShapesFeature.Execute(GetManager(), ConvertTableToShapesFeature.SourceDisposition.ReplaceOriginal);
                btn.ContextMenu.Items.Add(replace);
            }
            else if (btn.Name == "BtnTableDimensions")
            {
                var miCopyW = new MenuItem { Header = "Copy Column Widths" };
                miCopyW.Click += (s, e) => TableColumnWidthFeature.Execute(GetManager(), false);
                btn.ContextMenu.Items.Add(miCopyW);

                var miCopyH = new MenuItem { Header = "Copy Row Heights" };
                miCopyH.Click += (s, e) => TableRowHeightFeature.Execute(GetManager(), false);
                btn.ContextMenu.Items.Add(miCopyH);

                var miCopyB = new MenuItem { Header = "Copy Both (Width & Height)" };
                miCopyB.Click += (s, e) => TableDimensionsFeature.Execute(GetManager(), false);
                btn.ContextMenu.Items.Add(miCopyB);

                btn.ContextMenu.Items.Add(new Separator());

                var miPasteW = new MenuItem { Header = "Paste Column Widths" };
                miPasteW.Click += (s, e) => TableColumnWidthFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(miPasteW);

                var miPasteH = new MenuItem { Header = "Paste Row Heights" };
                miPasteH.Click += (s, e) => TableRowHeightFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(miPasteH);

                var miPasteB = new MenuItem { Header = "Paste Both (Width & Height)" };
                miPasteB.Click += (s, e) => TableDimensionsFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(miPasteB);
            }
            else if (btn.Name == "BtnSelectSameType")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var miFill = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Match Fill") };
                miFill.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.FillColor);
                btn.ContextMenu.Items.Add(miFill);

                var miLine = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Match Line") };
                miLine.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.LineColor);
                btn.ContextMenu.Items.Add(miLine);

                var miBoth = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Match Fill and Line") };
                miBoth.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.FillAndLineColor);
                btn.ContextMenu.Items.Add(miBoth);

                var miSize = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Match Size") };
                miSize.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.Size);
                btn.ContextMenu.Items.Add(miSize);

                var miWidth = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 4, "Match Width") };
                miWidth.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.Width);
                btn.ContextMenu.Items.Add(miWidth);

                var miHeight = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 5, "Match Height") };
                miHeight.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.Height);
                btn.ContextMenu.Items.Add(miHeight);

                var miType = new MenuItem { Header = "Select shapes with same type" };
                miType.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.Type);
                btn.ContextMenu.Items.Add(miType);

                var miFormatting = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 6, "Match Type and Formatting") };
                miFormatting.Click += (s, e) => SelectSameTypeFeature.Execute(GetManager(), AdvancedSelectionFeature.SelectSameCriteria.TypeAndFormatting);
                btn.ContextMenu.Items.Add(miFormatting);
            }
            else if (btn.Name == "BtnBringToFront")
            {
                AddZOrderContextMenu(
                    btn,
                    GetCatalogVariantLabel(btn.Name, 0, "Rank Selected Shapes (First Selected on Top)"),
                    GetCatalogVariantLabel(btn.Name, 1, "Place Directly Above Last Selected"),
                    GetCatalogVariantLabel(btn.Name, 2, "Place Directly Below Last Selected"),
                    BringToFrontFeature.ExecuteRanked);
            }
            else if (btn.Name == "BtnBringForward")
            {
                AddZOrderContextMenu(
                    btn,
                    GetCatalogVariantLabel(btn.Name, 0, "Rank Selected Shapes (First Selected on Top)"),
                    GetCatalogVariantLabel(btn.Name, 1, "Place Directly Above Last Selected"),
                    GetCatalogVariantLabel(btn.Name, 2, "Place Directly Below Last Selected"),
                    BringForwardFeature.ExecuteRanked);
            }
            else if (btn.Name == "BtnSendBackward")
            {
                AddZOrderContextMenu(
                    btn,
                    GetCatalogVariantLabel(btn.Name, 0, "Rank Selected Shapes (First Selected at Bottom)"),
                    GetCatalogVariantLabel(btn.Name, 1, "Place Directly Above Last Selected"),
                    GetCatalogVariantLabel(btn.Name, 2, "Place Directly Below Last Selected"),
                    SendBackwardFeature.ExecuteRanked);
            }
            else if (btn.Name == "BtnSendToBack")
            {
                AddZOrderContextMenu(
                    btn,
                    GetCatalogVariantLabel(btn.Name, 0, "Rank Selected Shapes (First Selected at Bottom)"),
                    GetCatalogVariantLabel(btn.Name, 1, "Place Directly Above Last Selected"),
                    GetCatalogVariantLabel(btn.Name, 2, "Place Directly Below Last Selected"),
                    SendToBackFeature.ExecuteRanked);
            }
            else if (btn.Name == "BtnRotateShape")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = "Rotate 90° Left" };
                mi1.Click += (s, e) => oPenEfficiency.Features.Visuals.RotateShapeFeature.ExecuteSpecific(GetManager(), -90f);
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = "Rotate 45° Right" };
                mi2.Click += (s, e) => oPenEfficiency.Features.Visuals.RotateShapeFeature.ExecuteSpecific(GetManager(), 45f);
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = "Rotate 45° Left" };
                mi3.Click += (s, e) => oPenEfficiency.Features.Visuals.RotateShapeFeature.ExecuteSpecific(GetManager(), -45f);
                btn.ContextMenu.Items.Add(mi3);
            }
            else if (btn.Name == "BtnChangeCase")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Sentence case") };
                mi1.Click += (s, e) => oPenEfficiency.Features.Text.ChangeCaseFeature.ExecuteSpecific(GetManager(), PowerPoint.PpChangeCase.ppCaseSentence);
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "lowercase") };
                mi2.Click += (s, e) => oPenEfficiency.Features.Text.ChangeCaseFeature.ExecuteSpecific(GetManager(), PowerPoint.PpChangeCase.ppCaseLower);
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "UPPERCASE") };
                mi3.Click += (s, e) => oPenEfficiency.Features.Text.ChangeCaseFeature.ExecuteSpecific(GetManager(), PowerPoint.PpChangeCase.ppCaseUpper);
                btn.ContextMenu.Items.Add(mi3);
                var mi4 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Capitalize Each Word") };
                mi4.Click += (s, e) => oPenEfficiency.Features.Text.ChangeCaseFeature.ExecuteSpecific(GetManager(), PowerPoint.PpChangeCase.ppCaseTitle);
                btn.ContextMenu.Items.Add(mi4);
                var mi5 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 4, "tOGGLE cASE") };
                mi5.Click += (s, e) => oPenEfficiency.Features.Text.ChangeCaseFeature.ExecuteSpecific(GetManager(), PowerPoint.PpChangeCase.ppCaseToggle);
                btn.ContextMenu.Items.Add(mi5);
            }
            else if (btn.Name == "BtnFormatBold")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var apply = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Apply Bold Formatting") };
                apply.Click += (s, e) => oPenEfficiency.Features.Text.FormatBoldFeature.Apply(GetManager());
                btn.ContextMenu.Items.Add(apply);
                var remove = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Remove Bold Formatting") };
                remove.Click += (s, e) => oPenEfficiency.Features.Text.FormatBoldFeature.Remove(GetManager());
                btn.ContextMenu.Items.Add(remove);
                var toggle = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Toggle Bold Formatting") };
                toggle.Click += (s, e) => oPenEfficiency.Features.Text.FormatBoldFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(toggle);
            }
            else if (btn.Name == "BtnFormatItalic")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var apply = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Apply Italic Formatting") };
                apply.Click += (s, e) => oPenEfficiency.Features.Text.FormatItalicFeature.Apply(GetManager());
                btn.ContextMenu.Items.Add(apply);
                var remove = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Remove Italic Formatting") };
                remove.Click += (s, e) => oPenEfficiency.Features.Text.FormatItalicFeature.Remove(GetManager());
                btn.ContextMenu.Items.Add(remove);
                var toggle = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Toggle Italic Formatting") };
                toggle.Click += (s, e) => oPenEfficiency.Features.Text.FormatItalicFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(toggle);
            }
            else if (btn.Name == "BtnFormatUnderline")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var apply = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Apply Underline Formatting") };
                apply.Click += (s, e) => oPenEfficiency.Features.Text.FormatUnderlineFeature.Apply(GetManager());
                btn.ContextMenu.Items.Add(apply);
                var remove = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Remove Underline Formatting") };
                remove.Click += (s, e) => oPenEfficiency.Features.Text.FormatUnderlineFeature.Remove(GetManager());
                btn.ContextMenu.Items.Add(remove);
                var toggle = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Toggle Underline Formatting") };
                toggle.Click += (s, e) => oPenEfficiency.Features.Text.FormatUnderlineFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(toggle);
            }
            else if (btn.Name == "BtnFormatStrikethrough")
            {
                // REG-TEXTFONTPROPS-001: PowerPoint's Font object model has no strikethrough
                // property, so only the native toggle is offered (explicit apply/remove were dead).
                var toggle = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Toggle Strikethrough Formatting") };
                toggle.Click += (s, e) => oPenEfficiency.Features.Text.FormatStrikethroughFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(toggle);
            }

            else if (btn.Name == "BtnFormatSubscript")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var apply = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Apply Subscript Formatting") };
                apply.Click += (s, e) => oPenEfficiency.Features.Text.FormatSubscriptFeature.Apply(GetManager());
                btn.ContextMenu.Items.Add(apply);
                var remove = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Remove Subscript Formatting") };
                remove.Click += (s, e) => oPenEfficiency.Features.Text.FormatSubscriptFeature.Remove(GetManager());
                btn.ContextMenu.Items.Add(remove);
                var toggle = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Toggle Subscript Formatting") };
                toggle.Click += (s, e) => oPenEfficiency.Features.Text.FormatSubscriptFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(toggle);
            }
            else if (btn.Name == "BtnFormatSuperscript")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var apply = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Apply Superscript Formatting") };
                apply.Click += (s, e) => oPenEfficiency.Features.Text.FormatSuperscriptFeature.Apply(GetManager());
                btn.ContextMenu.Items.Add(apply);
                var remove = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Remove Superscript Formatting") };
                remove.Click += (s, e) => oPenEfficiency.Features.Text.FormatSuperscriptFeature.Remove(GetManager());
                btn.ContextMenu.Items.Add(remove);
                var toggle = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Toggle Superscript Formatting") };
                toggle.Click += (s, e) => oPenEfficiency.Features.Text.FormatSuperscriptFeature.Execute(GetManager());
                btn.ContextMenu.Items.Add(toggle);
            }
            else if (btn.Name == "BtnDistributeHorizontal")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Distribute Across Slide") };
                mi1.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeHorizontalFeature.Execute(GetManager(), relativeToSlide: true);
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Distribute Within Selection") };
                mi2.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeHorizontalFeature.Execute(GetManager(), relativeToSlide: false);
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Align Middles & Distribute Within Selection") };
                mi3.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeHorizontalFeature.Execute(GetManager(), relativeToSlide: false, alignMiddles: true);
                btn.ContextMenu.Items.Add(mi3);
                var mi4 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Align Middles & Distribute Across Slide") };
                mi4.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeHorizontalFeature.Execute(GetManager(), relativeToSlide: true, alignMiddles: true);
                btn.ContextMenu.Items.Add(mi4);
            }
            else if (btn.Name == "BtnDistributeVertical")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Distribute Across Slide") };
                mi1.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeVerticalFeature.Execute(GetManager(), relativeToSlide: true);
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Distribute Within Selection") };
                mi2.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeVerticalFeature.Execute(GetManager(), relativeToSlide: false);
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Align Centers & Distribute Within Selection") };
                mi3.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeVerticalFeature.Execute(GetManager(), relativeToSlide: false, alignCenters: true);
                btn.ContextMenu.Items.Add(mi3);
                var mi4 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Align Centers & Distribute Across Slide") };
                mi4.Click += (s, e) => oPenEfficiency.Features.Alignment.DistributeVerticalFeature.Execute(GetManager(), relativeToSlide: true, alignCenters: true);
                btn.ContextMenu.Items.Add(mi4);
            }
            else if (btn.Name == "BtnDiagonalAlign")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = "Align 45° (Anchor: Top/Left-most shape)" };
                mi1.Click += (s, e) => oPenEfficiency.Features.Alignment.DiagonalAlignFeature.ExecuteSpecific(GetManager(), 45f, "First");
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = "Align 135° (Anchor: Top/Left-most shape)" };
                mi2.Click += (s, e) => oPenEfficiency.Features.Alignment.DiagonalAlignFeature.ExecuteSpecific(GetManager(), 135f, "First");
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = "Align -45° (Anchor: Top/Left-most shape)" };
                mi3.Click += (s, e) => oPenEfficiency.Features.Alignment.DiagonalAlignFeature.ExecuteSpecific(GetManager(), -45f, "First");
                btn.ContextMenu.Items.Add(mi3);
                var mi4 = new MenuItem { Header = "Align -135° (Anchor: Top/Left-most shape)" };
                mi4.Click += (s, e) => oPenEfficiency.Features.Alignment.DiagonalAlignFeature.ExecuteSpecific(GetManager(), -135f, "First");
                btn.ContextMenu.Items.Add(mi4);
            }
            else if (btn.Name == "BtnAlignText")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var miL = new MenuItem { Header = "Left" };
                miL.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), null, PowerPoint.PpParagraphAlignment.ppAlignLeft);
                btn.ContextMenu.Items.Add(miL);
                var miC = new MenuItem { Header = "Center" };
                miC.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), null, PowerPoint.PpParagraphAlignment.ppAlignCenter);
                btn.ContextMenu.Items.Add(miC);
                var miR = new MenuItem { Header = "Right" };
                miR.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), null, PowerPoint.PpParagraphAlignment.ppAlignRight);
                btn.ContextMenu.Items.Add(miR);
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = "Top" };
                mi1.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorTop);
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = "Middle" };
                mi2.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle);
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = "Bottom" };
                mi3.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorBottom);
                btn.ContextMenu.Items.Add(mi3);
                btn.ContextMenu.Items.Add(new Separator());
                var mi4 = new MenuItem { Header = "Top Centered" };
                mi4.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorTop, PowerPoint.PpParagraphAlignment.ppAlignCenter);
                btn.ContextMenu.Items.Add(mi4);
                var mi5 = new MenuItem { Header = "Middle Centered" };
                mi5.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle, PowerPoint.PpParagraphAlignment.ppAlignCenter);
                btn.ContextMenu.Items.Add(mi5);
                var mi6 = new MenuItem { Header = "Bottom Centered" };
                mi6.Click += (s, e) => oPenEfficiency.Features.Text.AlignTextFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorBottom, PowerPoint.PpParagraphAlignment.ppAlignCenter);
                btn.ContextMenu.Items.Add(mi6);
            }
            else if (btn.Name == "BtnTextDirection")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Horizontal") };
                mi1.Click += (s, e) => oPenEfficiency.Features.Text.TextDirectionFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal);
                btn.ContextMenu.Items.Add(mi1);
                var mi2 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Rotate 90°") };
                mi2.Click += (s, e) => oPenEfficiency.Features.Text.TextDirectionFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationDownward);
                btn.ContextMenu.Items.Add(mi2);
                var mi3 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Rotate 270°") };
                mi3.Click += (s, e) => oPenEfficiency.Features.Text.TextDirectionFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationUpward);
                btn.ContextMenu.Items.Add(mi3);
                var mi4 = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Stacked") };
                mi4.Click += (s, e) => oPenEfficiency.Features.Text.TextDirectionFeature.ExecuteSpecific(GetManager(), Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationVerticalFarEast);
                btn.ContextMenu.Items.Add(mi4);
            }
            else if (btn.Name == "BtnCharacterSpacing")
            {
                // REG-TEXTFONTPROPS-001: PowerPoint's Font object model has no spacing property,
                // so only the native dialog entry is offered (exact presets were dead).
                var more = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "More Spacing...") };
                more.Click += (s, e) => oPenEfficiency.Features.Text.CharacterSpacingFeature.ExecuteMoreSpacing(GetManager());
                btn.ContextMenu.Items.Add(more);
            }

            else if (btn.Name == "BtnTranslate")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var miAPI = new MenuItem { Header = "Translate via Configured API" };
                miAPI.Click += (s, e) => {
                    var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
                    bool hasKey = false;
                    if (config != null)
                    {
                        if (config.SelectedTranslationTool == "DeepL" && !string.IsNullOrEmpty(config.DeepLApiKey)) hasKey = true;
                        else if (config.SelectedTranslationTool == "OpenAI" && !string.IsNullOrEmpty(config.OpenAIApiKey)) hasKey = true;
                        else if (config.SelectedTranslationTool == "Claude" && !string.IsNullOrEmpty(config.ClaudeApiKey)) hasKey = true;
                        else if (config.SelectedTranslationTool == "LibreTranslate" && !string.IsNullOrEmpty(config.LibreTranslateBaseUrl)) hasKey = true;
                    }
                    if (hasKey) oPenEfficiency.Features.Text.TranslationFeatures.TranslateSelectionSafe(GetManager(), config);
                    else System.Windows.MessageBox.Show("Please configure an API provider in the settings first.", "Translation", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                };
                btn.ContextMenu.Items.Add(miAPI);

                var miFree = new MenuItem { Header = "Translate via Free API (MyMemory)" };
                miFree.Click += (s, e) => {
                    var config = JsonService.Load<SidebarConfig>("sidebar_config.json");
                    string lang = config?.TranslationTargetLanguage ?? "EN-US";
                    oPenEfficiency.Features.Text.TranslationFeatures.TranslateWithFreeApiSafe(GetManager(), lang);
                };
                btn.ContextMenu.Items.Add(miFree);
            }
            else if (btn.Name == "BtnTransparency")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var levels = new[] { 0f, 0.25f, 0.5f, 0.75f, 1f };
                foreach (var l in levels) {
                    var mi = new MenuItem { Header = $"Set {l * 100}%" };
                    mi.Click += (s, e) => oPenEfficiency.Features.Visuals.TransparencyFeature.SetTransparency(GetManager(), l);
                    btn.ContextMenu.Items.Add(mi);
                }
            }
        }
    }
}
