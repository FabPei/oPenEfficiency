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
        private void AttachTableAndObjectContextMenus(ButtonBase btn)
        {
            if (btn.Name == "BtnTableSort")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var mi1 = new MenuItem { Header = "Sort Selected Cells A to Z" };
                mi1.Click += (s, e) => TableSortFeature.Execute(GetManager(), TableSortFeature.SortType.AZ_Selection);
                btn.ContextMenu.Items.Add(mi1);

                var mi2 = new MenuItem { Header = "Sort Selected Cells Z to A" };
                mi2.Click += (s, e) => TableSortFeature.Execute(GetManager(), TableSortFeature.SortType.ZA_Selection);
                btn.ContextMenu.Items.Add(mi2);

                var mi3 = new MenuItem { Header = "Sort Whole Row Text A to Z" };
                mi3.Click += (s, e) => TableSortFeature.Execute(GetManager(), TableSortFeature.SortType.AZ_WholeRow);
                btn.ContextMenu.Items.Add(mi3);

                var mi4 = new MenuItem { Header = "Sort Whole Row Text Z to A" };
                mi4.Click += (s, e) => TableSortFeature.Execute(GetManager(), TableSortFeature.SortType.ZA_WholeRow);
                btn.ContextMenu.Items.Add(mi4);
            }
            else if (btn.Name == "BtnMoveTableRowColumn")
            {
                AddTableMoveContextMenu(
                    btn,
                    GetCatalogVariantLabel(btn.Name, 0, "Move Rows"),
                    GetCatalogVariantLabel(btn.Name, 1, "Move Columns"),
                    GetCatalogVariantLabel(btn.Name, 2, "Move to Position..."),
                    GetCatalogVariantLabel(btn.Name, 3, "Swap Selected"),
                    GetCatalogVariantLabel(btn.Name, 4, "Duplicate Selected"),
                    GetCatalogVariantLabel(btn.Name, 5, "Repeat Last Table Move"));
            }
            else if (btn.Name == "BtnUpdateExcelCharts")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var miUpdate = new MenuItem { Header = "Refresh Linked Excel Charts on This Slide" };
                miUpdate.Click += (s, e) => UpdateExcelChartsFeature.Execute(GetManager(), updateAllSlides: false);
                btn.ContextMenu.Items.Add(miUpdate);

                btn.ContextMenu.Items.Add(new Separator());
                var miAll = new MenuItem { Header = "Update All Charts in Presentation" };
                miAll.Click += (s, e) => UpdateExcelChartsFeature.Execute(GetManager(), updateAllSlides: true);
                btn.ContextMenu.Items.Add(miAll);
            }
            else if (btn.Name == "BtnFlightMode")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var miToggle = new MenuItem { Header = "Toggle Master Covers & Overlays" };
                miToggle.Click += (s, e) => ExecuteRegisteredCommand(
                    FlightModeCommand,
                    FlightModeFeature.FlightModeType.CoverLogos,
                    CommandInvocationSource.ContextMenu,
                    btn,
                    ToggleIntent.Toggle);
                btn.ContextMenu.Items.Add(miToggle);

                btn.ContextMenu.Items.Add(new Separator());

                var miCover = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Cover Logos") };
                miCover.Click += (s, e) => ExecuteRegisteredCommand(
                    FlightModeCommand,
                    FlightModeFeature.FlightModeType.CoverLogos,
                    CommandInvocationSource.ContextMenu,
                    btn,
                    ToggleIntent.EnsureOn);
                btn.ContextMenu.Items.Add(miCover);

                var miLight = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Light Overlay") };
                miLight.Click += (s, e) => ExecuteRegisteredCommand(
                    FlightModeCommand,
                    FlightModeFeature.FlightModeType.LightOverlay,
                    CommandInvocationSource.ContextMenu,
                    btn,
                    ToggleIntent.EnsureOn);
                btn.ContextMenu.Items.Add(miLight);

                var miHeavy = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Heavy Overlay") };
                miHeavy.Click += (s, e) => ExecuteRegisteredCommand(
                    FlightModeCommand,
                    FlightModeFeature.FlightModeType.HeavyOverlay,
                    CommandInvocationSource.ContextMenu,
                    btn,
                    ToggleIntent.EnsureOn);
                btn.ContextMenu.Items.Add(miHeavy);

                btn.ContextMenu.Items.Add(new Separator());

                var miDisable = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Disable Master Covers & Overlays") };
                miDisable.Click += (s, e) => ExecuteRegisteredCommand(
                    FlightModeCommand,
                    FlightModeFeature.FlightModeType.None,
                    CommandInvocationSource.ContextMenu,
                    btn,
                    ToggleIntent.EnsureOff);
                btn.ContextMenu.Items.Add(miDisable);
            }
            else if (btn.Name == "BtnInsertColumnLeft")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var preserveWidth = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Insert Column Left & Preserve Width") };
                preserveWidth.Click += (s, e) => InsertTableColumnLeftFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(preserveWidth);
                var duplicateFormat = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Duplicate Selected Column Left – Format Only") };
                duplicateFormat.Click += (s, e) => InsertTableColumnLeftFeature.DuplicateFormatOnly(GetManager());
                btn.ContextMenu.Items.Add(duplicateFormat);
                var insertMultiple = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Insert Multiple Columns Left...") };
                insertMultiple.Click += (s, e) =>
                {
                    var dialog = new UI.InsertTableColumnsDialog();
                    var manager = GetManager();
                    var app = manager?.GetApplication();
                    if (app != null) new System.Windows.Interop.WindowInteropHelper(dialog).Owner = (IntPtr)app.HWND;
                    if (dialog.ShowDialog() == true)
                        InsertTableColumnLeftFeature.ExecuteMultiple(manager, dialog.ItemCount, dialog.PreserveDimension);
                };
                btn.ContextMenu.Items.Add(insertMultiple);
            }
            else if (btn.Name == "BtnInsertColumnRight")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var preserveWidth = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Insert Column Right & Preserve Width") };
                preserveWidth.Click += (s, e) => InsertTableColumnRightFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(preserveWidth);

                var duplicateFormat = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Duplicate Selected Column Right – Format Only") };
                duplicateFormat.Click += (s, e) => InsertTableColumnRightFeature.DuplicateFormatOnly(GetManager());
                btn.ContextMenu.Items.Add(duplicateFormat);

                var insertMultiple = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Insert Multiple Columns...") };
                insertMultiple.Click += (s, e) =>
                {
                    var dialog = new UI.InsertTableColumnsDialog();
                    var manager = GetManager();
                    var app = manager?.GetApplication();
                    if (app != null)
                        new System.Windows.Interop.WindowInteropHelper(dialog).Owner = (IntPtr)app.HWND;
                    if (dialog.ShowDialog() == true)
                        InsertTableColumnRightFeature.ExecuteMultiple(manager, dialog.ItemCount, dialog.PreserveDimension);
                };
                btn.ContextMenu.Items.Add(insertMultiple);
            }
            else if (btn.Name == "BtnInsertRowTop")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var preserveHeight = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Insert Row Above & Preserve Height") };
                preserveHeight.Click += (s, e) => InsertTableRowTopFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(preserveHeight);
                var duplicateFormat = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Duplicate Selected Row Above – Format Only") };
                duplicateFormat.Click += (s, e) => InsertTableRowTopFeature.DuplicateFormatOnly(GetManager());
                btn.ContextMenu.Items.Add(duplicateFormat);
                var insertMultiple = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Insert Multiple Rows Above...") };
                insertMultiple.Click += (s, e) =>
                {
                    var dialog = new UI.InsertTableColumnsDialog(rows: true);
                    var manager = GetManager();
                    var app = manager?.GetApplication();
                    if (app != null) new System.Windows.Interop.WindowInteropHelper(dialog).Owner = (IntPtr)app.HWND;
                    if (dialog.ShowDialog() == true)
                        InsertTableRowTopFeature.ExecuteMultiple(manager, dialog.ItemCount, dialog.PreserveDimension);
                };
                btn.ContextMenu.Items.Add(insertMultiple);
            }
            else if (btn.Name == "BtnInsertRowBottom")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var preserveHeight = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Insert Row Below & Preserve Height") };
                preserveHeight.Click += (s, e) => InsertTableRowBottomFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(preserveHeight);
                var duplicateFormat = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Duplicate Selected Row Below – Format Only") };
                duplicateFormat.Click += (s, e) => InsertTableRowBottomFeature.DuplicateFormatOnly(GetManager());
                btn.ContextMenu.Items.Add(duplicateFormat);
                var insertMultiple = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Insert Multiple Rows Below...") };
                insertMultiple.Click += (s, e) =>
                {
                    var dialog = new UI.InsertTableColumnsDialog(rows: true);
                    var manager = GetManager();
                    var app = manager?.GetApplication();
                    if (app != null) new System.Windows.Interop.WindowInteropHelper(dialog).Owner = (IntPtr)app.HWND;
                    if (dialog.ShowDialog() == true)
                        InsertTableRowBottomFeature.ExecuteMultiple(manager, dialog.ItemCount, dialog.PreserveDimension);
                };
                btn.ContextMenu.Items.Add(insertMultiple);
            }
            else if (btn.Name == "BtnAlignToTableCell")
            {
                if (btn.ContextMenu.Items.Count > 0)
                    btn.ContextMenu.Items.Add(new Separator());

                var labels = new[]
                {
                    GetCatalogVariantLabel(btn.Name, 0, "Top Left"),
                    GetCatalogVariantLabel(btn.Name, 1, "Top Center"),
                    GetCatalogVariantLabel(btn.Name, 2, "Top Right"),
                    GetCatalogVariantLabel(btn.Name, 3, "Middle Left"),
                    GetCatalogVariantLabel(btn.Name, 4, "Center"),
                    GetCatalogVariantLabel(btn.Name, 5, "Middle Right"),
                    GetCatalogVariantLabel(btn.Name, 6, "Bottom Left"),
                    GetCatalogVariantLabel(btn.Name, 7, "Bottom Center"),
                    GetCatalogVariantLabel(btn.Name, 8, "Bottom Right")
                };
                var rows = new[]
                {
                    (Header: "Top", Values: new[] { (labels[0], "TOP_LEFT"), (labels[1], "TOP_CENTER"), (labels[2], "TOP_RIGHT") }),
                    (Header: "Middle", Values: new[] { (labels[3], "MIDDLE_LEFT"), (labels[4], "CENTER"), (labels[5], "MIDDLE_RIGHT") }),
                    (Header: "Bottom", Values: new[] { (labels[6], "BOTTOM_LEFT"), (labels[7], "BOTTOM_CENTER"), (labels[8], "BOTTOM_RIGHT") })
                };

                foreach (var row in rows)
                {
                    var submenu = new MenuItem { Header = row.Header };
                    foreach (var value in row.Values)
                    {
                        var alignment = value.Item2;
                        var item = new MenuItem { Header = value.Item1 };
                        item.Click += (s, e) => AlignToTableCellFeature.Align(GetManager(), alignment);
                        submenu.Items.Add(item);
                    }
                    btn.ContextMenu.Items.Add(submenu);
                }
            }
            else if (btn.Name == "BtnTableFormatPainter")
            {
                if (btn.ContextMenu.Items.Count > 0)
                    btn.ContextMenu.Items.Add(new Separator());

                var labels = new[]
                {
                    GetCatalogVariantLabel(btn.Name, 0, "All Formats Horizontally"),
                    GetCatalogVariantLabel(btn.Name, 1, "All Formats Vertically"),
                    GetCatalogVariantLabel(btn.Name, 2, "All Formats in All Directions"),
                    GetCatalogVariantLabel(btn.Name, 3, "Text Formatting Horizontally"),
                    GetCatalogVariantLabel(btn.Name, 4, "Text Formatting Vertically"),
                    GetCatalogVariantLabel(btn.Name, 5, "Text Formatting in All Directions"),
                    GetCatalogVariantLabel(btn.Name, 6, "Cell Fill Horizontally"),
                    GetCatalogVariantLabel(btn.Name, 7, "Cell Fill Vertically"),
                    GetCatalogVariantLabel(btn.Name, 8, "Cell Fill in All Directions")
                };
                AddTableFormatPainterPresetMenu(btn, "All formats", labels, 0, true, true);
                AddTableFormatPainterPresetMenu(btn, "Text formatting", labels, 3, true, false);
                AddTableFormatPainterPresetMenu(btn, "Cell fill", labels, 6, false, true);
            }
            else if (btn.Name == "BtnTableColumnWidth")
            {
                var copy = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Copy Column Widths") };
                copy.Click += (s, e) => TableColumnWidthFeature.Execute(GetManager(), false);
                btn.ContextMenu.Items.Add(copy);
                var paste = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Paste Column Widths") };
                paste.Click += (s, e) => TableColumnWidthFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(paste);
                var equalize = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Make Column Widths Equal") };
                equalize.Click += (s, e) => TableColumnWidthFeature.Equalize(GetManager());
                btn.ContextMenu.Items.Add(equalize);
                btn.ContextMenu.Items.Add(new Separator());
                var clear = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Clear Copied Column Widths") };
                clear.Click += (s, e) => TableColumnWidthFeature.ClearCopiedWidths();
                btn.ContextMenu.Items.Add(clear);
            }
            else if (btn.Name == "BtnTableRowHeight")
            {
                var copy = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Copy Row Heights") };
                copy.Click += (s, e) => TableRowHeightFeature.Execute(GetManager(), false);
                btn.ContextMenu.Items.Add(copy);
                var paste = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Paste Row Heights") };
                paste.Click += (s, e) => TableRowHeightFeature.Execute(GetManager(), true);
                btn.ContextMenu.Items.Add(paste);
                var equalize = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Make Row Heights Equal") };
                equalize.Click += (s, e) => TableRowHeightFeature.Equalize(GetManager());
                btn.ContextMenu.Items.Add(equalize);
                btn.ContextMenu.Items.Add(new Separator());
                var clear = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Clear Copied Row Heights") };
                clear.Click += (s, e) => TableRowHeightFeature.ClearCopiedHeights();
                btn.ContextMenu.Items.Add(clear);
            }
            else if (btn.Name == "BtnRemoveHorizontalGaps")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Remove Horizontal Gaps & Center" };
                mi.Click += (s, e) => RemoveHorizontalGapsFeature.ExecuteCentered(GetManager());
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnRemoveVerticalGaps")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Remove Vertical Gaps & Center" };
                mi.Click += (s, e) => RemoveVerticalGapsFeature.ExecuteCentered(GetManager());
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnTableSplit")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi = new MenuItem { Header = "Split Table by Row" };
                mi.Click += (s, e) => TableSplitFeature.Execute(GetManager(), false);
                btn.ContextMenu.Items.Add(mi);
            }
            else if (btn.Name == "BtnTableSum")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var mi1 = new MenuItem { Header = "Calculate Column Totals" };
                mi1.Click += (s, e) => TableSumFeature.Execute(GetManager(), false, true);
                btn.ContextMenu.Items.Add(mi1);

                var mi2 = new MenuItem { Header = "Calculate Row & Column Totals" };
                mi2.Click += (s, e) => TableSumFeature.Execute(GetManager(), true, true);
                btn.ContextMenu.Items.Add(mi2);
            }
            else if (btn.Name == "BtnAdjustSpacing")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var miIncH = new MenuItem { Header = "Increase Horizontal Spacing" };
                miIncH.Click += (s, e) => AdjustHorizontalSpacingFeature.Execute(GetManager(), GetSpacingIncrement());
                btn.ContextMenu.Items.Add(miIncH);

                var miDecH = new MenuItem { Header = "Decrease Horizontal Spacing" };
                miDecH.Click += (s, e) => AdjustHorizontalSpacingFeature.Execute(GetManager(), -GetSpacingIncrement());
                btn.ContextMenu.Items.Add(miDecH);

                var miDecHCentered = new MenuItem { Header = "Decrease Horizontal Spacing from Center" };
                miDecHCentered.Click += (s, e) => AdjustHorizontalSpacingFeature.ExecuteFromCenter(GetManager(), -GetSpacingIncrement());
                btn.ContextMenu.Items.Add(miDecHCentered);

                var miIncV = new MenuItem { Header = "Increase Vertical Spacing" };
                miIncV.Click += (s, e) => AdjustVerticalSpacingFeature.Execute(GetManager(), GetSpacingIncrement());
                btn.ContextMenu.Items.Add(miIncV);

                var miDecV = new MenuItem { Header = "Decrease Vertical Spacing" };
                miDecV.Click += (s, e) => AdjustVerticalSpacingFeature.Execute(GetManager(), -GetSpacingIncrement());
                btn.ContextMenu.Items.Add(miDecV);

                var miDecVCentered = new MenuItem { Header = "Decrease Vertical Spacing from Center" };
                miDecVCentered.Click += (s, e) => AdjustVerticalSpacingFeature.ExecuteFromCenter(GetManager(), -GetSpacingIncrement());
                btn.ContextMenu.Items.Add(miDecVCentered);
            }
            else if (btn.Name == "BtnSnapShape")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var left = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Snap Left") };
                left.Click += (s, e) => SnapShapeFeature.Execute(GetManager(), SnapShapeFeature.SnapDirection.Left);
                btn.ContextMenu.Items.Add(left);
                var right = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Snap Right") };
                right.Click += (s, e) => SnapShapeFeature.Execute(GetManager(), SnapShapeFeature.SnapDirection.Right);
                btn.ContextMenu.Items.Add(right);
                var above = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Snap Above") };
                above.Click += (s, e) => SnapShapeFeature.Execute(GetManager(), SnapShapeFeature.SnapDirection.Above);
                btn.ContextMenu.Items.Add(above);
                var below = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 3, "Snap Below") };
                below.Click += (s, e) => SnapShapeFeature.Execute(GetManager(), SnapShapeFeature.SnapDirection.Below);
                btn.ContextMenu.Items.Add(below);
            }
            else if (btn.Name == "BtnSnapToGrid")
            {
                var labels = new[]
                {
                    GetCatalogVariantLabel(btn.Name, 0, "Enable Snap to Grid"),
                    GetCatalogVariantLabel(btn.Name, 1, "Enable Snap to Objects"),
                    GetCatalogVariantLabel(btn.Name, 2, "Disable Snapping")
                };
                AddSnapContextMenu(btn, true, labels);
            }
            else if (btn.Name == "BtnSnapToObjects")
            {
                var labels = new[]
                {
                    GetCatalogVariantLabel(btn.Name, 0, "Enable Snap to Objects"),
                    GetCatalogVariantLabel(btn.Name, 1, "Enable Snap to Grid"),
                    GetCatalogVariantLabel(btn.Name, 2, "Disable Snapping")
                };
                AddSnapContextMenu(btn, false, labels);
            }
            else if (btn.Name == "BtnPickColor")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var miFill = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Pick Fill Color") };
                miFill.Click += (s, e) => Features.PickColorFeature.Execute(GetManager(), Features.PickColorFeature.ApplyTarget.Fill);
                btn.ContextMenu.Items.Add(miFill);

                var miLine = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Pick Line Color") };
                miLine.Click += (s, e) => Features.PickColorFeature.Execute(GetManager(), Features.PickColorFeature.ApplyTarget.Line);
                btn.ContextMenu.Items.Add(miLine);

                var miText = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Pick Text Color") };
                miText.Click += (s, e) => Features.PickColorFeature.Execute(GetManager(), Features.PickColorFeature.ApplyTarget.Text);
                btn.ContextMenu.Items.Add(miText);
            }
            else if (btn.Name == "BtnPickFillColor")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var fill = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Pick Fill Color") };
                fill.Click += (s, e) => PickColorFeature.Execute(GetManager(), PickColorFeature.ApplyTarget.Fill);
                btn.ContextMenu.Items.Add(fill);
                var line = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Pick Line Color") };
                line.Click += (s, e) => PickColorFeature.Execute(GetManager(), PickColorFeature.ApplyTarget.Line);
                btn.ContextMenu.Items.Add(line);
                var text = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Pick Text Color") };
                text.Click += (s, e) => PickColorFeature.Execute(GetManager(), PickColorFeature.ApplyTarget.Text);
                btn.ContextMenu.Items.Add(text);
            }
            else if (btn.Name == "BtnPickLineColor")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var fill = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Pick Fill Color") };
                fill.Click += (s, e) => PickColorFeature.Execute(GetManager(), PickColorFeature.ApplyTarget.Fill);
                btn.ContextMenu.Items.Add(fill);
                var line = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Pick Line Color") };
                line.Click += (s, e) => PickColorFeature.Execute(GetManager(), PickColorFeature.ApplyTarget.Line);
                btn.ContextMenu.Items.Add(line);
                var text = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Pick Text Color") };
                text.Click += (s, e) => PickColorFeature.Execute(GetManager(), PickColorFeature.ApplyTarget.Text);
                btn.ContextMenu.Items.Add(text);
            }
            else if (btn.Name == "BtnPickTextColor")
            {
                btn.ContextMenu.Items.Add(new Separator());
                var fill = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 0, "Pick Fill Color") };
                fill.Click += (s, e) => PickColorFeature.Execute(GetManager(), PickColorFeature.ApplyTarget.Fill);
                btn.ContextMenu.Items.Add(fill);
                var line = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 1, "Pick Line Color") };
                line.Click += (s, e) => PickColorFeature.Execute(GetManager(), PickColorFeature.ApplyTarget.Line);
                btn.ContextMenu.Items.Add(line);
                var text = new MenuItem { Header = GetCatalogVariantLabel(btn.Name, 2, "Pick Text Color") };
                text.Click += (s, e) => PickColorFeature.Execute(GetManager(), PickColorFeature.ApplyTarget.Text);
                btn.ContextMenu.Items.Add(text);
            }
            else if (btn.Name == "BtnRepeatShape")
            {
                btn.ContextMenu.Items.Add(new Separator());

                var miRight = new MenuItem { Header = "Repeat to the Right" };
                miRight.Click += (s, e) => ExecuteRegisteredCommand(
                    RepeatShapeCommand,
                    RepeatShapeFeature.RepeatDirection.Right,
                    CommandInvocationSource.ContextMenu,
                    btn);
                btn.ContextMenu.Items.Add(miRight);

                var miLeft = new MenuItem { Header = "Repeat to the Left" };
                miLeft.Click += (s, e) => ExecuteRegisteredCommand(
                    RepeatShapeCommand,
                    RepeatShapeFeature.RepeatDirection.Left,
                    CommandInvocationSource.ContextMenu,
                    btn);
                btn.ContextMenu.Items.Add(miLeft);

                var miUp = new MenuItem { Header = "Repeat to the Top" };
                miUp.Click += (s, e) => ExecuteRegisteredCommand(
                    RepeatShapeCommand,
                    RepeatShapeFeature.RepeatDirection.Up,
                    CommandInvocationSource.ContextMenu,
                    btn);
                btn.ContextMenu.Items.Add(miUp);

                var miDown = new MenuItem { Header = "Repeat to the Bottom" };
                miDown.Click += (s, e) => ExecuteRegisteredCommand(
                    RepeatShapeCommand,
                    RepeatShapeFeature.RepeatDirection.Down,
                    CommandInvocationSource.ContextMenu,
                    btn);
                btn.ContextMenu.Items.Add(miDown);

            }
        }
    }
}
