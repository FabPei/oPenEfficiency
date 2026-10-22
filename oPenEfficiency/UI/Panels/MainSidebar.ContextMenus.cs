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
        private void AttachCatalogContextMenu(System.Windows.Controls.Primitives.ButtonBase btn)
        {
            AttachLayoutContextMenus(btn);
            AttachWorkflowAndVisualContextMenus(btn);
            AttachTableAndObjectContextMenus(btn);
            AttachTextFormattingContextMenus(btn);
        }

        private void PerformHarveyBallCycle()
        {
            var manager = GetManager();
            var selection = manager.GetSelectedShapes();
            float nextPercentage = 0.25f;
            if (selection != null && selection.Count == 1)
            {
                var selected = selection[1];
                if (HarveyBallFeature.IsHarveyBall(selected))
                {
                    float current = HarveyBallFeature.GetCurrentPercentage(selected);
                    nextPercentage = oPenEfficiency.Core.Services.HarveyBallStatePolicy.NextPercentage(current);
                }
            }
            HarveyBallFeature.Execute(manager, nextPercentage);
        }

        private void AddZOrderContextMenu(
            ButtonBase button,
            string rankedLabel,
            string directlyAboveLabel,
            string directlyBelowLabel,
            Func<PowerPointManager, bool> rankedAction)
        {
            button.ContextMenu.Items.Add(new Separator());

            var ranked = new MenuItem { Header = rankedLabel };
            ranked.Click += (s, e) => rankedAction(GetManager());
            button.ContextMenu.Items.Add(ranked);

            var directlyAbove = new MenuItem { Header = directlyAboveLabel };
            directlyAbove.Click += (s, e) =>
                ZOrderFeatureHelper.ExecuteRelativeToLastSelected(
                    GetManager(),
                    true,
                    button.Name + ".PlaceDirectlyAboveLastSelected");
            button.ContextMenu.Items.Add(directlyAbove);

            var directlyBelow = new MenuItem { Header = directlyBelowLabel };
            directlyBelow.Click += (s, e) =>
                ZOrderFeatureHelper.ExecuteRelativeToLastSelected(
                    GetManager(),
                    false,
                    button.Name + ".PlaceDirectlyBelowLastSelected");
            button.ContextMenu.Items.Add(directlyBelow);
        }

        private void AddTableFormatPainterPresetMenu(ButtonBase button, string header, IReadOnlyList<string> labels, int variantOffset, bool formatText, bool formatCell)
        {
            var submenu = new MenuItem { Header = header };
            var presets = new[]
            {
                (Fallback: "Horizontal", Up: false, Down: false, Left: true, Right: true),
                (Fallback: "Vertical", Up: true, Down: true, Left: false, Right: false),
                (Fallback: "All directions", Up: true, Down: true, Left: true, Right: true)
            };

            for (int index = 0; index < presets.Length; index++)
            {
                var preset = presets[index];
                var item = new MenuItem { Header = labels[variantOffset + index] };
                item.Click += (s, e) => TableFormatPainterFeature.Paint(
                    GetManager(), preset.Up, preset.Down, preset.Left, preset.Right, formatText, formatCell, 0);
                submenu.Items.Add(item);
            }

            button.ContextMenu.Items.Add(submenu);
        }

        private void AddSnapContextMenu(ButtonBase button, bool isGridButton, IReadOnlyList<string> labels)
        {
            if (button.ContextMenu.Items.Count > 0)
                button.ContextMenu.Items.Add(new Separator());

            var primary = new MenuItem { Header = labels[0], IsCheckable = true };
            primary.Click += (s, e) =>
            {
                if (isGridButton) SnapToGridFeature.SetEnabled(true);
                else SnapToObjectsFeature.SetEnabled(true);
                UpdateButtonStates();
            };
            button.ContextMenu.Items.Add(primary);

            var alternate = new MenuItem { Header = labels[1], IsCheckable = true };
            alternate.Click += (s, e) =>
            {
                if (isGridButton) SnapToObjectsFeature.SetEnabled(true);
                else SnapToGridFeature.SetEnabled(true);
                UpdateButtonStates();
            };
            button.ContextMenu.Items.Add(alternate);

            var disable = new MenuItem { Header = labels[2], IsCheckable = true };
            disable.Click += (s, e) =>
            {
                SnapToGridFeature.SetEnabled(false);
                SnapToObjectsFeature.SetEnabled(false);
                UpdateButtonStates();
            };
            button.ContextMenu.Items.Add(disable);

            button.ContextMenu.Opened += (s, e) =>
            {
                primary.IsChecked = isGridButton ? SnapToGridFeature.IsEnabled : SnapToObjectsFeature.IsEnabled;
                alternate.IsChecked = isGridButton ? SnapToObjectsFeature.IsEnabled : SnapToGridFeature.IsEnabled;
                disable.IsChecked = !SnapToGridFeature.IsEnabled && !SnapToObjectsFeature.IsEnabled;
            };
        }

        private void AddCheckboxPresetMenu(ButtonBase button, string header, IReadOnlyList<string> labels, int variantOffset, bool hasFrame)
        {
            var submenu = new MenuItem { Header = header };
            var states = new[]
            {
                (Fallback: "Tick", State: CheckboxFeature.CheckboxState.Tick),
                (Fallback: "Cross", State: CheckboxFeature.CheckboxState.Cross),
                (Fallback: "Empty", State: CheckboxFeature.CheckboxState.Empty)
            };

            for (int index = 0; index < states.Length; index++)
            {
                var preset = states[index];
                var item = new MenuItem { Header = labels[variantOffset + index] };
                item.Click += (s, e) => CheckboxFeature.Execute(GetManager(), preset.State, hasFrame, 32f, -1);
                submenu.Items.Add(item);
            }

            button.ContextMenu.Items.Add(submenu);
        }

        private void InsertNumerationPreset(int number)
        {
            var settings = NumerationFeature.NumerationSettings.Default();
            settings.Number = number;
            NumerationFeature.Execute(GetManager(), null, settings);
        }

        private void InsertProgressSeriesPreset(int count, int activeIndex, string mode, string layout)
        {
            var settings = ProgressSeriesFeature.SeriesSettings.Default();
            settings.Count = count;
            settings.ActiveIndex = activeIndex;
            settings.Mode = mode;
            settings.Layout = layout;
            ProgressSeriesFeature.Execute(GetManager(), settings);
        }

        private void AddStretchContextMenu(ButtonBase button, IReadOnlyList<string> labels)
        {
            if (button.ContextMenu.Items.Count > 0)
                button.ContextMenu.Items.Add(new Separator());

            var gap = new MenuItem { Header = "Stretch Across Gap" };
            var gapActions = new Action[]
            {
                () => StretchToFirstLeftFeature.Execute(GetManager()),
                () => StretchToFirstRightFeature.Execute(GetManager()),
                () => StretchToFirstTopFeature.Execute(GetManager()),
                () => StretchToFirstBottomFeature.Execute(GetManager())
            };
            for (int index = 0; index < gapActions.Length; index++)
            {
                var action = gapActions[index];
                var item = new MenuItem { Header = labels[index] };
                item.Click += (s, e) => action();
                gap.Items.Add(item);
            }
            button.ContextMenu.Items.Add(gap);

            var edges = new MenuItem { Header = "Move Edge to Anchor" };
            var edgeActions = new Action[]
            {
                () => StretchToFirstLeftEdgeFeature.Execute(GetManager()),
                () => StretchToFirstRightEdgeFeature.Execute(GetManager()),
                () => StretchToFirstTopEdgeFeature.Execute(GetManager()),
                () => StretchToFirstBottomEdgeFeature.Execute(GetManager())
            };
            for (int index = 0; index < edgeActions.Length; index++)
            {
                var action = edgeActions[index];
                var item = new MenuItem { Header = labels[index + 4] };
                item.Click += (s, e) => action();
                edges.Items.Add(item);
            }
            button.ContextMenu.Items.Add(edges);
        }

        private void AddTableMoveContextMenu(ButtonBase button, string rowsLabel, string columnsLabel, string positionLabel, string swapLabel, string duplicateLabel, string repeatLabel)
        {
            button.ContextMenu.Items.Add(new Separator());

            var rows = new MenuItem { Header = rowsLabel };
            AddTableMenuAction(rows, "Move up", () => MoveTableRowColumnFeature.Execute(GetManager(), MoveTableRowColumnFeature.MoveDirection.RowUp, MoveTableRowColumnFeature.CurrentMode));
            AddTableMenuAction(rows, "Move down", () => MoveTableRowColumnFeature.Execute(GetManager(), MoveTableRowColumnFeature.MoveDirection.RowDown, MoveTableRowColumnFeature.CurrentMode));
            AddTableMenuAction(rows, "Move to first available row", () => MoveTableRowColumnFeature.MoveToBoundary(GetManager(), MoveTableRowColumnFeature.MoveAxis.Row, true, MoveTableRowColumnFeature.CurrentMode));
            AddTableMenuAction(rows, "Move to last available row", () => MoveTableRowColumnFeature.MoveToBoundary(GetManager(), MoveTableRowColumnFeature.MoveAxis.Row, false, MoveTableRowColumnFeature.CurrentMode));
            AddTableMenuAction(rows, positionLabel, () => MoveTableRowColumnFeature.MoveToPositionInteractive(GetManager(), MoveTableRowColumnFeature.MoveAxis.Row, MoveTableRowColumnFeature.CurrentMode));
            rows.Items.Add(new Separator());
            AddTableMenuAction(rows, swapLabel + " rows", () => MoveTableRowColumnFeature.SwapSelected(GetManager(), MoveTableRowColumnFeature.MoveAxis.Row, MoveTableRowColumnFeature.CurrentMode));
            AddTableMenuAction(rows, duplicateLabel + " rows above", () => MoveTableRowColumnFeature.DuplicateSelected(GetManager(), MoveTableRowColumnFeature.MoveAxis.Row, true));
            AddTableMenuAction(rows, duplicateLabel + " rows below", () => MoveTableRowColumnFeature.DuplicateSelected(GetManager(), MoveTableRowColumnFeature.MoveAxis.Row, false));
            button.ContextMenu.Items.Add(rows);

            var columns = new MenuItem { Header = columnsLabel };
            AddTableMenuAction(columns, "Move left", () => MoveTableRowColumnFeature.Execute(GetManager(), MoveTableRowColumnFeature.MoveDirection.ColumnLeft, MoveTableRowColumnFeature.CurrentMode));
            AddTableMenuAction(columns, "Move right", () => MoveTableRowColumnFeature.Execute(GetManager(), MoveTableRowColumnFeature.MoveDirection.ColumnRight, MoveTableRowColumnFeature.CurrentMode));
            AddTableMenuAction(columns, "Move to first column", () => MoveTableRowColumnFeature.MoveToBoundary(GetManager(), MoveTableRowColumnFeature.MoveAxis.Column, true, MoveTableRowColumnFeature.CurrentMode));
            AddTableMenuAction(columns, "Move to last column", () => MoveTableRowColumnFeature.MoveToBoundary(GetManager(), MoveTableRowColumnFeature.MoveAxis.Column, false, MoveTableRowColumnFeature.CurrentMode));
            AddTableMenuAction(columns, positionLabel, () => MoveTableRowColumnFeature.MoveToPositionInteractive(GetManager(), MoveTableRowColumnFeature.MoveAxis.Column, MoveTableRowColumnFeature.CurrentMode));
            columns.Items.Add(new Separator());
            AddTableMenuAction(columns, swapLabel + " columns", () => MoveTableRowColumnFeature.SwapSelected(GetManager(), MoveTableRowColumnFeature.MoveAxis.Column, MoveTableRowColumnFeature.CurrentMode));
            AddTableMenuAction(columns, duplicateLabel + " columns left", () => MoveTableRowColumnFeature.DuplicateSelected(GetManager(), MoveTableRowColumnFeature.MoveAxis.Column, true));
            AddTableMenuAction(columns, duplicateLabel + " columns right", () => MoveTableRowColumnFeature.DuplicateSelected(GetManager(), MoveTableRowColumnFeature.MoveAxis.Column, false));
            button.ContextMenu.Items.Add(columns);

            var modes = new MenuItem { Header = "Move mode" };
            var modeItems = new Dictionary<MoveTableRowColumnFeature.MoveMode, MenuItem>();
            AddTableMoveModeItem(modes, modeItems, "Keep table design", MoveTableRowColumnFeature.MoveMode.KeepTableDesign);
            AddTableMoveModeItem(modes, modeItems, "Content and all formatting", MoveTableRowColumnFeature.MoveMode.Complete);
            AddTableMoveModeItem(modes, modeItems, "Content and formatting, keep borders", MoveTableRowColumnFeature.MoveMode.IgnoreBorders);
            AddTableMoveModeItem(modes, modeItems, "Plain text only", MoveTableRowColumnFeature.MoveMode.TextOnly);
            button.ContextMenu.Items.Add(modes);

            var protection = new MenuItem { Header = "Protected rows" };
            var protectFirst = new MenuItem { Header = "Protect first row (header)", IsCheckable = true };
            protectFirst.Click += (s, e) => MoveTableRowColumnFeature.ProtectFirstRow = protectFirst.IsChecked;
            protection.Items.Add(protectFirst);
            var protectLast = new MenuItem { Header = "Protect last row (total)", IsCheckable = true };
            protectLast.Click += (s, e) => MoveTableRowColumnFeature.ProtectLastRow = protectLast.IsChecked;
            protection.Items.Add(protectLast);
            button.ContextMenu.Items.Add(protection);

            button.ContextMenu.Items.Add(new Separator());
            var repeat = new MenuItem { Header = repeatLabel };
            repeat.Click += (s, e) => MoveTableRowColumnFeature.RepeatLastMove(GetManager());
            button.ContextMenu.Items.Add(repeat);

            button.ContextMenu.Opened += (s, e) =>
            {
                foreach (var pair in modeItems) pair.Value.IsChecked = pair.Key == MoveTableRowColumnFeature.CurrentMode;
                protectFirst.IsChecked = MoveTableRowColumnFeature.ProtectFirstRow;
                protectLast.IsChecked = MoveTableRowColumnFeature.ProtectLastRow;
                repeat.IsEnabled = MoveTableRowColumnFeature.CanRepeatLastMove;
            };
        }

        private static void AddTableMenuAction(MenuItem parent, string header, Action action)
        {
            var item = new MenuItem { Header = header };
            item.Click += (s, e) => action();
            parent.Items.Add(item);
        }

        private static void AddTableMoveModeItem(MenuItem parent, IDictionary<MoveTableRowColumnFeature.MoveMode, MenuItem> items, string header, MoveTableRowColumnFeature.MoveMode mode)
        {
            var item = new MenuItem { Header = header, IsCheckable = true, StaysOpenOnClick = true };
            item.Click += (s, e) =>
            {
                MoveTableRowColumnFeature.CurrentMode = mode;
                foreach (var pair in items) pair.Value.IsChecked = pair.Key == mode;
            };
            items[mode] = item;
            parent.Items.Add(item);
        }

    }
}
