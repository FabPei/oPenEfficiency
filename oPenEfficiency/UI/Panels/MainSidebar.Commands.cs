using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls.Primitives;
using oPenEfficiency.Features;
using oPenEfficiency.Features.Utilities;
using oPenEfficiency.Features.Text;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Commands;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using oPenEfficiency.UI;

namespace oPenEfficiency
{
    public partial class MainSidebar
    {
        private static readonly CommandKey<RepeatShapeFeature.RepeatDirection> RepeatShapeCommand =
            new CommandKey<RepeatShapeFeature.RepeatDirection>("BtnRepeatShape");
        private static readonly CommandKey<FlightModeFeature.FlightModeType> FlightModeCommand =
            new CommandKey<FlightModeFeature.FlightModeType>("BtnFlightMode");

        private static readonly Lazy<CommandRegistry> Registry =
            new Lazy<CommandRegistry>(BuildCommandRegistry, true);

        private static CommandRegistry BuildCommandRegistry()
        {
            try
            {
                var registry = new CommandRegistry();
                RegisterGeneratedFeatureCommands(registry);
                registry.RegisterAction(
                    RepeatShapeCommand,
                    RepeatShapeFeature.RepeatDirection.Right,
                    (context, direction) => FromFeatureResult(
                        RepeatShapeFeature.Execute(context.PowerPoint, direction)));
                registry.RegisterToggle(
                    FlightModeCommand,
                    FlightModeFeature.FlightModeType.CoverLogos,
                    ExecuteFlightModeCommand);
                RegisterSpecializedCommands(registry);
                var validationErrors = FeatureCatalogValidator.ValidateCommands(
                    FeatureLibrary.Catalog,
                    registry.Commands.Select(command => new FeatureCommandContract
                    {
                        Id = command.Id,
                        IsToggle = command.Kind == CommandKind.Toggle,
                        HasValidPrimaryArguments = true
                    }));
                if (validationErrors.Count > 0)
                    throw new CommandRegistrationException(
                        "Command catalog validation failed:" + Environment.NewLine +
                        string.Join(Environment.NewLine, validationErrors));
                registry.Freeze();
                return registry;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MainSidebar.BuildCommandRegistry");
                throw;
            }
        }

        private static void RegisterAction(
            CommandRegistry registry,
            string id,
            Func<CommandContext, bool> execute)
        {
            registry.RegisterAction(
                new CommandKey<NoCommandArguments>(id),
                NoCommandArguments.Value,
                (context, arguments) => FromFeatureResult(execute(context)));
        }

        private static void RegisterVoidAction(
            CommandRegistry registry,
            string id,
            Action<CommandContext> execute)
        {
            registry.RegisterAction(
                new CommandKey<NoCommandArguments>(id),
                NoCommandArguments.Value,
                (context, arguments) =>
                {
                    execute(context);
                    return CommandDispatchResult.Succeeded();
                });
        }

        private static void RegisterToggle(
            CommandRegistry registry,
            string id,
            Func<CommandContext, ToggleIntent, CommandDispatchResult> execute)
        {
            registry.RegisterToggle(
                new CommandKey<NoCommandArguments>(id),
                NoCommandArguments.Value,
                (context, arguments, intent) => execute(context, intent));
        }

        private static void RegisterHostAction(
            CommandRegistry registry,
            string id,
            CommandHostAction action)
        {
            registry.RegisterAction(
                new CommandKey<NoCommandArguments>(id),
                NoCommandArguments.Value,
                (context, arguments) =>
                    context.Ui.ExecuteHostAction(action, ToggleIntent.Toggle));
        }

        private static void RegisterSpecializedCommands(CommandRegistry registry)
        {
            RegisterHostAction(registry, "BtnAdjustSpacing", CommandHostAction.AdjustSpacing);
            RegisterHostAction(registry, "BtnDecreaseHorizontalSpacing", CommandHostAction.DecreaseHorizontalSpacing);
            RegisterHostAction(registry, "BtnDecreaseVerticalSpacing", CommandHostAction.DecreaseVerticalSpacing);
            RegisterHostAction(registry, "BtnExploreFeatures", CommandHostAction.OpenFeatureExplorer);
            RegisterHostAction(registry, "BtnIncreaseHorizontalSpacing", CommandHostAction.IncreaseHorizontalSpacing);
            RegisterHostAction(registry, "BtnIncreaseVerticalSpacing", CommandHostAction.IncreaseVerticalSpacing);
            RegisterHostAction(registry, "BtnParagraphDialog", CommandHostAction.OpenParagraphDialog);
            RegisterHostAction(registry, "BtnSettings", CommandHostAction.OpenSettings);
            RegisterHostAction(registry, "BtnSpecialChars", CommandHostAction.OpenSpecialCharacters);
            RegisterHostAction(registry, "BtnTemplateManager", CommandHostAction.OpenTemplateManager);
            RegisterHostAction(registry, "BtnUpdateExcelCharts", CommandHostAction.UpdateExcelCharts);

            RegisterAction(registry, "BtnArrangeCircle", context =>
                ArrangeInShapeFeature.Execute(context.PowerPoint, "circle"));
            RegisterAction(registry, "BtnMoveTableRowColumn", context =>
                context.Ui.TryOpenContextMenu() || MoveTableRowColumnFeature.Execute(
                    context.PowerPoint,
                    MoveTableRowColumnFeature.MoveDirection.RowUp,
                    MoveTableRowColumnFeature.CurrentMode));
            RegisterAction(registry, "BtnTableSort", context =>
                context.Ui.TryOpenContextMenu() || TableSortFeature.Execute(context.PowerPoint));
            RegisterAction(registry, "BtnTranslate", context =>
                TranslationFeatures.Execute(context.PowerPoint));

            registry.RegisterAction(
                new CommandKey<NoCommandArguments>("BtnApplyTextTool"),
                NoCommandArguments.Value,
                (context, arguments) => SetModelessWindow(
                    context,
                    "BtnApplyTextTool",
                    ToggleIntent.Toggle,
                    () => new UI.ApplyTextWindow(context.PowerPoint)));
            RegisterToggle(registry, "BtnAssetManager", (context, intent) =>
                context.Ui.ExecuteHostAction(CommandHostAction.ToggleAssetManager, intent));
            RegisterToggle(registry, "BtnGanttChart", (context, intent) =>
                context.Ui.ExecuteHostAction(CommandHostAction.ShowGantt, intent));
            RegisterToggle(registry, "BtnMagicResizer", (context, intent) =>
                SetModelessWindow(context, "BtnMagicResizer", intent,
                    () => new UI.MagicResizerWindow(context.PowerPoint)));
            RegisterToggle(registry, "BtnExportWizard", (context, intent) =>
                SetModelessWindow(context, "BtnExportWizard", intent,
                    () => new UI.ExportWizard(manager: context.PowerPoint)));
            RegisterToggle(registry, "BtnQRCode", (context, intent) =>
                SetModelessWindow(context, "BtnQRCode", intent,
                    () => new UI.QRCodeWindow(context.PowerPoint)));
            RegisterToggle(registry, "BtnSeriesGenerator", (context, intent) =>
                SetModelessWindow(context, "BtnSeriesGenerator", intent,
                    () => new UI.Dialogs.SeriesGeneratorWindow(context.PowerPoint)));
            RegisterToggle(registry, "BtnSlideNotes", (context, intent) =>
                SetModelessWindow(context, "BtnSlideNotes", intent,
                    () => new UI.Toolbars.SlideNotesWindow(context.PowerPoint, null)));
            RegisterToggle(registry, "BtnTransparency", (context, intent) =>
                SetModelessWindow(context, "BtnTransparency", intent,
                    () => new UI.Toolbars.TransparencyToolbar(context.PowerPoint)));

            RegisterToggle(registry, "BtnSnapToGrid", (context, intent) =>
                ExecuteBooleanToggle(
                    intent,
                    () => SnapToGridFeature.IsEnabled,
                    enabled => SnapToGridFeature.SetEnabled(enabled)));
            RegisterToggle(registry, "BtnSnapToObjects", (context, intent) =>
                ExecuteBooleanToggle(
                    intent,
                    () => SnapToObjectsFeature.IsEnabled,
                    enabled => SnapToObjectsFeature.SetEnabled(enabled)));
            RegisterToggle(registry, "BtnQuickContextMenu", (context, intent) =>
                ExecuteBooleanToggle(
                    intent,
                    () => QuickContextMenuFeature.IsEnabled,
                    enabled => QuickContextMenuFeature.SetEnabled(enabled)));
            RegisterToggle(registry, "BtnPerformanceMode", (context, intent) =>
                ExecuteBooleanToggle(
                    intent,
                    () => ExternalAddInPerformanceModeFeature.IsEnabled(context.PowerPoint),
                    enabled => ExternalAddInPerformanceModeFeature.SetEnabled(context.PowerPoint, enabled)));
            RegisterToggle(registry, "BtnLockDimensions", ExecuteDimensionLockCommand);
            RegisterToggle(registry, "BtnShapeLock", ExecuteShapeLockCommand);
        }

        private static CommandDispatchResult SetModelessWindow(
            CommandContext context,
            string id,
            ToggleIntent intent,
            Func<Window> createWindow)
        {
            bool state = context.Ui.SetModelessWindowState(id, intent, createWindow);
            return CommandDispatchResult.Succeeded(state);
        }

        private static CommandDispatchResult ExecuteBooleanToggle(
            ToggleIntent intent,
            Func<bool> getState,
            Func<bool, bool> setState)
        {
            bool requested = intent == ToggleIntent.EnsureOn ||
                (intent == ToggleIntent.Toggle && !getState());
            bool succeeded = setState(requested);
            bool actual = getState();
            return succeeded
                ? CommandDispatchResult.Succeeded(actual)
                : CommandDispatchResult.Failed(toggleState: actual);
        }

        private static CommandDispatchResult ExecuteDimensionLockCommand(
            CommandContext context,
            ToggleIntent intent)
        {
            var shapes = context.PowerPoint.GetSelectedShapes();
            bool current = shapes != null && shapes.Count > 0 &&
                shapes.Cast<Microsoft.Office.Interop.PowerPoint.Shape>()
                    .All(DimensionLockFeature.IsDimensionLocked);
            bool requested = intent == ToggleIntent.EnsureOn ||
                (intent == ToggleIntent.Toggle && !current);
            bool succeeded = DimensionLockFeature.SetSelectionDimensionLock(
                context.PowerPoint,
                requested);
            shapes = context.PowerPoint.GetSelectedShapes();
            bool actual = shapes != null && shapes.Count > 0 &&
                shapes.Cast<Microsoft.Office.Interop.PowerPoint.Shape>()
                    .All(DimensionLockFeature.IsDimensionLocked);
            return succeeded
                ? CommandDispatchResult.Succeeded(actual)
                : CommandDispatchResult.Failed(toggleState: actual);
        }

        private static CommandDispatchResult ExecuteShapeLockCommand(
            CommandContext context,
            ToggleIntent intent)
        {
            var shapes = context.PowerPoint.GetSelectedShapes();
            bool current = shapes != null && shapes.Count > 0 &&
                shapes.Cast<Microsoft.Office.Interop.PowerPoint.Shape>()
                    .All(ShapeLockingFeature.IsShapeLocked);
            bool requested = intent == ToggleIntent.EnsureOn ||
                (intent == ToggleIntent.Toggle && !current);
            bool succeeded = ShapeLockingFeature.SetSelectionShapeLock(
                context.PowerPoint,
                requested);
            shapes = context.PowerPoint.GetSelectedShapes();
            bool actual = shapes != null && shapes.Count > 0 &&
                shapes.Cast<Microsoft.Office.Interop.PowerPoint.Shape>()
                    .All(ShapeLockingFeature.IsShapeLocked);
            return succeeded
                ? CommandDispatchResult.Succeeded(actual)
                : CommandDispatchResult.Failed(toggleState: actual);
        }

        private static CommandDispatchResult FromFeatureResult(bool succeeded)
        {
            return succeeded
                ? CommandDispatchResult.Succeeded()
                : CommandDispatchResult.Failed();
        }

        private static CommandDispatchResult ExecuteFlightModeCommand(
            CommandContext context,
            FlightModeFeature.FlightModeType requestedMode,
            ToggleIntent intent)
        {
            var manager = context.PowerPoint;
            var currentMode = FlightModeFeature.GetMode(manager);
            bool succeeded;

            if (intent == ToggleIntent.EnsureOff ||
                requestedMode == FlightModeFeature.FlightModeType.None)
            {
                succeeded = currentMode == FlightModeFeature.FlightModeType.None ||
                    FlightModeFeature.Disable(manager);
            }
            else if (intent == ToggleIntent.Toggle)
            {
                succeeded = currentMode == FlightModeFeature.FlightModeType.None
                    ? EnableFlightMode(manager, requestedMode)
                    : FlightModeFeature.Disable(manager);
            }
            else
            {
                succeeded = currentMode == requestedMode ||
                    EnableFlightMode(manager, requestedMode);
            }

            bool isEnabled = FlightModeFeature.GetMode(manager) !=
                FlightModeFeature.FlightModeType.None;
            return succeeded
                ? CommandDispatchResult.Succeeded(isEnabled)
                : CommandDispatchResult.Failed(toggleState: isEnabled);
        }

        private static bool EnableFlightMode(
            PowerPointManager manager,
            FlightModeFeature.FlightModeType mode)
        {
            switch (mode)
            {
                case FlightModeFeature.FlightModeType.CoverLogos:
                    return FlightModeFeature.EnableCoverLogos(manager);
                case FlightModeFeature.FlightModeType.LightOverlay:
                    return FlightModeFeature.EnableLightOverlay(manager);
                case FlightModeFeature.FlightModeType.HeavyOverlay:
                    return FlightModeFeature.EnableHeavyOverlay(manager);
                case FlightModeFeature.FlightModeType.HideBackground:
                    return FlightModeFeature.EnableHideBackground(manager);
                default:
                    return FlightModeFeature.Disable(manager);
            }
        }

        private CommandDispatchResult ExecuteRegisteredCommand<TArguments>(
            CommandKey<TArguments> key,
            TArguments arguments,
            CommandInvocationSource source,
            FrameworkElement target = null,
            ToggleIntent toggleIntent = ToggleIntent.Toggle)
        {
            var ui = new SidebarCommandUiAdapter(this, target);
            var manager = GetManager();
            if (manager == null)
            {
                ExceptionLogger.Log(
                    new InvalidOperationException(
                        $"PowerPoint manager is unavailable for '{key.Id}' from {source}."),
                    "MainSidebar.ExecuteRegisteredCommand");
                var failure = CommandDispatchResult.Failed(
                    "openEfficiency could not access PowerPoint for this command. Please try again. Details were written to the local diagnostics log.");
                CompleteRegisteredCommand(failure, ui);
                return failure;
            }

            var context = new CommandContext(manager, source, ui);
            var result = Registry.Value.Execute(key, context, arguments, toggleIntent);
            CompleteRegisteredCommand(result, ui);
            return result;
        }

        private static ToggleIntent GetToggleIntent(
            CommandInvocationSource source,
            FrameworkElement target)
        {
            if (source != CommandInvocationSource.Sidebar || !(target is ToggleButton toggleButton))
                return ToggleIntent.Toggle;
            if (toggleButton.IsChecked == true) return ToggleIntent.EnsureOn;
            if (toggleButton.IsChecked == false) return ToggleIntent.EnsureOff;
            return ToggleIntent.Toggle;
        }

        private static void CompleteRegisteredCommand(
            CommandDispatchResult result,
            ICommandUiAdapter ui)
        {
            if (result == null || ui == null) return;
            if (result.ToggleState.HasValue)
                ui.ApplyToggleState(result.ToggleState);
            if (!string.IsNullOrWhiteSpace(result.UserMessage))
                ui.ShowError(result.UserMessage);
            if (result.ToggleState.HasValue)
                ui.RefreshCommandStates();
        }

        private sealed class SidebarCommandUiAdapter : ICommandUiAdapter
        {
            private readonly MainSidebar _sidebar;
            private readonly FrameworkElement _target;

            public SidebarCommandUiAdapter(MainSidebar sidebar, FrameworkElement target)
            {
                _sidebar = sidebar ?? throw new ArgumentNullException(nameof(sidebar));
                _target = target;
            }

            public bool TryOpenContextMenu()
            {
                if (_target?.ContextMenu == null) return false;
                _target.ContextMenu.PlacementTarget = _target;
                _target.ContextMenu.IsOpen = true;
                return true;
            }

            public bool? ShowDialog(Func<Window> createDialog)
            {
                if (createDialog == null) throw new ArgumentNullException(nameof(createDialog));
                var dialog = createDialog();
                if (dialog == null) throw new InvalidOperationException("Dialog factory returned null.");

                var app = _sidebar.GetManager()?.GetApplication();
                if (app != null)
                    new System.Windows.Interop.WindowInteropHelper(dialog).Owner = (IntPtr)app.HWND;
                return dialog.ShowDialog();
            }

            public bool SetModelessWindowState(
                string commandId,
                ToggleIntent intent,
                Func<Window> createWindow)
            {
                bool isOpen = _sidebar._floatingWindows.TryGetValue(commandId, out var existing);
                if (isOpen && !existing.IsLoaded)
                {
                    _sidebar._floatingWindows.Remove(commandId);
                    existing = null;
                    isOpen = false;
                }

                if (intent == ToggleIntent.EnsureOff ||
                    (intent == ToggleIntent.Toggle && isOpen))
                {
                    if (isOpen) existing.Close();
                    _sidebar._floatingWindows.Remove(commandId);
                    return false;
                }

                if (isOpen)
                {
                    SetPowerPointOwner(existing);
                    if (!existing.IsVisible) existing.Show();
                    if (existing.WindowState == WindowState.Minimized)
                        existing.WindowState = WindowState.Normal;
                    existing.Activate();
                    existing.Focus();
                    return true;
                }

                if (createWindow == null) throw new ArgumentNullException(nameof(createWindow));
                var window = createWindow();
                if (window == null) throw new InvalidOperationException("Window factory returned null.");
                SetPowerPointOwner(window);
                window.Closed += (sender, args) =>
                {
                    if (_sidebar._floatingWindows.TryGetValue(commandId, out var tracked) &&
                        ReferenceEquals(tracked, window))
                    {
                        _sidebar._floatingWindows.Remove(commandId);
                    }
                    ApplyToggleState(false);
                };
                try
                {
                    window.Show();
                    if (!window.IsLoaded || !window.IsVisible)
                        throw new InvalidOperationException(
                            $"Window for '{commandId}' closed before it became visible.");
                    _sidebar._floatingWindows[commandId] = window;
                }
                catch
                {
                    try { window.Close(); }
                    catch (Exception closeEx)
                    {
                        ExceptionLogger.Log(closeEx, "MainSidebar.ModelessWindow.ShowFailureCleanup");
                    }
                    _sidebar._floatingWindows.Remove(commandId);
                    throw;
                }
                return true;
            }

            private void SetPowerPointOwner(Window window)
            {
                var app = _sidebar.GetManager()?.GetApplication();
                if (app != null)
                    new System.Windows.Interop.WindowInteropHelper(window).Owner = (IntPtr)app.HWND;
            }

            public CommandDispatchResult ExecuteHostAction(
                CommandHostAction action,
                ToggleIntent intent)
            {
                var manager = _sidebar.GetManager();
                switch (action)
                {
                    case CommandHostAction.AdjustSpacing:
                        if (TryOpenContextMenu()) return CommandDispatchResult.Succeeded();
                        return FromFeatureResult(AdjustHorizontalSpacingFeature.Execute(
                            manager,
                            _sidebar.GetSpacingIncrement()));
                    case CommandHostAction.IncreaseHorizontalSpacing:
                        return FromFeatureResult(AdjustHorizontalSpacingFeature.Execute(
                            manager,
                            _sidebar.GetSpacingIncrement()));
                    case CommandHostAction.DecreaseHorizontalSpacing:
                        return FromFeatureResult(AdjustHorizontalSpacingFeature.Execute(
                            manager,
                            -_sidebar.GetSpacingIncrement()));
                    case CommandHostAction.IncreaseVerticalSpacing:
                        return FromFeatureResult(AdjustVerticalSpacingFeature.Execute(
                            manager,
                            _sidebar.GetSpacingIncrement()));
                    case CommandHostAction.DecreaseVerticalSpacing:
                        return FromFeatureResult(AdjustVerticalSpacingFeature.Execute(
                            manager,
                            -_sidebar.GetSpacingIncrement()));
                    case CommandHostAction.UpdateExcelCharts:
                        if (TryOpenContextMenu()) return CommandDispatchResult.Succeeded();
                        return FromFeatureResult(UpdateExcelChartsFeature.Execute(
                            manager,
                            updateAllSlides: false));
                    case CommandHostAction.OpenSettings:
                        _sidebar.OpenSettings();
                        return CommandDispatchResult.Succeeded();
                    case CommandHostAction.OpenTemplateManager:
                        _sidebar.OpenTemplateManager();
                        return CommandDispatchResult.Succeeded();
                    case CommandHostAction.OpenFeatureExplorer:
                        if (_sidebar._host == null)
                            throw new InvalidOperationException("Sidebar host is unavailable for Feature Explorer.");
                        _sidebar._host.OpenFeatureExplorer();
                        return CommandDispatchResult.Succeeded();
                    case CommandHostAction.ToggleAssetManager:
                        if (_sidebar._host == null)
                            throw new InvalidOperationException("Sidebar host is unavailable for Asset Manager.");
                        _sidebar._host.ToggleAssetManager();
                        return CommandDispatchResult.Succeeded();
                    case CommandHostAction.ShowGantt:
                        _sidebar.ShowOrCenterGanttWindow(manager, _target as ToggleButton);
                        return CommandDispatchResult.Succeeded(
                            (_target as ToggleButton)?.IsChecked);
                    case CommandHostAction.OpenParagraphDialog:
                        return _sidebar.OpenParagraphDialogFromCommand(manager);
                    case CommandHostAction.OpenSpecialCharacters:
                        _sidebar.OpenSpecialCharactersWindow(null);
                        return CommandDispatchResult.Succeeded();
                    default:
                        return CommandDispatchResult.Failed(
                            $"Unknown host action '{action}'.");
                }
            }

            public void ApplyToggleState(bool? isChecked)
            {
                if (_target is ToggleButton toggleButton)
                    toggleButton.IsChecked = isChecked;
            }

            public void RefreshCommandStates()
            {
                _sidebar.UpdateButtonStates();
            }

            public void ShowError(string message)
            {
                if (string.IsNullOrWhiteSpace(message)) return;
                MessageBox.Show(message, "openEfficiency", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private CommandDispatchResult OpenParagraphDialogFromCommand(
            PowerPointManager manager)
        {
            try
            {
                var app = manager?.GetApplication();
                if (app?.CommandBars == null)
                    return CommandDispatchResult.Failed(
                        "Cannot open Paragraph & Line Spacing. Please try again.");

                app.CommandBars.ExecuteMso("ParagraphDialog");
                return CommandDispatchResult.Succeeded();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MainSidebar.ParagraphDialog");
                return CommandDispatchResult.Failed(
                    $"Cannot open Paragraph & Line Spacing: {ex.Message}");
            }
        }
    }
}
