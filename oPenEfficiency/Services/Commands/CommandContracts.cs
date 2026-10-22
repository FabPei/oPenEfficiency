using System;
using System.Windows;

namespace oPenEfficiency.Services.Commands
{
    public sealed class CommandKey<TArguments>
    {
        public CommandKey(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
                throw new ArgumentException("Command ID must not be empty.", nameof(id));

            Id = id;
        }

        public string Id { get; }
    }

    public struct NoCommandArguments
    {
        public static readonly NoCommandArguments Value = new NoCommandArguments();
    }

    public enum CommandKind
    {
        Action,
        Toggle
    }

    public enum CommandInvocationSource
    {
        Sidebar,
        Search,
        Shortcut,
        ContextMenu,
        Ribbon,
        Internal
    }

    public enum ToggleIntent
    {
        Toggle,
        EnsureOn,
        EnsureOff
    }

    public enum CommandHostAction
    {
        AdjustSpacing,
        IncreaseHorizontalSpacing,
        DecreaseHorizontalSpacing,
        IncreaseVerticalSpacing,
        DecreaseVerticalSpacing,
        UpdateExcelCharts,
        OpenSettings,
        OpenTemplateManager,
        OpenFeatureExplorer,
        ToggleAssetManager,
        ShowGantt,
        OpenParagraphDialog,
        OpenSpecialCharacters
    }

    public interface ICommandUiAdapter
    {
        bool TryOpenContextMenu();
        bool? ShowDialog(Func<Window> createDialog);
        bool SetModelessWindowState(string commandId, ToggleIntent intent, Func<Window> createWindow);
        CommandDispatchResult ExecuteHostAction(CommandHostAction action, ToggleIntent intent);
        void ApplyToggleState(bool? isChecked);
        void RefreshCommandStates();
        void ShowError(string message);
    }

    public sealed class CommandContext
    {
        public CommandContext(
            PowerPointManager powerPoint,
            CommandInvocationSource source,
            ICommandUiAdapter ui)
        {
            PowerPoint = powerPoint ?? throw new ArgumentNullException(nameof(powerPoint));
            Ui = ui ?? throw new ArgumentNullException(nameof(ui));
            Source = source;
        }

        public PowerPointManager PowerPoint { get; }
        public CommandInvocationSource Source { get; }
        public ICommandUiAdapter Ui { get; }
    }

    public enum CommandDispatchStatus
    {
        NotRegistered,
        Succeeded,
        Cancelled,
        Failed
    }

    public sealed class CommandDispatchResult
    {
        private CommandDispatchResult(
            CommandDispatchStatus status,
            bool? toggleState,
            string userMessage)
        {
            Status = status;
            ToggleState = toggleState;
            UserMessage = userMessage;
        }

        public CommandDispatchStatus Status { get; }
        public bool? ToggleState { get; }
        public string UserMessage { get; }

        public static CommandDispatchResult NotRegistered() =>
            new CommandDispatchResult(CommandDispatchStatus.NotRegistered, null, null);

        public static CommandDispatchResult Succeeded(bool? toggleState = null) =>
            new CommandDispatchResult(CommandDispatchStatus.Succeeded, toggleState, null);

        public static CommandDispatchResult Cancelled(bool? toggleState = null) =>
            new CommandDispatchResult(CommandDispatchStatus.Cancelled, toggleState, null);

        public static CommandDispatchResult Failed(
            string userMessage = null,
            bool? toggleState = null) =>
            new CommandDispatchResult(CommandDispatchStatus.Failed, toggleState, userMessage);
    }

    public sealed class CommandDescriptor
    {
        internal CommandDescriptor(string id, CommandKind kind, Type argumentType)
        {
            Id = id;
            Kind = kind;
            ArgumentType = argumentType;
        }

        public string Id { get; }
        public CommandKind Kind { get; }
        public Type ArgumentType { get; }
    }

    public sealed class CommandRegistrationException : InvalidOperationException
    {
        public CommandRegistrationException(string message)
            : base(message)
        {
        }
    }
}
