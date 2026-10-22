using System;
using System.Collections.Generic;
using System.Linq;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services.Commands
{
    public sealed class CommandRegistry
    {
        internal const string UserVisibleFailureMessage =
            "openEfficiency could not run this command. Please try again. Details were written to the local diagnostics log.";

        private readonly Dictionary<string, ICommandRegistration> _registrations =
            new Dictionary<string, ICommandRegistration>(StringComparer.Ordinal);
        private bool _isFrozen;

        public IReadOnlyList<CommandDescriptor> Commands => _registrations.Values
            .Select(registration => registration.Descriptor)
            .OrderBy(descriptor => descriptor.Id, StringComparer.Ordinal)
            .ToList()
            .AsReadOnly();

        public bool IsRegistered(string id)
        {
            return !string.IsNullOrWhiteSpace(id) && _registrations.ContainsKey(id);
        }

        public void RegisterAction<TArguments>(
            CommandKey<TArguments> key,
            TArguments primaryArguments,
            Func<CommandContext, TArguments, CommandDispatchResult> execute)
        {
            if (execute == null) throw new ArgumentNullException(nameof(execute));
            Register(new ActionRegistration<TArguments>(key, primaryArguments, execute));
        }

        public void RegisterToggle<TArguments>(
            CommandKey<TArguments> key,
            TArguments primaryArguments,
            Func<CommandContext, TArguments, ToggleIntent, CommandDispatchResult> execute)
        {
            if (execute == null) throw new ArgumentNullException(nameof(execute));
            Register(new ToggleRegistration<TArguments>(key, primaryArguments, execute));
        }

        public void Freeze()
        {
            _isFrozen = true;
        }

        public CommandDispatchResult Execute<TArguments>(
            CommandKey<TArguments> key,
            CommandContext context,
            TArguments arguments,
            ToggleIntent toggleIntent = ToggleIntent.Toggle)
        {
            if (key == null) throw new ArgumentNullException(nameof(key));
            return ExecuteRegistration(key.Id, context, arguments, typeof(TArguments), toggleIntent);
        }

        public CommandDispatchResult ExecuteDefault(
            string id,
            CommandContext context,
            ToggleIntent toggleIntent = ToggleIntent.Toggle)
        {
            if (string.IsNullOrWhiteSpace(id)) return CommandDispatchResult.NotRegistered();
            if (!_registrations.TryGetValue(id, out var registration))
                return CommandDispatchResult.NotRegistered();

            return ExecuteRegistration(
                id,
                context,
                registration.PrimaryArguments,
                registration.Descriptor.ArgumentType,
                toggleIntent);
        }

        private void Register(ICommandRegistration registration)
        {
            if (_isFrozen)
                throw new CommandRegistrationException("The command registry is frozen.");
            if (registration == null)
                throw new ArgumentNullException(nameof(registration));
            if (_registrations.ContainsKey(registration.Descriptor.Id))
                throw new CommandRegistrationException(
                    $"Duplicate command ID '{registration.Descriptor.Id}'.");

            _registrations.Add(registration.Descriptor.Id, registration);
        }

        private CommandDispatchResult ExecuteRegistration(
            string id,
            CommandContext context,
            object arguments,
            Type suppliedArgumentType,
            ToggleIntent toggleIntent)
        {
            if (!_registrations.TryGetValue(id, out var registration))
                return CommandDispatchResult.NotRegistered();

            if (context == null)
                throw new ArgumentNullException(nameof(context));

            if (registration.Descriptor.ArgumentType != suppliedArgumentType ||
                !registration.Accepts(arguments))
            {
                var exception = new InvalidOperationException(
                    $"Command '{id}' expects arguments of type " +
                    $"'{registration.Descriptor.ArgumentType.FullName}', but received " +
                    $"'{suppliedArgumentType?.FullName ?? "<null>"}'.");
                ExceptionLogger.Log(exception, $"CommandRegistry.Execute({id},{context.Source})");
                return CommandDispatchResult.Failed(UserVisibleFailureMessage);
            }

            try
            {
                var result = registration.Execute(context, arguments, toggleIntent);
                if (result != null) return result;

                var exception = new InvalidOperationException(
                    $"Command '{id}' returned no dispatch result.");
                ExceptionLogger.Log(exception, $"CommandRegistry.Execute({id},{context.Source})");
                return CommandDispatchResult.Failed(UserVisibleFailureMessage);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, $"CommandRegistry.Execute({id},{context.Source})");
                return CommandDispatchResult.Failed(UserVisibleFailureMessage);
            }
        }

        private interface ICommandRegistration
        {
            CommandDescriptor Descriptor { get; }
            object PrimaryArguments { get; }
            bool Accepts(object arguments);
            CommandDispatchResult Execute(
                CommandContext context,
                object arguments,
                ToggleIntent toggleIntent);
        }

        private abstract class CommandRegistration<TArguments> : ICommandRegistration
        {
            protected CommandRegistration(
                CommandKey<TArguments> key,
                TArguments primaryArguments,
                CommandKind kind)
            {
                if (key == null) throw new ArgumentNullException(nameof(key));
                Descriptor = new CommandDescriptor(key.Id, kind, typeof(TArguments));
                PrimaryArguments = primaryArguments;
            }

            public CommandDescriptor Descriptor { get; }
            public object PrimaryArguments { get; }

            public bool Accepts(object arguments)
            {
                if (arguments is TArguments) return true;
                return arguments == null &&
                    (!typeof(TArguments).IsValueType || Nullable.GetUnderlyingType(typeof(TArguments)) != null);
            }

            public abstract CommandDispatchResult Execute(
                CommandContext context,
                object arguments,
                ToggleIntent toggleIntent);

            protected static TArguments Cast(object arguments)
            {
                return arguments == null ? default(TArguments) : (TArguments)arguments;
            }
        }

        private sealed class ActionRegistration<TArguments> : CommandRegistration<TArguments>
        {
            private readonly Func<CommandContext, TArguments, CommandDispatchResult> _execute;

            public ActionRegistration(
                CommandKey<TArguments> key,
                TArguments primaryArguments,
                Func<CommandContext, TArguments, CommandDispatchResult> execute)
                : base(key, primaryArguments, CommandKind.Action)
            {
                _execute = execute;
            }

            public override CommandDispatchResult Execute(
                CommandContext context,
                object arguments,
                ToggleIntent toggleIntent)
            {
                return _execute(context, Cast(arguments));
            }
        }

        private sealed class ToggleRegistration<TArguments> : CommandRegistration<TArguments>
        {
            private readonly Func<CommandContext, TArguments, ToggleIntent, CommandDispatchResult> _execute;

            public ToggleRegistration(
                CommandKey<TArguments> key,
                TArguments primaryArguments,
                Func<CommandContext, TArguments, ToggleIntent, CommandDispatchResult> execute)
                : base(key, primaryArguments, CommandKind.Toggle)
            {
                _execute = execute;
            }

            public override CommandDispatchResult Execute(
                CommandContext context,
                object arguments,
                ToggleIntent toggleIntent)
            {
                return _execute(context, Cast(arguments), toggleIntent);
            }
        }
    }
}
