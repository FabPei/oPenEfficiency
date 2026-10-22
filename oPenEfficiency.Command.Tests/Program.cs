using System;
using System.Windows;
using System.Runtime.Serialization;
using System.Reflection;
using System.Xml;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Commands;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace oPenEfficiency.Command.Tests
{
    [TestClass]
    public class Program
    {
        private enum Direction
        {
            Right,
            Left
        }

        [TestMethod] public void DefaultArgumentsAreExplicit() => ExecuteDefaultUsesExplicitPrimaryArguments();
        [TestMethod] public void TypedArgumentsAreForwarded() => ExecuteUsesTypedArguments();
        [TestMethod] public void ToggleIntentsAreForwarded() => ToggleReceivesEveryIntent();
        [TestMethod] public void UnknownIdsAreRejected() => UnknownIdIsNotRegistered();
        [TestMethod] public void DuplicateAndFrozenRegistrationsFail() => DuplicateAndFrozenRegistrationsFailControlled();
        [TestMethod] public void HandlerAndArgumentFailuresAreTerminal() => WrongArgumentTypeAndHandlerExceptionsAreTerminalFailures();
        [TestMethod] public void HandlerFailuresReturnActionableUserMessage() => HandlerFailureIncludesUserMessage();
        [TestMethod] public void AgendaXmlParserCreatesUsableLayout() => AgendaXmlParsingProducesUsableShapeBounds();
        [TestMethod] public void AgendaXmlImportMergesWithoutReplacingLayouts() => AgendaXmlImportMergeIsAtomic();
        [TestMethod] public void AgendaXmlImportReportsStorageFailure() => AgendaXmlImportStorageFailureIsTerminal();
        [TestMethod] public void AgendaXmlFolderImportFindsBothNames() => AgendaXmlFolderImportSupportsCurrentAndLegacyNames();
        [TestMethod] public void AgendaTransactionUsesUndoBeforeFallback() => AgendaTransactionRollbackOrderIsStable();
        [TestMethod] public void AgendaSynchronizationRequiresUndo() => AgendaTransactionRequiresUndoForSynchronization();
        [TestMethod] public void AgendaRollbackVerifiesFullStructure() => AgendaRollbackSnapshotVerificationIsStable();

        private static void ExecuteDefaultUsesExplicitPrimaryArguments()
        {
            var key = new CommandKey<Direction>("Parameterized");
            var registry = new CommandRegistry();
            Direction received = Direction.Left;
            registry.RegisterAction(
                key,
                Direction.Right,
                (context, direction) =>
                {
                    received = direction;
                    return CommandDispatchResult.Succeeded();
                });
            registry.Freeze();

            var result = registry.ExecuteDefault("Parameterized", CreateContext());
            Assert(result.Status == CommandDispatchStatus.Succeeded,
                "Default execution did not succeed.");
            Assert(received == Direction.Right,
                "Default execution did not use the explicit primary arguments.");
        }

        private static void ExecuteUsesTypedArguments()
        {
            var key = new CommandKey<Direction>("Typed");
            var registry = new CommandRegistry();
            Direction received = Direction.Right;
            registry.RegisterAction(
                key,
                Direction.Right,
                (context, direction) =>
                {
                    received = direction;
                    return CommandDispatchResult.Succeeded();
                });
            registry.Freeze();

            var result = registry.Execute(key, CreateContext(), Direction.Left);
            Assert(result.Status == CommandDispatchStatus.Succeeded && received == Direction.Left,
                "Typed execution lost the supplied argument.");
            Assert(registry.Commands.Count == 1 &&
                registry.Commands[0].ArgumentType == typeof(Direction) &&
                registry.Commands[0].Kind == CommandKind.Action,
                "Command descriptor does not expose the registered typed contract.");
        }

        private static void ToggleReceivesEveryIntent()
        {
            var key = new CommandKey<NoCommandArguments>("Toggle");
            var registry = new CommandRegistry();
            ToggleIntent received = ToggleIntent.Toggle;
            registry.RegisterToggle(
                key,
                NoCommandArguments.Value,
                (context, arguments, intent) =>
                {
                    received = intent;
                    return CommandDispatchResult.Succeeded(intent != ToggleIntent.EnsureOff);
                });
            registry.Freeze();

            foreach (var intent in new[]
            {
                ToggleIntent.Toggle,
                ToggleIntent.EnsureOn,
                ToggleIntent.EnsureOff
            })
            {
                var result = registry.Execute(
                    key,
                    CreateContext(),
                    NoCommandArguments.Value,
                    intent);
                Assert(result.Status == CommandDispatchStatus.Succeeded && received == intent,
                    $"Toggle intent '{intent}' was not forwarded.");
                Assert(result.ToggleState == (intent != ToggleIntent.EnsureOff),
                    $"Toggle state for '{intent}' was not preserved.");
            }
        }

        private static void UnknownIdIsNotRegistered()
        {
            var registry = new CommandRegistry();
            registry.Freeze();
            var result = registry.ExecuteDefault("Unknown", CreateContext());
            Assert(result.Status == CommandDispatchStatus.NotRegistered,
                "Unknown command ID did not return NotRegistered.");
            Assert(!registry.IsRegistered("Unknown"),
                "Unknown command ID was reported as registered.");
        }

        private static void DuplicateAndFrozenRegistrationsFailControlled()
        {
            var key = new CommandKey<NoCommandArguments>("Duplicate");
            var registry = new CommandRegistry();
            registry.RegisterAction(
                key,
                NoCommandArguments.Value,
                (context, arguments) => CommandDispatchResult.Succeeded());

            AssertThrows<CommandRegistrationException>(() =>
                registry.RegisterAction(
                    key,
                    NoCommandArguments.Value,
                    (context, arguments) => CommandDispatchResult.Succeeded()),
                "Duplicate command ID was accepted.");

            registry.Freeze();
            AssertThrows<CommandRegistrationException>(() =>
                registry.RegisterAction(
                    new CommandKey<NoCommandArguments>("AfterFreeze"),
                    NoCommandArguments.Value,
                    (context, arguments) => CommandDispatchResult.Succeeded()),
                "Registration after Freeze was accepted.");
        }

        private static void WrongArgumentTypeAndHandlerExceptionsAreTerminalFailures()
        {
            var integerKey = new CommandKey<int>("TypedFailure");
            var registry = new CommandRegistry();
            registry.RegisterAction(
                integerKey,
                1,
                (context, argument) => CommandDispatchResult.Succeeded());
            registry.RegisterAction(
                new CommandKey<NoCommandArguments>("Throws"),
                NoCommandArguments.Value,
                (context, argument) => throw new InvalidOperationException("Expected test failure"));
            registry.Freeze();

            var wrongKey = new CommandKey<string>("TypedFailure");
            var wrongType = registry.Execute(wrongKey, CreateContext(), "wrong");
            Assert(wrongType.Status == CommandDispatchStatus.Failed,
                "Wrong command argument type was not a controlled failure.");

            var thrown = registry.ExecuteDefault("Throws", CreateContext());
            Assert(thrown.Status == CommandDispatchStatus.Failed,
                "Handler exception was not converted to Failed.");
        }

        private static void HandlerFailureIncludesUserMessage()
        {
            string commandId = "BrokenWindow-" + Guid.NewGuid().ToString("N");
            var registry = new CommandRegistry();
            registry.RegisterAction(
                new CommandKey<NoCommandArguments>(commandId),
                NoCommandArguments.Value,
                (context, argument) => throw new InvalidOperationException("Sensitive detail"));
            registry.Freeze();

            CommandDispatchResult result = registry.ExecuteDefault(commandId, CreateContext());

            Assert(result.Status == CommandDispatchStatus.Failed,
                "Handler exception was not converted to Failed.");
            Assert(!string.IsNullOrWhiteSpace(result.UserMessage),
                "Handler exception did not produce a user-visible failure message.");
            Assert(!result.UserMessage.Contains("Sensitive detail"),
                "User-visible failure message leaked internal exception detail.");
            string diagnostics = File.ReadAllText(ExceptionLogger.GetLogFilePath());
            Assert(diagnostics.Contains("CommandRegistry.Execute(" + commandId + ",Internal)"),
                "Handler exception diagnostics omitted the command ID and invocation source.");
        }

        private static void AgendaXmlParsingProducesUsableShapeBounds()
        {
            var document = new XmlDocument();
            document.LoadXml(
                "<layouts><layout name='Imported Layout'>" +
                "<standard><font name='Arial' size='18' color='1'/><paragraphformat alignment='1'/></standard>" +
                "<position left='40' top='120'/>" +
                "<cases>" +
                "<case level='1' selected='1'><element type='text'><font bold='1'/></element></case>" +
                "<case level='2' selected='1'><element type='text'><font italic='1'/></element></case>" +
                "</cases></layout></layouts>");

            MethodInfo parser = typeof(ExternalAgendaImportService).GetMethod(
                "ParseLayout",
                BindingFlags.Instance | BindingFlags.NonPublic);
            Assert(parser != null, "Agenda XML parser could not be resolved.");

            var layout = (AgendaLayoutConfig)parser.Invoke(
                new ExternalAgendaImportService(),
                new object[] { document.SelectSingleNode("//layout") });

            Assert(layout != null && layout.LayoutName == "Imported Layout",
                "Agenda XML layout metadata was not parsed.");
            Assert(layout.ItemTextShape != null &&
                   layout.ItemTextShape.Width > 0 &&
                   layout.ItemTextShape.Height > 0 &&
                   layout.ItemTextShape.Left == 40 &&
                   layout.ItemTextShape.Top == 120,
                "Imported agenda item geometry is not usable by Agenda Creator.");
            Assert(layout.SubItemTextShape != null &&
                   layout.SubItemTextShape.Left > layout.ItemTextShape.Left,
                "Imported agenda subitem geometry does not preserve hierarchy.");

            document.LoadXml("<layouts><layout name='Minimal Layout'/></layouts>");
            var minimalLayout = (AgendaLayoutConfig)parser.Invoke(
                new ExternalAgendaImportService(),
                new object[] { document.SelectSingleNode("//layout") });
            Assert(minimalLayout.ItemTextShape != null && minimalLayout.ItemTextShape.Width > 0,
                "Minimal external agenda XML did not receive usable fallback geometry.");
        }

        private static void AgendaXmlImportMergeIsAtomic()
        {
            string directory = Path.Combine(Path.GetTempPath(), "oPE_AgendaImport_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            string xmlPath = Path.Combine(directory, "AgendaCreator.xml");
            try
            {
                File.WriteAllText(
                    xmlPath,
                    "<layouts>" +
                    "<layout name='existing layout'><cases><case level='1' selected='1'/></cases></layout>" +
                    "<layout name='Imported Layout'><position left='30' top='90'/><cases><case level='1' selected='1'/></cases></layout>" +
                    "</layouts>");

                var store = new FakeAgendaLayoutStore
                {
                    Layouts = new List<AgendaLayoutConfig>
                    {
                        new AgendaLayoutConfig { LayoutName = "Existing Layout" }
                    }
                };
                var imported = new ExternalAgendaImportService(store).ImportFromXml(xmlPath);

                Assert(imported.SequenceEqual(new[] { "Imported Layout" }),
                    "Agenda XML import did not skip a case-insensitive duplicate.");
                Assert(store.SaveCalls == 1 && store.SavedLayouts.Count == 2,
                    "Agenda XML import did not merge existing and imported layouts atomically.");
                Assert(store.SavedLayouts.Any(layout => layout.LayoutName == "Existing Layout") &&
                       store.SavedLayouts.Any(layout =>
                           layout.LayoutName == "Imported Layout" &&
                           layout.ItemTextShape != null &&
                           layout.ItemTextShape.Width > 0),
                    "Agenda XML import replaced existing data or saved an unusable layout.");
            }
            finally
            {
                if (Directory.Exists(directory)) Directory.Delete(directory, true);
            }
        }

        private static void AgendaXmlImportStorageFailureIsTerminal()
        {
            string directory = Path.Combine(Path.GetTempPath(), "oPE_AgendaImport_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            string xmlPath = Path.Combine(directory, "AgendaWizard.xml");
            try
            {
                File.WriteAllText(
                    xmlPath,
                    "<layouts><layout name='Imported Layout'><cases><case level='1' selected='1'/></cases></layout></layouts>");
                var store = new FakeAgendaLayoutStore { SaveResult = false };

                var imported = new ExternalAgendaImportService(store).ImportFromXml(xmlPath);

                Assert(imported.Count == 0 && store.SaveCalls == 1,
                    "Agenda XML import reported success after storage rejected the write.");
            }
            finally
            {
                if (Directory.Exists(directory)) Directory.Delete(directory, true);
            }
        }

        private static void AgendaXmlFolderImportSupportsCurrentAndLegacyNames()
        {
            string directory = Path.Combine(Path.GetTempPath(), "oPE_AgendaImport_" + Guid.NewGuid().ToString("N"));
            string nestedDirectory = Path.Combine(directory, "Nested");
            Directory.CreateDirectory(nestedDirectory);
            try
            {
                File.WriteAllText(
                    Path.Combine(directory, "AgendaCreator.xml"),
                    "<layouts><layout name='Current XML'><cases><case level='1' selected='1'/></cases></layout></layouts>");
                File.WriteAllText(
                    Path.Combine(nestedDirectory, "AgendaWizard.xml"),
                    "<layouts><layout name='Legacy XML'><cases><case level='1' selected='1'/></cases></layout></layouts>");
                File.WriteAllText(
                    Path.Combine(directory, "Unrelated.xml"),
                    "<layouts><layout name='Must Not Import'/></layouts>");

                var store = new FakeAgendaLayoutStore();
                var imported = new ExternalAgendaImportService(store).ImportFromFolder(directory);

                Assert(imported.Count == 2 &&
                       imported.Contains("Current XML") &&
                       imported.Contains("Legacy XML") &&
                       !imported.Contains("Must Not Import"),
                    "Agenda folder import did not select both supported XML file names exclusively.");
                Assert(store.Layouts.Count == 2,
                    "Agenda folder import did not persist both discovered layouts.");
            }
            finally
            {
                if (Directory.Exists(directory)) Directory.Delete(directory, true);
            }
        }

        private static void AgendaTransactionRollbackOrderIsStable()
        {
            var undoHost = new FakeAgendaMutationHost { UndoResult = true };
            var transaction = new AgendaMutationTransaction(undoHost);
            transaction.Begin(true);
            transaction.MarkMutationAttempted();
            transaction.RollBack();
            Assert(undoHost.UndoCalls == 1 && undoHost.FallbackCalls == 0,
                "Agenda rollback did not prefer the native Undo transaction.");

            var fallbackHost = new FakeAgendaMutationHost { UndoResult = false };
            transaction = new AgendaMutationTransaction(fallbackHost);
            transaction.Begin(false);
            transaction.MarkMutationAttempted();
            transaction.RollBack();
            Assert(fallbackHost.UndoCalls == 1 && fallbackHost.FallbackCalls == 1,
                "Agenda rollback did not run fallback cleanup after Undo failed.");

            var untouchedHost = new FakeAgendaMutationHost();
            transaction = new AgendaMutationTransaction(untouchedHost);
            transaction.Begin(false);
            transaction.RollBack();
            Assert(untouchedHost.UndoCalls == 0 && untouchedHost.FallbackCalls == 0,
                "Agenda rollback mutated state even though no mutation had started.");
        }

        private static void AgendaTransactionRequiresUndoForSynchronization()
        {
            var host = new FakeAgendaMutationHost { StartUndoResult = false };
            var transaction = new AgendaMutationTransaction(host);

            AssertThrows<InvalidOperationException>(
                () => transaction.Begin(true),
                "Agenda synchronization started without a safe Undo entry.");
            Assert(host.CaptureCalls == 1 && host.UndoCalls == 0 && host.FallbackCalls == 0,
                "Failed transaction setup attempted rollback mutations.");
        }

        private static void AgendaRollbackSnapshotVerificationIsStable()
        {
            var snapshot = new AgendaMutationSnapshot(
                new[] { 10, 20 },
                "Layout",
                new[] { "Introduction", "Chapter" },
                new Dictionary<int, int> { [10] = 3, [20] = 5 });

            Assert(snapshot.MatchesPresentationStructure(
                    new[] { 10, 20 },
                    new[] { "Introduction", "Chapter" },
                    new Dictionary<int, int> { [10] = 3, [20] = 5 }),
                "Agenda rollback rejected an exact presentation snapshot.");
            Assert(!snapshot.MatchesPresentationStructure(
                    new[] { 20, 10 },
                    new[] { "Introduction", "Chapter" },
                    new Dictionary<int, int> { [10] = 3, [20] = 5 }),
                "Agenda rollback accepted changed slide order.");
            Assert(!snapshot.MatchesPresentationStructure(
                    new[] { 10, 20 },
                    new[] { "Introduction", "Changed" },
                    new Dictionary<int, int> { [10] = 3, [20] = 5 }),
                "Agenda rollback accepted changed sections.");
            Assert(!snapshot.MatchesPresentationStructure(
                    new[] { 10, 20 },
                    new[] { "Introduction", "Chapter" },
                    new Dictionary<int, int> { [10] = 2, [20] = 5 }),
                "Agenda rollback accepted changed slide shapes.");
        }

        private static CommandContext CreateContext()
        {
            return new CommandContext(
                (PowerPointManager)FormatterServices.GetUninitializedObject(typeof(PowerPointManager)),
                CommandInvocationSource.Internal,
                new FakeUiAdapter());
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        private static void AssertThrows<TException>(Action action, string message)
            where TException : Exception
        {
            try
            {
                action();
            }
            catch (TException)
            {
                return;
            }

            throw new InvalidOperationException(message);
        }

        private sealed class FakeUiAdapter : ICommandUiAdapter
        {
            public bool TryOpenContextMenu() => false;
            public bool? ShowDialog(Func<Window> createDialog) => null;
            public bool SetModelessWindowState(
                string commandId,
                ToggleIntent intent,
                Func<Window> createWindow) => false;
            public CommandDispatchResult ExecuteHostAction(
                CommandHostAction action,
                ToggleIntent intent) => CommandDispatchResult.Succeeded();
            public void ApplyToggleState(bool? isChecked) { }
            public void RefreshCommandStates() { }
            public void ShowError(string message) { }
        }

        private sealed class FakeAgendaLayoutStore : IAgendaLayoutStore
        {
            public List<AgendaLayoutConfig> Layouts { get; set; } = new List<AgendaLayoutConfig>();
            public bool SaveResult { get; set; } = true;
            public int SaveCalls { get; private set; }
            public List<AgendaLayoutConfig> SavedLayouts { get; private set; }

            public List<AgendaLayoutConfig> Load()
            {
                return new List<AgendaLayoutConfig>(Layouts);
            }

            public bool TrySave(List<AgendaLayoutConfig> layouts)
            {
                SaveCalls++;
                SavedLayouts = new List<AgendaLayoutConfig>(layouts);
                if (SaveResult) Layouts = new List<AgendaLayoutConfig>(layouts);
                return SaveResult;
            }
        }

        private sealed class FakeAgendaMutationHost : IAgendaMutationHost
        {
            public bool StartUndoResult { get; set; } = true;
            public bool UndoResult { get; set; }
            public int CaptureCalls { get; private set; }
            public int UndoCalls { get; private set; }
            public int FallbackCalls { get; private set; }

            public AgendaMutationSnapshot Capture()
            {
                CaptureCalls++;
                return new AgendaMutationSnapshot(new[] { 10, 20 }, "Existing Layout");
            }

            public bool TryStartUndoEntry()
            {
                return StartUndoResult;
            }

            public bool TryUndo(AgendaMutationSnapshot snapshot)
            {
                UndoCalls++;
                return UndoResult;
            }

            public void RunFallbackCleanup(AgendaMutationSnapshot snapshot)
            {
                FallbackCalls++;
            }
        }
    }
}
