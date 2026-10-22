using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.DocumentAutomation;
using oPenEfficiency.Services;

namespace oPenEfficiency.Core.Tests
{
    [TestClass]
    public sealed class DocumentAutomationOutputTransactionTests
    {
        private const string OutputPath = @"C:\automation\result.pptx";
        private const string TemporaryPath = @"C:\automation\working.pptx";

        [TestMethod]
        public void NewOutputPublishesOnlyAfterCopyWasCompleted()
        {
            var fixture = new TransactionFixture();

            string result = fixture.Execute();

            Assert.AreEqual(OutputPath, result);
            CollectionAssert.AreEqual(new[] { "create", "apply", "move", "open" }, fixture.Events);
            Assert.AreEqual("complete", fixture.FileSystem.Read(OutputPath));
            Assert.IsFalse(fixture.FileSystem.FileExists(TemporaryPath));
        }

        [TestMethod]
        public void ExistingOutputIsAtomicallyReplacedAfterCompletion()
        {
            var fixture = new TransactionFixture();
            fixture.FileSystem.Write(OutputPath, "original");

            fixture.Execute();

            CollectionAssert.AreEqual(new[] { "create", "apply", "replace", "open" }, fixture.Events);
            Assert.AreEqual("complete", fixture.FileSystem.Read(OutputPath));
        }

        [TestMethod]
        public void CopyFailurePreservesExistingTargetAndRemovesPartialWorkingCopy()
        {
            var fixture = new TransactionFixture { Failure = FailureStage.Create };
            fixture.FileSystem.Write(OutputPath, "original");

            InvalidOperationException error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Execute());

            Assert.AreEqual("create failed", error.Message);
            Assert.AreEqual("original", fixture.FileSystem.Read(OutputPath));
            Assert.IsFalse(fixture.FileSystem.FileExists(TemporaryPath));
            CollectionAssert.AreEqual(new[] { "create", "delete" }, fixture.Events);
        }

        [TestMethod]
        public void ApplyOrSaveFailurePreservesExistingTargetAndRemovesWorkingCopy()
        {
            var fixture = new TransactionFixture { Failure = FailureStage.Apply };
            fixture.FileSystem.Write(OutputPath, "original");

            InvalidOperationException error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Execute());

            Assert.AreEqual("apply failed", error.Message);
            Assert.AreEqual("original", fixture.FileSystem.Read(OutputPath));
            Assert.IsFalse(fixture.FileSystem.FileExists(TemporaryPath));
            CollectionAssert.AreEqual(new[] { "create", "apply", "delete" }, fixture.Events);
        }

        [TestMethod]
        public void PublishFailurePreservesExistingTargetAndRemovesCompletedWorkingCopy()
        {
            var fixture = new TransactionFixture { Failure = FailureStage.Publish };
            fixture.FileSystem.Write(OutputPath, "original");

            IOException error = Assert.ThrowsException<IOException>(() => fixture.Execute());

            Assert.AreEqual("publish failed", error.Message);
            Assert.AreEqual("original", fixture.FileSystem.Read(OutputPath));
            Assert.IsFalse(fixture.FileSystem.FileExists(TemporaryPath));
            CollectionAssert.AreEqual(new[] { "create", "apply", "replace", "delete" }, fixture.Events);
        }

        [TestMethod]
        public void OpenFailureLeavesACompletedPublishedOutputButReportsFailure()
        {
            var fixture = new TransactionFixture { Failure = FailureStage.Open };
            fixture.FileSystem.Write(OutputPath, "original");

            InvalidOperationException error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Execute());

            Assert.AreEqual("open failed", error.Message);
            Assert.AreEqual("complete", fixture.FileSystem.Read(OutputPath));
            Assert.IsFalse(fixture.FileSystem.FileExists(TemporaryPath));
            CollectionAssert.AreEqual(new[] { "create", "apply", "replace", "open" }, fixture.Events);
        }

        [TestMethod]
        public void CleanupFailureDoesNotMaskPrimaryFailureAndIsReported()
        {
            var fixture = new TransactionFixture { Failure = FailureStage.Apply };
            fixture.FileSystem.ThrowOnDelete = true;

            InvalidOperationException error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Execute());

            Assert.AreEqual("apply failed", error.Message);
            Assert.AreEqual(1, fixture.CleanupFailures.Count);
            Assert.AreEqual("DocumentAutomationOutputTransaction.DeleteTemporaryCopy", fixture.CleanupFailures[0].Context);
            Assert.IsInstanceOfType(fixture.CleanupFailures[0].Exception, typeof(IOException));
        }

        [TestMethod]
        public void MissingOutputDirectoryFailsBeforeCreatingACopy()
        {
            var fixture = new TransactionFixture();
            fixture.FileSystem.OutputDirectoryExists = false;

            Assert.ThrowsException<DirectoryNotFoundException>(() => fixture.Execute());

            Assert.AreEqual(0, fixture.Events.Count);
        }

        [TestMethod]
        public void WorkingCopyCannotAliasOutputPath()
        {
            var fileSystem = new FakeFileSystem(new List<string>());
            var operations = new FakeOperations(fileSystem, new List<string>());
            var transaction = new DocumentAutomationOutputTransaction<string>(fileSystem, directory => OutputPath);

            Assert.ThrowsException<InvalidOperationException>(() => transaction.Execute(OutputPath, operations));
            Assert.AreEqual(0, operations.Events.Count);
        }

        [TestMethod]
        public void StoreWriteFailureNeverDeletesExistingValues()
        {
            var events = new List<string>();

            Assert.ThrowsException<InvalidOperationException>(() =>
                AtomicStoreTransaction.AddBeforeDelete(
                    () => { events.Add("add"); throw new InvalidOperationException("write failed"); },
                    new Action[] { () => events.Add("delete-old") }));

            CollectionAssert.AreEqual(new[] { "add" }, events);
        }

        [TestMethod]
        public void StoreDeletesOldValuesOnlyAfterNewValueWasWritten()
        {
            var events = new List<string>();

            AtomicStoreTransaction.AddBeforeDelete(
                () => events.Add("add"),
                new Action[] { () => events.Add("delete-old-1"), () => events.Add("delete-old-2") });

            CollectionAssert.AreEqual(new[] { "add", "delete-old-1", "delete-old-2" }, events);
        }

        [TestMethod]
        public void StoreDeleteFailureLeavesTheNewValueAvailableAndStopsFurtherDeletion()
        {
            var events = new List<string>();

            Assert.ThrowsException<IOException>(() =>
                AtomicStoreTransaction.AddBeforeDelete(
                    () => events.Add("add"),
                    new Action[]
                    {
                        () => { events.Add("delete-old-1"); throw new IOException("delete failed"); },
                        () => events.Add("delete-old-2")
                    }));

            CollectionAssert.AreEqual(new[] { "add", "delete-old-1" }, events);
        }

        [TestMethod]
        public void ConcurrentRunsForTheSameOutputFailFastWithoutTouchingTheSecondCopy()
        {
            var firstEvents = new List<string>();
            var firstFileSystem = new FakeFileSystem(firstEvents);
            var entered = new ManualResetEventSlim();
            var release = new ManualResetEventSlim();
            var firstOperations = new BlockingOperations(firstFileSystem, entered, release);
            var firstTransaction = new DocumentAutomationOutputTransaction<string>(firstFileSystem, directory => @"C:\automation\first-working.pptx");
            Task<string> first = Task.Run(() => firstTransaction.Execute(OutputPath, firstOperations));
            Assert.IsTrue(entered.Wait(TimeSpan.FromSeconds(5)), "The first transaction did not acquire the output lease.");

            var secondEvents = new List<string>();
            var secondFileSystem = new FakeFileSystem(secondEvents);
            var secondOperations = new FakeOperations(secondFileSystem, secondEvents);
            var secondTransaction = new DocumentAutomationOutputTransaction<string>(secondFileSystem, directory => @"C:\automation\second-working.pptx");
            InvalidOperationException error = Assert.ThrowsException<InvalidOperationException>(() => secondTransaction.Execute(OutputPath, secondOperations));

            Assert.IsTrue(error.Message.Contains("already writing"));
            Assert.AreEqual(0, secondEvents.Count);
            release.Set();
            Assert.AreEqual(OutputPath, first.GetAwaiter().GetResult());
        }

        [TestMethod]
        public void ConcurrentRunsForDifferentOutputsCanProceedIndependently()
        {
            var enteredOne = new ManualResetEventSlim();
            var enteredTwo = new ManualResetEventSlim();
            var release = new ManualResetEventSlim();
            var fileSystemOne = new FakeFileSystem(new List<string>());
            var fileSystemTwo = new FakeFileSystem(new List<string>());
            var transactionOne = new DocumentAutomationOutputTransaction<string>(fileSystemOne, directory => @"C:\automation\one-working.pptx");
            var transactionTwo = new DocumentAutomationOutputTransaction<string>(fileSystemTwo, directory => @"C:\automation\two-working.pptx");
            Task<string> first = Task.Run(() => transactionOne.Execute(OutputPath, new BlockingOperations(fileSystemOne, enteredOne, release)));
            Task<string> second = Task.Run(() => transactionTwo.Execute(@"C:\automation\second-result.pptx", new BlockingOperations(fileSystemTwo, enteredTwo, release)));

            Assert.IsTrue(enteredOne.Wait(TimeSpan.FromSeconds(5)) && enteredTwo.Wait(TimeSpan.FromSeconds(5)), "Different targets were globally serialized.");
            release.Set();
            Task.WaitAll(first, second);
        }

        [TestMethod]
        public void FailedRunReleasesItsOutputLeaseForRetry()
        {
            var failed = new TransactionFixture { Failure = FailureStage.Apply };
            Assert.ThrowsException<InvalidOperationException>(() => failed.Execute());
            var retry = new TransactionFixture();
            Assert.AreEqual(OutputPath, retry.Execute());
        }

        [TestMethod]
        public void DefaultTemporaryPathsAreUniqueAndStayInTheRequestedDirectory()
        {
            string[] paths = Enumerable.Range(0, 1000).Select(index => DocumentAutomationOutputPath.CreateTemporaryPath(@"C:\automation")).ToArray();
            Assert.AreEqual(paths.Length, paths.Distinct(StringComparer.OrdinalIgnoreCase).Count());
            Assert.IsTrue(paths.All(path => string.Equals(Path.GetDirectoryName(path), @"C:\automation", StringComparison.OrdinalIgnoreCase)));
            Assert.IsTrue(paths.All(path => Path.GetFileName(path).StartsWith("oPE_DocumentAutomation_") && path.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase)));
        }

        private enum FailureStage { None, Create, Apply, Publish, Open }

        private sealed class TransactionFixture
        {
            public readonly List<string> Events = new List<string>();
            public readonly List<CleanupFailure> CleanupFailures = new List<CleanupFailure>();
            public readonly FakeFileSystem FileSystem;
            public FailureStage Failure { get; set; }

            public TransactionFixture()
            {
                FileSystem = new FakeFileSystem(Events);
            }

            public string Execute()
            {
                FileSystem.ThrowOnPublish = Failure == FailureStage.Publish;
                var operations = new FakeOperations(FileSystem, Events) { Failure = Failure };
                var transaction = new DocumentAutomationOutputTransaction<string>(
                    FileSystem,
                    directory => TemporaryPath,
                    (exception, context) => CleanupFailures.Add(new CleanupFailure { Exception = exception, Context = context }));
                return transaction.Execute(OutputPath, operations);
            }
        }

        private sealed class FakeOperations : IDocumentAutomationOutputOperations<string>
        {
            private readonly FakeFileSystem _fileSystem;
            public readonly List<string> Events;
            public FailureStage Failure { get; set; }

            public FakeOperations(FakeFileSystem fileSystem, List<string> events)
            {
                _fileSystem = fileSystem;
                Events = events;
            }

            public void CreateWorkingCopy(string temporaryPath)
            {
                Events.Add("create");
                _fileSystem.Write(temporaryPath, "partial-copy");
                if (Failure == FailureStage.Create) throw new InvalidOperationException("create failed");
            }

            public void ApplyAndSaveWorkingCopy(string temporaryPath)
            {
                Events.Add("apply");
                _fileSystem.Write(temporaryPath, "partial-automation");
                if (Failure == FailureStage.Apply) throw new InvalidOperationException("apply failed");
                _fileSystem.Write(temporaryPath, "complete");
            }

            public string OpenCompletedOutput(string outputPath)
            {
                Events.Add("open");
                if (Failure == FailureStage.Open) throw new InvalidOperationException("open failed");
                return outputPath;
            }
        }

        private sealed class BlockingOperations : IDocumentAutomationOutputOperations<string>
        {
            private readonly FakeFileSystem _fileSystem;
            private readonly ManualResetEventSlim _entered;
            private readonly ManualResetEventSlim _release;

            public BlockingOperations(FakeFileSystem fileSystem, ManualResetEventSlim entered, ManualResetEventSlim release)
            {
                _fileSystem = fileSystem;
                _entered = entered;
                _release = release;
            }

            public void CreateWorkingCopy(string temporaryPath)
            {
                _fileSystem.Write(temporaryPath, "copy");
                _entered.Set();
                if (!_release.Wait(TimeSpan.FromSeconds(10))) throw new TimeoutException("The concurrency test was not released.");
            }

            public void ApplyAndSaveWorkingCopy(string temporaryPath) => _fileSystem.Write(temporaryPath, "complete");
            public string OpenCompletedOutput(string outputPath) => outputPath;
        }

        private sealed class FakeFileSystem : IDocumentAutomationFileSystem
        {
            private readonly Dictionary<string, string> _files = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private readonly List<string> _events;
            public bool OutputDirectoryExists { get; set; } = true;
            public bool ThrowOnPublish { get; set; }
            public bool ThrowOnDelete { get; set; }

            public FakeFileSystem(List<string> events) { _events = events; }
            public bool DirectoryExists(string path) => OutputDirectoryExists;
            public bool FileExists(string path) => _files.ContainsKey(path);
            public string Read(string path) => _files[path];
            public void Write(string path, string content) => _files[path] = content;

            public void Replace(string sourcePath, string destinationPath)
            {
                _events.Add("replace");
                Publish(sourcePath, destinationPath);
            }

            public void Move(string sourcePath, string destinationPath)
            {
                _events.Add("move");
                Publish(sourcePath, destinationPath);
            }

            public void Delete(string path)
            {
                _events.Add("delete");
                if (ThrowOnDelete) throw new IOException("cleanup failed");
                _files.Remove(path);
            }

            private void Publish(string sourcePath, string destinationPath)
            {
                if (ThrowOnPublish) throw new IOException("publish failed");
                _files[destinationPath] = _files[sourcePath];
                _files.Remove(sourcePath);
            }
        }

        private sealed class CleanupFailure
        {
            public Exception Exception { get; set; }
            public string Context { get; set; }
        }
    }
}
