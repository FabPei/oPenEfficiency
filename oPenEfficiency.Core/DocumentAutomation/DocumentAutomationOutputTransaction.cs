using System;
using System.IO;
using System.Collections.Concurrent;
using oPenEfficiency.Utils;

namespace oPenEfficiency.DocumentAutomation
{
    public interface IDocumentAutomationOutputOperations<TResult>
    {
        void CreateWorkingCopy(string temporaryPath);
        void ApplyAndSaveWorkingCopy(string temporaryPath);
        TResult OpenCompletedOutput(string outputPath);
    }

    public interface IDocumentAutomationFileSystem
    {
        bool DirectoryExists(string path);
        bool FileExists(string path);
        void Replace(string sourcePath, string destinationPath);
        void Move(string sourcePath, string destinationPath);
        void Delete(string path);
    }

    public sealed class PhysicalDocumentAutomationFileSystem : IDocumentAutomationFileSystem
    {
        public bool DirectoryExists(string path) => Directory.Exists(path);
        public bool FileExists(string path) => File.Exists(path);
        public void Replace(string sourcePath, string destinationPath) => File.Replace(sourcePath, destinationPath, null);
        public void Move(string sourcePath, string destinationPath) => File.Move(sourcePath, destinationPath);
        public void Delete(string path) => File.Delete(path);
    }

    public static class DocumentAutomationOutputPath
    {
        public static string CreateTemporaryPath(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory)) throw new ArgumentException("A temporary directory is required.", nameof(directory));
            return Path.Combine(directory, "oPE_DocumentAutomation_" + Guid.NewGuid().ToString("N") + ".pptx");
        }
    }

    internal static class DocumentAutomationOutputLease
    {
        private static readonly ConcurrentDictionary<string, byte> ActivePaths = new ConcurrentDictionary<string, byte>(StringComparer.OrdinalIgnoreCase);

        public static IDisposable Acquire(string outputPath)
        {
            if (!ActivePaths.TryAdd(outputPath, 0))
                throw new InvalidOperationException("Another document automation run is already writing to this output path.");
            return new Lease(outputPath);
        }

        private sealed class Lease : IDisposable
        {
            private string _path;
            public Lease(string path) { _path = path; }
            public void Dispose()
            {
                string path = _path;
                if (path == null) return;
                _path = null;
                ActivePaths.TryRemove(path, out byte _);
            }
        }
    }

    /// <summary>
    /// Publishes a fully processed working copy without exposing a partial output. The target is
    /// replaced only after copy creation and mutation both succeed.
    /// </summary>
    public sealed class DocumentAutomationOutputTransaction<TResult>
    {
        private readonly IDocumentAutomationFileSystem _fileSystem;
        private readonly Func<string, string> _temporaryPathFactory;
        private readonly Action<Exception, string> _cleanupFailureReporter;

        public DocumentAutomationOutputTransaction(
            IDocumentAutomationFileSystem fileSystem = null,
            Func<string, string> temporaryPathFactory = null,
            Action<Exception, string> cleanupFailureReporter = null)
        {
            _fileSystem = fileSystem ?? new PhysicalDocumentAutomationFileSystem();
            _temporaryPathFactory = temporaryPathFactory ?? DocumentAutomationOutputPath.CreateTemporaryPath;
            _cleanupFailureReporter = cleanupFailureReporter ?? ExceptionLogger.Log;
        }

        public TResult Execute(string outputPath, IDocumentAutomationOutputOperations<TResult> operations)
        {
            if (operations == null) throw new ArgumentNullException(nameof(operations));
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));

            outputPath = Path.GetFullPath(outputPath);
            string directory = Path.GetDirectoryName(outputPath);
            if (string.IsNullOrEmpty(directory) || !_fileSystem.DirectoryExists(directory))
                throw new DirectoryNotFoundException("The output folder does not exist.");

            using (DocumentAutomationOutputLease.Acquire(outputPath))
            {
                string temporaryPath = _temporaryPathFactory(directory);
                if (string.IsNullOrWhiteSpace(temporaryPath)) throw new InvalidOperationException("The temporary output path is empty.");
                temporaryPath = Path.GetFullPath(temporaryPath);
                if (string.Equals(temporaryPath, outputPath, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The working copy path must differ from the output path.");

                bool published = false;
                try
                {
                    operations.CreateWorkingCopy(temporaryPath);
                    operations.ApplyAndSaveWorkingCopy(temporaryPath);
                    if (_fileSystem.FileExists(outputPath)) _fileSystem.Replace(temporaryPath, outputPath);
                    else _fileSystem.Move(temporaryPath, outputPath);
                    published = true;
                    return operations.OpenCompletedOutput(outputPath);
                }
                finally
                {
                    if (!published && _fileSystem.FileExists(temporaryPath))
                    {
                        try { _fileSystem.Delete(temporaryPath); }
                        catch (Exception ex) { _cleanupFailureReporter(ex, "DocumentAutomationOutputTransaction.DeleteTemporaryCopy"); }
                    }
                }
            }
        }
    }

}
