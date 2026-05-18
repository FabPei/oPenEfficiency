using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using oPenEfficiency;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Central service for managing asset cache (thumbnails, previews) in a user-specific location.
    /// Cache is stored in %APPDATA%\oPenEfficiency\cache\ using SHA256 hash of source file path.
    /// For shared folders (OneDrive, network shares), cache can be stored locally next to source files.
    /// </summary>
    public class AssetCacheService
    {
        private readonly string _cacheBasePath;
        private const string SourcePathFile = "source.path";
        private const string TimestampFile = "cache.timestamp";
        private const string SharedFoldersConfig = "shared_cache_folders.json";

        public AssetCacheService()
        {
            // %APPDATA%\oPenEfficiency\cache\
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _cacheBasePath = Path.Combine(appDataPath, "oPenEfficiency", "cache");

            // Ensure base cache directory exists
            if (!Directory.Exists(_cacheBasePath))
            {
                Directory.CreateDirectory(_cacheBasePath);
            }
        }

        /// <summary>
        /// Gets the list of shared folders where local caching is enabled.
        /// </summary>
        public List<string> GetSharedFolders()
        {
            var list = JsonService.Load<List<string>>(SharedFoldersConfig);
            return list ?? new List<string>();
        }

        /// <summary>
        /// Adds a shared folder to the configuration.
        /// </summary>
        public void AddSharedFolder(string path)
        {
            var list = GetSharedFolders();
            if (!list.Contains(path, StringComparer.OrdinalIgnoreCase))
            {
                list.Add(path);
                JsonService.Save(list, SharedFoldersConfig);
            }
        }

        /// <summary>
        /// Removes a shared folder from the configuration.
        /// </summary>
        public void RemoveSharedFolder(string path)
        {
            var list = GetSharedFolders();
            if (list.Remove(path))
            {
                JsonService.Save(list, SharedFoldersConfig);
            }
        }

        /// <summary>
        /// Checks if a file path is within any of the configured shared folders.
        /// </summary>
        public bool IsInSharedFolder(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return false;

            var sharedFolders = GetSharedFolders();
            string normalizedPath = Path.GetFullPath(filePath).ToLowerInvariant();

            foreach (var sharedFolder in sharedFolders)
            {
                string normalizedShared = Path.GetFullPath(sharedFolder).ToLowerInvariant();
                if (normalizedPath.StartsWith(normalizedShared))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Gets the cache folder for a given source file path.
        /// For files in shared folders, cache is stored locally next to the source file.
        /// For other files, cache is stored in %APPDATA%\oPenEfficiency\cache\[hash]\.
        /// </summary>
        /// <param name="filePath">Full path to the source file</param>
        /// <returns>Cache folder path</returns>
        public string GetCacheFolder(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return null;

            // Check if file is in a shared folder - use local adjacent cache
            if (IsInSharedFolder(filePath))
            {
                string dir = Path.GetDirectoryName(filePath);
                string name = Path.GetFileName(filePath);
                // Use hash-based folder name even for local cache to avoid long path issues
                string localHash = GeneratePathHash(filePath);
                return Path.Combine(dir, ".oPenEff_cache", localHash);
            }

            // Default: central AppData cache
            string cacheHash = GeneratePathHash(filePath);
            return Path.Combine(_cacheBasePath, cacheHash);
        }

        /// <summary>
        /// Registers the source file path by storing it in the cache folder.
        /// This enables cache validation and stale cache cleanup.
        /// </summary>
        /// <param name="filePath">Full path to the source file</param>
        /// <param name="cacheFolder">Cache folder path (from GetCacheFolder)</param>
        public void RegisterSourcePath(string filePath, string cacheFolder)
        {
            if (string.IsNullOrEmpty(filePath) || string.IsNullOrEmpty(cacheFolder))
                return;

            try
            {
                Directory.CreateDirectory(cacheFolder);
                string sourcePathFile = Path.Combine(cacheFolder, SourcePathFile);
                File.WriteAllText(sourcePathFile, filePath, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AssetCacheService.RegisterSourcePath error: {ex.Message}");
            }
        }

        /// <summary>
        /// Retrieves the original source file path from a cache folder.
        /// </summary>
        /// <param name="cacheFolder">Cache folder path</param>
        /// <returns>Original source file path, or null if not found</returns>
        public string GetSourcePath(string cacheFolder)
        {
            if (string.IsNullOrEmpty(cacheFolder) || !Directory.Exists(cacheFolder))
                return null;

            try
            {
                string sourcePathFile = Path.Combine(cacheFolder, SourcePathFile);
                if (File.Exists(sourcePathFile))
                {
                    return File.ReadAllText(sourcePathFile, Encoding.UTF8).Trim();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AssetCacheService.GetSourcePath error: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Validates if the cache is still valid by comparing timestamps.
        /// </summary>
        /// <param name="filePath">Full path to the source file</param>
        /// <returns>True if cache exists and is up-to-date, false otherwise</returns>
        public bool IsCacheValid(string filePath)
        {
            try
            {
                string cacheDir = GetCacheFolder(filePath);
                if (!Directory.Exists(cacheDir))
                    return false;

                string tsFile = Path.Combine(cacheDir, TimestampFile);
                if (!File.Exists(tsFile))
                    return false;

                long cachedTicks = long.Parse(File.ReadAllText(tsFile).Trim());
                long currentTicks = File.GetLastWriteTimeUtc(filePath).Ticks;
                return cachedTicks == currentTicks;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AssetCacheService.IsCacheValid error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Writes the current timestamp to the cache folder.
        /// <summary>
        /// Writes the current timestamp to the cache folder.
        /// </summary>
        public void WriteTimestamp(string filePath)
        {
            try
            {
                string cacheDir = GetCacheFolder(filePath);
                Directory.CreateDirectory(cacheDir);

                string tsFile = Path.Combine(cacheDir, TimestampFile);
                File.WriteAllText(tsFile, File.GetLastWriteTimeUtc(filePath).Ticks.ToString());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AssetCacheService.WriteTimestamp error: {ex.Message}");
            }
        }

        /// <summary>
        /// Writes a special timestamp indicating the file was processed but some items might have failed.
        /// This prevents constant re-scanning of corrupted presentations while still marking them as 'scanned'.
        /// </summary>
        public void WritePartialTimestamp(string filePath)
        {
            try
            {
                string cacheDir = GetCacheFolder(filePath);
                Directory.CreateDirectory(cacheDir);

                string tsFile = Path.Combine(cacheDir, TimestampFile);
                // Mark with current ticks to stop immediate re-scans
                File.WriteAllText(tsFile, File.GetLastWriteTimeUtc(filePath).Ticks.ToString());

                // Add a marker file to indicate it was a partial/faulty run
                File.WriteAllText(Path.Combine(cacheDir, "cache.partial"), DateTime.UtcNow.ToString());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AssetCacheService.WritePartialTimestamp error: {ex.Message}");
            }
        }
        /// <summary>
        /// Cleans up stale cache entries where the source file no longer exists.
        /// Also removes legacy adjacent cache folders (.oPenEff_thumbs, .oPenEff_shape_thumbs).
        /// </summary>
        /// <returns>Cleanup report with counts of removed items</returns>
        public CacheCleanupReport CleanupStaleCache()
        {
            var report = new CacheCleanupReport();

            try
            {
                // Clean central cache (AppData)
                if (Directory.Exists(_cacheBasePath))
                {
                    foreach (var cacheDir in Directory.GetDirectories(_cacheBasePath))
                    {
                        try
                        {
                            string sourcePath = GetSourcePath(cacheDir);
                            if (string.IsNullOrEmpty(sourcePath) || !File.Exists(sourcePath))
                            {
                                // Source file doesn't exist or can't be determined - delete cache
                                Directory.Delete(cacheDir, true);
                                report.StaleCacheFoldersDeleted++;
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"AssetCacheService.CleanupStaleCache error deleting {cacheDir}: {ex.Message}");
                        }
                    }
                }

                // Clean local cache folders in shared folders
                foreach (var sharedFolder in GetSharedFolders())
                {
                    if (Directory.Exists(sharedFolder))
                    {
                        try
                        {
                            var localCacheFolders = Directory.GetDirectories(sharedFolder, ".oPenEff_cache", SearchOption.AllDirectories);
                            foreach (var cacheDir in localCacheFolders)
                            {
                                try
                                {
                                    string sourcePath = GetSourcePath(cacheDir);
                                    if (string.IsNullOrEmpty(sourcePath) || !File.Exists(sourcePath))
                                    {
                                        Directory.Delete(cacheDir, true);
                                        report.StaleCacheFoldersDeleted++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Error cleaning local cache {cacheDir}: {ex.Message}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error scanning shared folder {sharedFolder}: {ex.Message}");
                        }
                    }
                }

                // Clean legacy adjacent cache folders
                report.LegacyCacheFoldersDeleted = CleanupLegacyAdjacentCache();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AssetCacheService.CleanupStaleCache error: {ex.Message}");
            }

            return report;
        }

        /// <summary>
        /// Cleans up legacy adjacent cache folders (.oPenEff_thumbs and .oPenEff_shape_thumbs)
        /// from all registered library folders.
        /// </summary>
        /// <returns>Number of legacy cache folders deleted</returns>
        private int CleanupLegacyAdjacentCache()
        {
            int deletedCount = 0;
            const string SlideCacheName = ".oPenEff_thumbs";
            const string ShapeCacheName = ".oPenEff_shape_thumbs";

            try
            {
                // Get library folders from settings
                var settings = Properties.Settings.Default;
                if (settings.LibraryFolders == null)
                    return deletedCount;

                var folders = new System.Collections.Generic.List<string>();
                foreach (var path in settings.LibraryFolders)
                {
                    if (path is string s)
                        folders.Add(s);
                }

                // Scan each library folder for legacy cache directories
                foreach (var folder in folders)
                {
                    if (!Directory.Exists(folder))
                        continue;

                    try
                    {
                        // Check for legacy cache folders at all levels
                        var legacySlideFolders = Directory.GetDirectories(folder, SlideCacheName, SearchOption.AllDirectories);
                        foreach (var legacyDir in legacySlideFolders)
                        {
                            try
                            {
                                Directory.Delete(legacyDir, true);
                                deletedCount++;
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Failed to delete legacy cache {legacyDir}: {ex.Message}");
                            }
                        }

                        var legacyShapeFolders = Directory.GetDirectories(folder, ShapeCacheName, SearchOption.AllDirectories);
                        foreach (var legacyDir in legacyShapeFolders)
                        {
                            try
                            {
                                Directory.Delete(legacyDir, true);
                                deletedCount++;
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Failed to delete legacy cache {legacyDir}: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"AssetCacheService.CleanupLegacyAdjacentCache error scanning {folder}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AssetCacheService.CleanupLegacyAdjacentCache error: {ex.Message}");
            }

            return deletedCount;
        }

        /// <summary>
        /// Generates a SHA256 hash of the file path for cache folder naming.
        /// Uses first 16 characters which provides sufficient uniqueness.
        /// </summary>
        private string GeneratePathHash(string fullPath)
        {
            using (var sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(fullPath));
                string hash = BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                return hash.Substring(0, 16);
            }
        }

        /// <summary>
        /// Gets the total size of the cache directory in bytes.
        /// </summary>
        public long GetCacheSizeBytes()
        {
            try
            {
                if (!Directory.Exists(_cacheBasePath))
                    return 0;

                long totalSize = 0;
                foreach (var file in Directory.EnumerateFiles(_cacheBasePath, "*.*", SearchOption.AllDirectories))
                {
                    try
                    {
                        var info = new FileInfo(file);
                        totalSize += info.Length;
                    }
                    catch
                    {
                        // File might be in use or inaccessible
                    }
                }
                return totalSize;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AssetCacheService.GetCacheSizeBytes error: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// Gets a human-readable cache size string (e.g., "15.3 MB").
        /// </summary>
        public string GetCacheSizeFormatted()
        {
            long bytes = GetCacheSizeBytes();

            if (bytes < 1024)
                return $"{bytes} B";
            if (bytes < 1024 * 1024)
                return $"{bytes / 1024.0:F1} KB";
            if (bytes < 1024 * 1024 * 1024)
                return $"{bytes / (1024.0 * 1024.0):F1} MB";
            return $"{bytes / (1024.0 * 1024.0 * 1024.0):F2} GB";
        }

        /// <summary>
        /// Gets the total cache size including local caches in shared folders.
        /// </summary>
        public string GetTotalCacheSizeFormatted()
        {
            long totalBytes = GetCacheSizeBytes();

            // Add sizes of local caches in shared folders
            foreach (var sharedFolder in GetSharedFolders())
            {
                if (Directory.Exists(sharedFolder))
                {
                    try
                    {
                        var cacheFolders = Directory.GetDirectories(sharedFolder, ".oPenEff_cache", SearchOption.AllDirectories);
                        foreach (var cacheDir in cacheFolders)
                        {
                            foreach (var file in Directory.EnumerateFiles(cacheDir, "*.*", SearchOption.AllDirectories))
                            {
                                try
                                {
                                    var info = new FileInfo(file);
                                    totalBytes += info.Length;
                                }
                                catch { }
                            }
                        }
                    }
                    catch { }
                }
            }

            if (totalBytes < 1024)
                return $"{totalBytes} B";
            if (totalBytes < 1024 * 1024)
                return $"{totalBytes / 1024.0:F1} KB";
            if (totalBytes < 1024 * 1024 * 1024)
                return $"{totalBytes / (1024.0 * 1024.0):F1} MB";
            return $"{totalBytes / (1024.0 * 1024.0 * 1024.0):F2} GB";
        }
    }

    /// <summary>
    /// Report from cache cleanup operation.
    /// </summary>
    public class CacheCleanupReport
    {
        public int StaleCacheFoldersDeleted { get; set; }
        public int LegacyCacheFoldersDeleted { get; set; }
        public int TotalDeleted => StaleCacheFoldersDeleted + LegacyCacheFoldersDeleted;
    }
}
