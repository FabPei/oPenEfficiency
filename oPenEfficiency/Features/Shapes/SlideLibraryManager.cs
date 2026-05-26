using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Models;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Slide = Microsoft.Office.Interop.PowerPoint.Slide;

namespace oPenEfficiency.Features
{
    public class SlideLibraryManager
    {
        private readonly PowerPointManager _manager;
        private readonly oPenEfficiency.Properties.Settings _settings;
        private readonly AssetCacheService _cacheService;

        public SlideLibraryManager(PowerPointManager manager)
        {
            _manager = manager;
            _settings = oPenEfficiency.Properties.Settings.Default;
            _cacheService = new AssetCacheService();
        }

        // --- Step 1: Library Source Management ---

        public List<string> GetLibraryFolders()
        {
            try
            {
                if (_settings.LibraryFolders == null)
                {
                    _settings.LibraryFolders = new System.Collections.Specialized.StringCollection();
                    _settings.Save();
                }
                
                var list = new List<string>();
                
                // Always add the user's personal "Favorites" folder
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string myAssetsPath = Path.Combine(appDataPath, "oPenEfficiency", "favorites", "slides");
                if (!Directory.Exists(myAssetsPath)) Directory.CreateDirectory(myAssetsPath);
                list.Add(myAssetsPath);

                foreach (var path in _settings.LibraryFolders)
                {
                    if (path is string s && !list.Contains(s, StringComparer.OrdinalIgnoreCase)) 
                        list.Add(s);
                }
                return list;
            }
            catch
            {
                return new List<string>();
            }
        }

        public void AddLibraryFolder(string path)
        {
            if (Directory.Exists(path) && !_settings.LibraryFolders.Contains(path))
            {
                _settings.LibraryFolders.Add(path);
                _settings.Save();
            }
        }

        public void RemoveLibraryFolder(string path)
        {
            if (_settings.LibraryFolders.Contains(path))
            {
                _settings.LibraryFolders.Remove(path);
                _settings.Save();
            }
        }

        public async Task<List<PresentationFile>> ScanLibrariesAsync()
        {
            return await Task.Run(() =>
            {
                var scannedFiles = new List<PresentationFile>();
                var folders = GetLibraryFolders();

                foreach (var folder in folders)
                {
                    if (Directory.Exists(folder))
                    {
                        try
                        {
                            var files = oPenEfficiency.Utils.CommonUtils.SafeGetFiles(folder, "*.pptx");
                            foreach (var file in files)
                            {
                                string fullPath = Path.GetFullPath(file);
                                if (scannedFiles.Any(f => f.FilePath.Equals(fullPath, StringComparison.OrdinalIgnoreCase))) continue;

                                // Compute the relative subfolder path
                                string fileDir = Path.GetDirectoryName(file);
                                string relativeSub = "";
                                if (fileDir.Length > folder.Length)
                                    relativeSub = fileDir.Substring(folder.Length).TrimStart(Path.DirectorySeparatorChar);

                                scannedFiles.Add(new PresentationFile
                                {
                                    FilePath = fullPath, 
                                    FileName = Path.GetFileNameWithoutExtension(file),
                                    FolderName = new DirectoryInfo(folder).Name,
                                    RootFolder = folder,
                                    SubFolder = relativeSub
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"SlideLibraryManager.ScanLibrariesAsync error: {ex.Message}");
                        }
                    }
                }
                return scannedFiles;
            });
        }

        /// <summary>
        /// Builds a hierarchical tree of FolderTreeNodes from scanned presentation files.
        /// </summary>
        public List<FolderTreeNode> BuildFolderTree(List<PresentationFile> files, bool foldersOnly = false)
        {
            var roots = new List<FolderTreeNode>();

            // Group by root library folder
            var byRoot = files.GroupBy(f => f.RootFolder);

            foreach (var rootGroup in byRoot)
            {
                string rootPath = rootGroup.Key;
                string rootName = new DirectoryInfo(rootPath).Name;
                var rootNode = new FolderTreeNode
                {
                    Name = rootName,
                    FullPath = rootPath,
                    FileCount = rootGroup.Count(f => string.IsNullOrEmpty(f.SubFolder))
                };

                // Group subfolder paths and build child nodes
                var withSub = rootGroup.Where(f => !string.IsNullOrEmpty(f.SubFolder)).ToList();
                BuildChildNodes(rootNode, withSub, rootPath, foldersOnly);

                if (!foldersOnly)
                {
                    // Add files directly in the root folder as leaf nodes
                    var directFiles = rootGroup.Where(f => string.IsNullOrEmpty(f.SubFolder)).ToList();
                    foreach (var file in directFiles)
                    {
                        rootNode.Children.Add(new FolderTreeNode
                        {
                            Name = file.FileName,
                            FullPath = file.FilePath,
                            IsFile = true,
                            IsCached = IsCacheValid(file.FilePath),
                            FilePath = file.FilePath
                        });
                    }
                }

                // Update total count (self + all descendants)
                rootNode.TotalFileCount = rootGroup.Count();
                roots.Add(rootNode);
            }

            return roots;
        }

        private void BuildChildNodes(FolderTreeNode parent, List<PresentationFile> files, string rootPath, bool foldersOnly = false)
        {
            // Group by the first path segment relative to the parent
            var groups = files.GroupBy(f =>
            {
                string rel = f.SubFolder;
                // If parent is not the root, strip the parent's relative path
                string parentRel = "";
                if (parent.FullPath.Length > rootPath.Length)
                    parentRel = parent.FullPath.Substring(rootPath.Length).TrimStart(Path.DirectorySeparatorChar);

                if (!string.IsNullOrEmpty(parentRel) && rel.StartsWith(parentRel + Path.DirectorySeparatorChar))
                    rel = rel.Substring(parentRel.Length + 1);
                else if (!string.IsNullOrEmpty(parentRel) && rel == parentRel)
                    return ""; // file is directly in parent

                int sep = rel.IndexOf(Path.DirectorySeparatorChar);
                return sep >= 0 ? rel.Substring(0, sep) : rel;
            });

            foreach (var g in groups)
            {
                if (string.IsNullOrEmpty(g.Key)) continue; // files directly in parent

                string childPath = Path.Combine(parent.FullPath, g.Key);
                var child = new FolderTreeNode
                {
                    Name = g.Key,
                    FullPath = childPath,
                    FileCount = g.Count(f => Path.GetDirectoryName(f.FilePath).Equals(childPath, StringComparison.OrdinalIgnoreCase)),
                    TotalFileCount = g.Count()
                };

                // Recurse for deeper levels
                var deeper = g.Where(f => !Path.GetDirectoryName(f.FilePath).Equals(childPath, StringComparison.OrdinalIgnoreCase)).ToList();
                if (deeper.Count > 0)
                    BuildChildNodes(child, deeper, rootPath, foldersOnly);

                if (!foldersOnly)
                {
                    // Add files directly in this subfolder as leaf nodes
                    var directInChild = g.Where(f => Path.GetDirectoryName(f.FilePath).Equals(childPath, StringComparison.OrdinalIgnoreCase)).ToList();
                    foreach (var file in directInChild)
                    {
                        child.Children.Add(new FolderTreeNode
                        {
                            Name = file.FileName,
                            FullPath = file.FilePath,
                            IsFile = true,
                            IsCached = IsCacheValid(file.FilePath),
                            FilePath = file.FilePath
                        });
                    }
                }

                parent.Children.Add(child);
            }
        }

        // --- Step 2: Thumbnail Generation ---
        
        public List<SlideTemplate> GetSlidesFromFile(string filePath)
        {
            var slides = new List<SlideTemplate>();
            try
            {
                int count = GetSlideCount(filePath);
                string fileName = Path.GetFileNameWithoutExtension(filePath);

                for (int i = 1; i <= count; i++)
                {
                    var template = new oPenEfficiency.Models.SlideTemplate
                    {
                        Title = count > 1 ? $"{fileName} - Slide {i}" : fileName,
                        FilePath = filePath,
                        SlideIndex = i,
                        PresentationSlideCount = count
                    };

                    // Load thumbnail if cached
                    if (IsCacheValid(filePath))
                    {
                        template.ThumbnailSource = LoadCachedThumbnail(filePath, i);
                    }
                    else if (i == 1)
                    {
                        // Fallback: extract the internal thumbnail for the first slide
                        template.ThumbnailSource = GetThumbnail(filePath);
                    }

                    slides.Add(template);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SlideLibraryManager.GetSlidesFromFile error: {ex.Message}");
            }
            return slides;
        }

        public async Task GenerateAllThumbnails()
        {
            try
            {
                var files = await ScanLibrariesAsync();
                foreach (var file in files)
                {
                    if (!IsCacheValid(file.FilePath))
                    {
                        // Run on UI thread but yield slightly to not completely freeze the app
                        await Task.Delay(100);
                        await GenerateThumbnailCache(file.FilePath);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SlideLibraryManager.GenerateAllThumbnails error: {ex.Message}");
            }
        }

        public BitmapImage GetThumbnail(string filePath)
        {
            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var archive = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Read))
                {
                    var thumbEntry = archive.GetEntry("docProps/thumbnail.jpeg");
                    if (thumbEntry != null)
                    {
                        using (var stream = thumbEntry.Open())
                        using (var ms = new MemoryStream())
                        {
                            stream.CopyTo(ms);
                            ms.Position = 0;
                            var image = new BitmapImage();
                            image.BeginInit();
                            image.CacheOption = BitmapCacheOption.OnLoad;
                            image.StreamSource = ms;
                            image.EndInit();
                            image.Freeze();
                            return image;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SlideLibraryManager.GetThumbnail error: {ex.Message}");
            }
            return null;
        }

        public async Task<BitmapImage> GetThumbnailAsync(string filePath)
        {
            return await Task.Run(() =>
            {
                try
                {
                    var fileInfo = new FileInfo(filePath);
                    if (!fileInfo.Exists) return null;

                    string cacheFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "oPenEfficiency", "thumbnails");
                    if (!Directory.Exists(cacheFolder)) Directory.CreateDirectory(cacheFolder);

                    // Create a unique cache key based on file path and modification time
                    string hashInput = filePath.ToLowerInvariant() + "_" + fileInfo.LastWriteTimeUtc.Ticks.ToString();
                    string hashStr;
                    using (var sha = System.Security.Cryptography.SHA256.Create())
                    {
                        var bytes = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(hashInput));
                        var sb = new System.Text.StringBuilder();
                        foreach (byte b in bytes) sb.Append(b.ToString("x2"));
                        hashStr = sb.ToString();
                    }

                    string cacheFile = Path.Combine(cacheFolder, hashStr + ".jpg");

                    // 1. Try to load from disk cache
                    if (File.Exists(cacheFile))
                    {
                        try
                        {
                            var cachedImg = new BitmapImage();
                            cachedImg.BeginInit();
                            cachedImg.CacheOption = BitmapCacheOption.OnLoad;
                            cachedImg.UriSource = new Uri(cacheFile);
                            cachedImg.DecodePixelWidth = 200; // Downscale cached image to save RAM
                            cachedImg.EndInit();
                            cachedImg.Freeze();
                            return cachedImg;
                        }
                        catch { /* Ignore cache corruption, fall back to generation */ }
                    }

                    // 2. Generate thumbnail from PPTX
                    using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var archive = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Read))
                    {
                        var thumbEntry = archive.GetEntry("docProps/thumbnail.jpeg");
                        if (thumbEntry != null)
                        {
                            using (var stream = thumbEntry.Open())
                            using (var ms = new MemoryStream())
                            {
                                stream.CopyTo(ms);
                                ms.Position = 0;
                                var image = new BitmapImage();
                                image.BeginInit();
                                image.CacheOption = BitmapCacheOption.OnLoad;
                                image.StreamSource = ms;
                                image.DecodePixelWidth = 200; // Important: Downscale the massive raw thumbnail
                                image.EndInit();
                                image.Freeze();

                                // 3. Save to disk cache for next time
                                try
                                {
                                    var encoder = new System.Windows.Media.Imaging.JpegBitmapEncoder();
                                    encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(image));
                                    using (var cacheStream = new FileStream(cacheFile, FileMode.Create, FileAccess.Write))
                                    {
                                        encoder.Save(cacheStream);
                                    }
                                }
                                catch { /* Ignore cache write errors */ }

                                return image;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"SlideLibraryManager.GetThumbnailAsync error: {ex.Message}");
                }
                return null;
            });
        }

        /// <summary>
        /// Extracts thumbnails for every slide in a PPTX.
        /// pptx slide images are at ppt/slides/thumbnails/thumbnailN.jpeg (if exported) or
        /// we fall back to the first-slide (docProps/thumbnail) for slide 1, and null for the rest.
        /// </summary>
        public async Task<List<BitmapImage>> GetSlideThumbnailsAsync(string filePath, int slideCount)
        {
            return await Task.Run(() =>
            {
                var results = new List<BitmapImage>();
                try
                {
                    using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var archive = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Read))
                    {
                        for (int i = 1; i <= slideCount; i++)
                        {
                            BitmapImage img = null;

                            // Newer exports may include per-slide previews at ppt/slides/thumbnails/slide{i}.jpg
                            var entry = archive.GetEntry($"ppt/slides/thumbnails/slide{i}.jpg") ??
                                        archive.GetEntry($"ppt/slides/thumbnails/slide{i}.jpeg");

                            // Fallback to docProps thumbnail for slide 1
                            if (entry == null && i == 1)
                                entry = archive.GetEntry("docProps/thumbnail.jpeg");

                            if (entry != null)
                            {
                                try
                                {
                                    using (var stream = entry.Open())
                                    using (var ms = new MemoryStream())
                                    {
                                        stream.CopyTo(ms);
                                        ms.Position = 0;
                                        img = new BitmapImage();
                                        img.BeginInit();
                                        img.CacheOption = BitmapCacheOption.OnLoad;
                                        img.StreamSource = ms;
                                        img.EndInit();
                                        img.Freeze();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"SlideLibraryManager.GetSlideThumbnailsAsync.extract error: {ex.Message}");
                                    img = null;
                                }
                            }

                            results.Add(img); // null if no thumbnail available
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"SlideLibraryManager.GetSlideThumbnailsAsync error: {ex.Message}");
                }
                return results;
            });
        }

        /// <summary>
        /// Counts slides in a PPTX file by counting slide XML files in the zip.
        /// </summary>
        public int GetSlideCount(string filePath)
        {
            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var archive = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Read))
                {
                    return archive.Entries
                            .Count(e => e.FullName.StartsWith("ppt/slides/slide") &&
                                    e.FullName.EndsWith(".xml") &&
                                    !e.FullName.Contains("_rels"));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SlideLibraryManager.GetSlideCount error: {ex.Message}");
                return 0;
            }
        }

        // ─── Thumbnail Cache ────────────────────────────────────────────────────

        /// <summary>Returns the cache folder for a given PPTX file.</summary>
        public string GetCacheFolder(string filePath)
        {
            return _cacheService.GetCacheFolder(filePath);
        }

        /// <summary>Returns true when cached thumbnails exist and the PPTX hasn't changed since.</summary>
        public bool IsCacheValid(string filePath)
        {
            return _cacheService.IsCacheValid(filePath);
        }

        /// <summary>
        /// Returns the path to the cached JPG for slide <paramref name="slideIndex"/> (1-based),
        /// or null if it doesn't exist.
        /// </summary>
        public string GetCachedThumbnailPath(string filePath, int slideIndex)
        {
            string path = Path.Combine(GetCacheFolder(filePath), $"s{slideIndex}.jpg");
            return File.Exists(path) ? path : null;
        }

        /// <summary>
        /// Loads a cached slide thumbnail into a frozen BitmapImage, or returns null.
        /// Safe to call from any thread.
        /// </summary>
        public BitmapImage LoadCachedThumbnail(string filePath, int slideIndex)
        {
            string thumbPath = GetCachedThumbnailPath(filePath, slideIndex);
            if (thumbPath == null) return null;

            try
            {
                var img = new BitmapImage();
                img.BeginInit();
                img.CacheOption = BitmapCacheOption.OnLoad;
                img.UriSource   = new Uri(thumbPath);
                img.EndInit();
                img.Freeze();
                return img;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SlideLibraryManager.LoadCachedThumbnail error: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Opens the PPTX presentation via PowerPoint COM (must be called on the STA/UI thread),
        /// exports every slide to the cache folder as a JPG, and writes the timestamp marker.
        /// Calls <paramref name="onSlideReady"/> after each slide so the UI can update progressively.
        /// </summary>
        public async Task GenerateThumbnailCache(string filePath,
                                           Action<int, BitmapImage> onSlideReady = null)
        {
            // Ensure we are on the UI thread for COM operations
            if (!System.Windows.Application.Current.Dispatcher.CheckAccess())
            {
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(async () => await GenerateThumbnailCache(filePath, onSlideReady));
                return;
            }

            Microsoft.Office.Interop.PowerPoint.Presentation pres = null;
            try
            {
                string cacheDir = GetCacheFolder(filePath);
                if (!Directory.Exists(cacheDir)) Directory.CreateDirectory(cacheDir);

                var app  = _manager.GetApplication();
                pres = app.Presentations.Open(
                    filePath,
                    Microsoft.Office.Core.MsoTriState.msoTrue,   // ReadOnly
                    Microsoft.Office.Core.MsoTriState.msoFalse,  // Untitled
                    Microsoft.Office.Core.MsoTriState.msoFalse); // WithWindow = hidden

                int count = pres.Slides.Count;
                int successCount = 0;

                for (int i = 1; i <= count; i++)
                {
                    try 
                    {
                        string thumbPath = Path.Combine(cacheDir, $"s{i}.jpg");
                        pres.Slides[i].Export(thumbPath, "JPG", 640, 360);

                        if (onSlideReady != null && File.Exists(thumbPath))
                        {
                            var img = new BitmapImage();
                            img.BeginInit();
                            img.CacheOption = BitmapCacheOption.OnLoad;
                            img.UriSource   = new Uri(thumbPath);
                            img.EndInit();
                            img.Freeze();
                            onSlideReady(i, img);
                        }
                        successCount++;
                        
                        // Yield slightly to keep the UI thread responsive
                        if (i % 3 == 0) await Task.Delay(1);
                    }
                    catch (Exception slideEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to export slide {i} from {Path.GetFileName(filePath)}: {slideEx.Message}");
                    }
                }

                // Register source path and write timestamp via cache service
                _cacheService.RegisterSourcePath(filePath, cacheDir);
                
                if (successCount > 0 || count == 0)
                {
                    _cacheService.WriteTimestamp(filePath);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GenerateThumbnailCache failed for {Path.GetFileName(filePath)}: {ex.Message}");
            }
            finally
            {
                try { pres?.Close(); } catch { }
                if (pres != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(pres);
            }
        }

        // ─── Insertion ──────────────────────────────────────────────────────────

        public void InsertSlide(string filePath, int slideIndex, bool matchTheme)
        {
            InsertLibrarySlideAtIndex(filePath, matchTheme, slideIndex);
        }

        public (bool Success, string ErrorMessage) InsertLibrarySlide(string sourcePath, bool matchTheme)
            => InsertLibrarySlideAtIndex(sourcePath, matchTheme, -1);

        public (bool Success, string ErrorMessage) InsertLibrarySlideAtIndex(string sourcePath, bool matchTheme, int slideIndex)
        {
            try
            {
                var app = _manager.GetApplication();
                
                // Check if PowerPoint application is accessible
                if (app == null)
                {
                    string error = "PowerPoint application is not accessible";
                    System.Diagnostics.Debug.WriteLine("Insert Error: " + error);
                    return (false, error);
                }

                // Check if there's an active presentation
                if (app.ActivePresentation == null)
                {
                    string error = "No active presentation found";
                    System.Diagnostics.Debug.WriteLine("Insert Error: " + error);
                    return (false, error);
                }

                // Check if the presentation has slides
                if (app.ActivePresentation.Slides == null || app.ActivePresentation.Slides.Count == 0)
                {
                    string error = "Presentation has no slides or slides are not accessible";
                    System.Diagnostics.Debug.WriteLine("Insert Error: " + error);
                    return (false, error);
                }

                // Check if source file exists and is accessible
                if (!System.IO.File.Exists(sourcePath))
                {
                    string error = $"Source file not found: {sourcePath}";
                    System.Diagnostics.Debug.WriteLine($"Insert Error: {error}");
                    return (false, error);
                }

                // Verify actual slide count in source file using PowerPoint
                int actualSlideCount = 0;
                Presentation sourcePres = null;
                try
                {
                    sourcePres = app.Presentations.Open(
                        sourcePath,
                        Microsoft.Office.Core.MsoTriState.msoFalse,  // ReadOnly = false (allow editing)
                        Microsoft.Office.Core.MsoTriState.msoFalse,  // Untitled
                        Microsoft.Office.Core.MsoTriState.msoFalse); // WithWindow = hidden
                    actualSlideCount = sourcePres.Slides.Count;
                    System.Diagnostics.Debug.WriteLine($"Source file verification: {sourcePath} has {actualSlideCount} slides");
                    
                    // Unhide all hidden slides to workaround InsertFromFile limitations
                    int hiddenCount = 0;
                    foreach (Slide slide in sourcePres.Slides)
                    {
                        if (slide.SlideShowTransition.Hidden == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            slide.SlideShowTransition.Hidden = Microsoft.Office.Core.MsoTriState.msoFalse;
                            hiddenCount++;
                        }
                    }
                    if (hiddenCount > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"Unhid {hiddenCount} slides in source file");
                        // Save the unhidden state back to the file
                        sourcePres.Save();
                    }
                    
                    sourcePres.Close();
                    
                    // Validate requested slide index against actual count
                    if (slideIndex > 0 && slideIndex > actualSlideCount)
                    {
                        string error = $"Requested slide index {slideIndex} exceeds the actual slide count of {actualSlideCount} in the source file. " +
                                     "The library may have stale data. Please refresh the library or check the source file.";
                        System.Diagnostics.Debug.WriteLine($"Insert Error: {error}");
                        return (false, error);
                    }
                    
                    // If verification succeeded but actualSlideCount is 0, something went wrong
                    if (actualSlideCount == 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"Warning: Could not determine actual slide count, proceeding with requested index {slideIndex}");
                    }
                    else if (slideIndex > 0 && slideIndex <= actualSlideCount)
                    {
                        System.Diagnostics.Debug.WriteLine($"Slide index {slideIndex} is within valid range (1-{actualSlideCount})");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Warning: Could not verify source slide count: {ex.Message}");
                }

                // Start Undo Entry
                try { app.StartNewUndoEntry(); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"StartNewUndoEntry error: {ex.Message}"); }

                // Calculate insert index - PowerPoint slides are 1-based, but InsertFromFile range parameters might be 0-based
                int insertIndex = 1;
                try
                {
                    if (app.ActiveWindow != null && app.ActiveWindow.ViewType == PpViewType.ppViewNormal)
                    {
                        // Check if we can access the selection
                        try
                        {
                            insertIndex = app.ActiveWindow.Selection.SlideRange.SlideIndex + 1;
                        }
                        catch
                        {
                            // If we can't access the selection, use the end
                            insertIndex = app.ActivePresentation.Slides.Count + 1;
                        }
                    }
                    else
                    {
                        insertIndex = app.ActivePresentation.Slides.Count + 1;
                    }
                }
                catch 
                { 
                    insertIndex = app.ActivePresentation.Slides.Count + 1; 
                }

                // For InsertFromFile, the slideIndex parameter is 1-based for the source presentation
                // but we need to ensure we don't try to insert beyond the current presentation
                if (insertIndex < 1) insertIndex = 1;
                // Don't limit the insertIndex - PowerPoint should handle inserting at the end properly

                if (!matchTheme)
                {
                    // Validate slideIndex before attempting insertion
                    if (slideIndex > 0 && slideIndex > 50) // Sanity check - most template files won't have more than 50 slides
                    {
                        string error = $"Requested slide index {slideIndex} seems unusually high. Please verify the slide number in the template file.";
                        System.Diagnostics.Debug.WriteLine($"Insert Error: {error}");
                        return (false, error);
                    }
                    
                    // PowerPoint InsertFromFile uses 0-based indexing for slide ranges (0 to slideCount-1)
                    // But our library uses 1-based indexing (1 to slideCount)
                    // So we need to convert: 1-based index -> 0-based index
                    int ppSlideIndex = slideIndex > 0 ? slideIndex - 1 : slideIndex;
                    
                    // If we verified the actual slide count and it's different from what InsertFromFile might expect,
                    // we should still try with the original index first, as InsertFromFile may see a different structure
                    System.Diagnostics.Debug.WriteLine($"InsertFromFile Debug: sourcePath={sourcePath}, insertIndex={insertIndex}, slideIndex={slideIndex}, ppSlideIndex={ppSlideIndex}, verifiedCount={actualSlideCount}");
                    
                    // Insert only the target slide (or all if slideIndex == -1)
                    try
                    {
                        if (slideIndex > 0)
                        {
                            app.ActivePresentation.Slides.InsertFromFile(sourcePath, insertIndex, ppSlideIndex, ppSlideIndex);
                        }
                        else
                        {
                            app.ActivePresentation.Slides.InsertFromFile(sourcePath, insertIndex);
                        }
                    }
                    catch (Exception ex)
                    {
                        string error = $"Failed to insert slide from file: {ex.Message}";
                        System.Diagnostics.Debug.WriteLine($"InsertFromFile Error: {ex.Message}");
                        
                        // Try fallback using copy-paste method if InsertFromFile fails
                        if (ex.Message.Contains("Integer out of range") || ex.Message.Contains("unknown member"))
                        {
                            System.Diagnostics.Debug.WriteLine($"Attempting fallback copy-paste method...");
                            bool fallbackSuccess = TryInsertSlideUsingCopyPaste(app, sourcePath, insertIndex, slideIndex);
                            if (fallbackSuccess)
                            {
                                return (true, null);
                            }
                        }
                        
                        // Add helpful context for common errors
                        if (ex.Message.Contains("Integer out of range"))
                        {
                            error += $"\n\nThe source file shows {actualSlideCount} slides when verified, " +
                                     $"but InsertFromFile reports a different count. " +
                                     $"Requested index: {slideIndex} (0-based: {ppSlideIndex}).\n\n" +
                                     "This may be due to:\n" +
                                     "- Hidden slides in the source file\n" +
                                     "- File format compatibility issues\n" +
                                     "- Corrupted slide structure\n\n" +
                                     "Try:\n" +
                                     "1. Re-saving the source file in PowerPoint\n" +
                                     "2. Using a different template file\n" +
                                     "3. Inserting all slides instead of a specific one";
                        }
                        
                        return (false, error);
                    }
                }
                else
                {
                    // Checked: Match Destination Theme
                    // 1. Insert slide
                    // InsertFromFile always keeps source formatting initially in recent PPT versions unless pasted.
                    // But we can clean it up.
                    
                    int countBefore = app.ActivePresentation.Slides.Count;
                    try
                    {
                        if (slideIndex > 0)
                        {
                            // Convert 1-based slide index to 0-based for PowerPoint API
                            int ppSlideIndex = slideIndex - 1;
                            app.ActivePresentation.Slides.InsertFromFile(sourcePath, insertIndex, ppSlideIndex, ppSlideIndex);
                        }
                        else
                        {
                            app.ActivePresentation.Slides.InsertFromFile(sourcePath, insertIndex);
                        }
                    }
                    catch (Exception ex)
                    {
                        string error = $"Failed to insert slide from file: {ex.Message}";
                        System.Diagnostics.Debug.WriteLine($"InsertFromFile Error: {ex.Message}");
                        
                        // Try fallback using copy-paste method if InsertFromFile fails
                        if (ex.Message.Contains("Integer out of range") || ex.Message.Contains("unknown member"))
                        {
                            System.Diagnostics.Debug.WriteLine($"Attempting fallback copy-paste method...");
                            bool fallbackSuccess = TryInsertSlideUsingCopyPaste(app, sourcePath, insertIndex, slideIndex);
                            if (fallbackSuccess)
                            {
                                return (true, null);
                            }
                        }
                        
                        // Add helpful context for common errors
                        if (ex.Message.Contains("Integer out of range"))
                        {
                            int ppSlideIndex = slideIndex > 0 ? slideIndex - 1 : slideIndex;
                            error += $"\n\nThis usually means the slide index {slideIndex} (0-based: {ppSlideIndex}) doesn't exist in the source file. " +
                                     "The source file might have fewer slides than expected, or the slide numbering might be different.";
                        }
                        
                        return (false, error);
                    }
                    
                    int countAfter = app.ActivePresentation.Slides.Count;
                    int insertedCount = countAfter - countBefore;

                    if (insertedCount > 0)
                    {
                        try
                        {
                            var insertedSlide = app.ActivePresentation.Slides[insertIndex];
                            
                            // 2. Layout Cleaning & Theme Matching
                            // Find a layout in the current presentation that matches the inserted slide's layout name
                            CustomLayout targetLayout = null;
                            string sourceLayoutName = insertedSlide.CustomLayout.Name;

                            foreach (CustomLayout layout in app.ActivePresentation.SlideMaster.CustomLayouts)
                            {
                                if (layout.Name.Equals(sourceLayoutName, StringComparison.OrdinalIgnoreCase))
                                {
                                    targetLayout = layout;
                                    break;
                                }
                            }

                            if (targetLayout != null)
                            {
                                insertedSlide.CustomLayout = targetLayout;
                            }
                            
                            // 3. Color/Font Sync
                            // Resetting the theme to the destination master
                            insertedSlide.Design = app.ActivePresentation.SlideMaster.Design;
                            insertedSlide.ColorScheme = app.ActivePresentation.SlideMaster.ColorScheme;
                            
                            // Clean up: If the insertion created a new Master (common with InsertFromFile), delete it if unused
                            // This is complex because we just reassigned the layout. 
                            // The old master might now be unused.
                            CleanupUnusedMasters(app.ActivePresentation);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Theme matching Error: {ex.Message}");
                            // Don't fail the whole operation if theme matching fails
                        }
                    }
                }

                return (true, null);
            }
            catch (Exception ex)
            {
                string error = $"An error occurred while inserting the slide: {ex.Message}";
                System.Diagnostics.Debug.WriteLine("Insert Error: " + ex.Message);
                return (false, error);
            }
        }

        private void CleanupUnusedMasters(Presentation pres)
        {
            try
            {
                // Iterate backwards to safely remove unused designs
                for (int i = pres.Designs.Count; i >= 2; i--) // Keep at least 1 design
                {
                    var design = pres.Designs[i];
                    bool used = false;
                    foreach (Slide s in pres.Slides)
                    {
                        if (s.Design.Name == design.Name)
                        {
                            used = true;
                            break;
                        }
                    }
                    if (!used)
                    {
                        try { design.Delete(); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"CleanupUnusedMasters.Delete error: {ex.Message}"); }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SlideLibraryManager.CleanupUnusedMasters error: {ex.Message}");
            }
        }

        /// <summary>
        /// Fallback method using copy-paste when InsertFromFile fails.
        /// Opens source presentation, copies the specific slide, pastes into target presentation.
        /// </summary>
        private bool TryInsertSlideUsingCopyPaste(Application app, string sourcePath, int insertIndex, int slideIndex)
        {
            Presentation sourcePres = null;
            try
            {
                System.Diagnostics.Debug.WriteLine($"Fallback: Opening source presentation: {sourcePath}");
                
                // Open source presentation
                sourcePres = app.Presentations.Open(
                    sourcePath,
                    Microsoft.Office.Core.MsoTriState.msoTrue,   // ReadOnly = true (safer)
                    Microsoft.Office.Core.MsoTriState.msoFalse,  // Untitled
                    Microsoft.Office.Core.MsoTriState.msoFalse); // WithWindow = hidden

                // Verify slide exists in source
                if (slideIndex < 1 || slideIndex > sourcePres.Slides.Count)
                {
                    System.Diagnostics.Debug.WriteLine($"Fallback: Slide index {slideIndex} out of range (1-{sourcePres.Slides.Count})");
                    return false;
                }

                System.Diagnostics.Debug.WriteLine($"Fallback: Copying slide {slideIndex} of {sourcePres.Slides.Count}");
                
                // Select and copy the slide from source
                var sourceSlide = sourcePres.Slides[slideIndex];
                
                // Copy the slide (no need to select first for Copy)
                sourceSlide.Copy();
                System.Diagnostics.Debug.WriteLine($"Fallback: Slide copied to clipboard");

                // Paste into target presentation
                var targetPres = app.ActivePresentation;
                
                // Validate insertIndex - ensure it's within valid range
                if (insertIndex > targetPres.Slides.Count)
                    insertIndex = targetPres.Slides.Count + 1;
                if (insertIndex < 1)
                    insertIndex = 1;

                System.Diagnostics.Debug.WriteLine($"Fallback: Pasting at position {insertIndex} in target presentation with {targetPres.Slides.Count} slides");
                
                // Paste the slide
                SlideRange pastedSlides = targetPres.Slides.Paste(insertIndex);
                
                if (pastedSlides != null && pastedSlides.Count > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"Fallback: Successfully pasted {pastedSlides.Count} slide(s) using copy-paste method");
                    return true;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"Fallback: Paste returned null or empty");
                    return false;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Fallback copy-paste failed: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Fallback exception type: {ex.GetType().Name}");
                return false;
            }
            finally
            {
                // Always close source presentation
                if (sourcePres != null)
                {
                    try 
                    { 
                        sourcePres.Close();
                        System.Diagnostics.Debug.WriteLine($"Fallback: Source presentation closed");
                    } 
                    catch (Exception closeEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"Fallback: Error closing source presentation: {closeEx.Message}");
                    }
                }
            }
        }
    }
}