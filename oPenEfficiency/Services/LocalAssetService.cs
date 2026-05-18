using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Slide = Microsoft.Office.Interop.PowerPoint.Slide;
using oPenEfficiency.Services;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Service for managing user's personal asset folders.
    /// Assets are stored in %APPDATA%\oPenEfficiency\favorites\[type]\
    /// </summary>
    public class LocalAssetService
    {
        private readonly string _baseAssetsPath;
        private readonly PowerPointManager _ppManager;

        private const string SlidesFolder = "slides";
        private const string ImagesFolder = "images";
        private const string ShapesFolder = "shapes";

        public LocalAssetService(PowerPointManager ppManager)
        {
            _ppManager = ppManager;
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _baseAssetsPath = Path.Combine(appDataPath, "oPenEfficiency", "favorites");

            // Ensure base directories exist
            EnsureDirectoryExists(Path.Combine(_baseAssetsPath, SlidesFolder));
            EnsureDirectoryExists(Path.Combine(_baseAssetsPath, ImagesFolder));
            EnsureDirectoryExists(Path.Combine(_baseAssetsPath, ShapesFolder));
        }

        /// <summary>
        /// Gets the base path for a specific asset type.
        /// </summary>
        public string GetAssetFolderPath(AssetType assetType)
        {
            string typeFolder;
            if (assetType == AssetType.Slides)
                typeFolder = SlidesFolder;
            else if (assetType == AssetType.Images)
                typeFolder = ImagesFolder;
            else if (assetType == AssetType.Shapes)
                typeFolder = ShapesFolder;
            else
                throw new ArgumentException("Unknown asset type: " + assetType);

            return Path.Combine(_baseAssetsPath, typeFolder);
        }

        /// <summary>
        /// Gets all folders (including subfolders) for a specific asset type.
        /// </summary>
        public List<AssetFolder> GetAssetFolders(AssetType assetType)
        {
            string basePath = GetAssetFolderPath(assetType);
            var folders = new List<AssetFolder>();

            if (!Directory.Exists(basePath))
                return folders;

            // Add root folder
            folders.Add(new AssetFolder
            {
                Name = "Favorites",
                FullPath = basePath,
                ParentPath = null,
                IsRoot = true
            });

            // Add subfolders
            try
            {
                var subDirs = Directory.GetDirectories(basePath, "*", SearchOption.AllDirectories);
                foreach (var subDir in subDirs.OrderBy(d => d))
                {
                    var dirInfo = new DirectoryInfo(subDir);
                    string parentPath = dirInfo.Parent?.FullName;

                    folders.Add(new AssetFolder
                    {
                        Name = dirInfo.Name,
                        FullPath = subDir,
                        ParentPath = parentPath,
                        IsRoot = false
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LocalAssetService.GetAssetFolders error: {ex.Message}");
            }

            return folders;
        }

        /// <summary>
        /// Creates a new folder in the specified asset type's root directory.
        /// </summary>
        public string CreateFolder(AssetType assetType, string folderName)
        {
            string basePath = GetAssetFolderPath(assetType);
            string safeName = MakeValidFileName(folderName);
            string newPath = Path.Combine(basePath, safeName);

            if (Directory.Exists(newPath))
            {
                throw new IOException($"Folder '{folderName}' already exists.");
            }

            Directory.CreateDirectory(newPath);
            return newPath;
        }

        /// <summary>
        /// Creates a subfolder within an existing folder.
        /// </summary>
        public string CreateSubFolder(string parentPath, string folderName)
        {
            string safeName = MakeValidFileName(folderName);
            string newPath = Path.Combine(parentPath, safeName);

            if (Directory.Exists(newPath))
            {
                throw new IOException($"Folder '{folderName}' already exists.");
            }

            Directory.CreateDirectory(newPath);
            return newPath;
        }

        /// <summary>
        /// Saves the currently selected slide(s) to the user's assets folder.
        /// </summary>
        public async Task<SaveResult> SaveSelectedSlides(string targetFolder, string fileName, string keywords = null, bool autoTag = false)
        {
            return await Task.Run(() =>
            {
                try
                {
                    var app = _ppManager.GetApplication();
                    if (app == null || app.ActivePresentation == null)
                        return new SaveResult { Success = false, ErrorMessage = "No active presentation found." };

                    var selection = app.ActiveWindow.Selection;
                    if (selection == null || selection.Type == PpSelectionType.ppSelectionNone)
                        return new SaveResult { Success = false, ErrorMessage = "No slide(s) selected." };

                    // Get selected slides
                    var selectedSlides = new List<Slide>();
                    try
                    {
                        SlideRange slideRange = selection.SlideRange;
                        if (slideRange != null)
                        {
                            for (int i = 1; i <= slideRange.Count; i++)
                            {
                                selectedSlides.Add(slideRange[i]);
                            }
                        }
                    }
                    catch
                    {
                        // If SlideRange access fails, try getting active slide
                        try
                        {
                            Slide activeSlide = (Slide)app.ActiveWindow.View.Slide;
                            if (activeSlide != null)
                                selectedSlides.Add(activeSlide);
                        }
                        catch (Exception ex)
                        {
                            return new SaveResult { Success = false, ErrorMessage = "Could not get selected slides: " + ex.Message };
                        }
                    }

                    if (selectedSlides.Count == 0)
                        return new SaveResult { Success = false, ErrorMessage = "No slides selected." };

                    // Auto-tagging
                    string finalKeywords = keywords ?? "";
                    if (autoTag)
                    {
                        var autoTags = new List<string>();
                        foreach (var slide in selectedSlides)
                        {
                            autoTags.AddRange(AutoTaggingService.IdentifyTags(slide));
                        }
                        
                        var existing = finalKeywords.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim());
                        var combined = existing.Concat(autoTags).Distinct();
                        finalKeywords = string.Join(", ", combined);
                    }

                    // Ensure target folder exists
                    EnsureDirectoryExists(targetFolder);

                    // Generate unique filename
                    string safeFileName = MakeValidFileName(fileName);
                    string targetPath = GetUniqueFilePath(targetFolder, safeFileName, ".pptx");

                    // Create new presentation and copy selected slides
                    Presentation newPres = app.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoFalse);
                    try
                    {
                        foreach (var slide in selectedSlides)
                        {
                            slide.Copy();
                            newPres.Slides.Paste(newPres.Slides.Count + 1);
                        }

                        if (!string.IsNullOrEmpty(finalKeywords))
                        {
                            try
                            {
                                Microsoft.Office.Core.DocumentProperties props = (Microsoft.Office.Core.DocumentProperties)newPres.BuiltInDocumentProperties;
                                props["Keywords"].Value = finalKeywords;
                            }
                            catch { }
                        }

                        newPres.SaveAs(targetPath);
                        return new SaveResult
                        {
                            Success = true,
                            SavedPath = targetPath,
                            ItemCount = selectedSlides.Count
                        };
                    }
                    finally
                    {
                        newPres.Close();
                    }
                }
                catch (Exception ex)
                {
                    return new SaveResult { Success = false, ErrorMessage = ex.Message };
                }
            });
        }

        /// <summary>
        /// Saves selected images/shapes from the active slide to the user's assets folder.
        /// </summary>
        public async Task<SaveResult> SaveSelectedShapes(string targetFolder, string fileName, string keywords = null, bool autoTag = false)
        {
            return await Task.Run(() =>
            {
                try
                {
                    var app = _ppManager.GetApplication();
                    if (app == null || app.ActivePresentation == null)
                        return new SaveResult { Success = false, ErrorMessage = "No active presentation found." };

                    var selection = app.ActiveWindow.Selection;
                    if (selection == null || (selection.Type != PpSelectionType.ppSelectionShapes && selection.Type != PpSelectionType.ppSelectionText))
                        return new SaveResult { Success = false, ErrorMessage = "No shape(s) selected." };

                    // Auto-tagging (shapes can also be tagged by slide context)
                    string finalKeywords = keywords ?? "";
                    if (autoTag)
                    {
                        try
                        {
                            Slide activeSlide = (Slide)app.ActiveWindow.View.Slide;
                            var autoTags = AutoTaggingService.IdentifyTags(activeSlide);
                            
                            var existing = finalKeywords.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim());
                            var combined = existing.Concat(autoTags).Distinct();
                            finalKeywords = string.Join(", ", combined);
                        }
                        catch { }
                    }

                    EnsureDirectoryExists(targetFolder);
                    string safeFileName = MakeValidFileName(fileName);

                    string targetPath = GetUniqueFilePath(targetFolder, safeFileName, ".pptx");

                    try
                    {
                        ShapeRange shapeRange = null;
                        if (selection.Type == PpSelectionType.ppSelectionShapes)
                            shapeRange = selection.ShapeRange;
                        else if (selection.Type == PpSelectionType.ppSelectionText)
                            shapeRange = selection.ShapeRange;

                        if (shapeRange != null && shapeRange.Count > 0)
                        {
                            shapeRange.Copy();

                            Presentation newPres = app.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoFalse);
                            try
                            {
                                newPres.PageSetup.SlideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                                newPres.PageSetup.SlideHeight = app.ActivePresentation.PageSetup.SlideHeight;

                                // Layout 7 is typically Blank
                                Slide slide = newPres.Slides.AddSlide(1, newPres.SlideMaster.CustomLayouts[7]);
                                
                                var pastedRange = slide.Shapes.Paste();
                                
                                var metadataList = new List<ShapeMetadataItem>();
                                for (int i = 1; i <= pastedRange.Count; i++)
                                {
                                    Shape s = pastedRange[i];
                                    metadataList.Add(new ShapeMetadataItem
                                    {
                                        OriginalShapeId = s.Id,
                                        Name = s.Name,
                                        Left = s.Left,
                                        Top = s.Top,
                                        Width = s.Width,
                                        Height = s.Height,
                                        SlideWidth = app.ActivePresentation.PageSetup.SlideWidth,
                                        SlideHeight = app.ActivePresentation.PageSetup.SlideHeight,
                                        Keywords = finalKeywords
                                    });
                                }

                                newPres.SaveAs(targetPath);

                                string metaPath = targetPath + ".meta.json";
                                JsonService.Save(metadataList, metaPath);

                                return new SaveResult
                                {
                                    Success = true,
                                    SavedPath = targetPath,
                                    ItemCount = pastedRange.Count
                                };
                            }
                            finally
                            {
                                newPres.Close();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(newPres);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        return new SaveResult { Success = false, ErrorMessage = "Could not save shapes: " + ex.Message };
                    }

                    return new SaveResult { Success = false, ErrorMessage = "No shapes could be exported." };
                }
                catch (Exception ex)
                {
                    return new SaveResult { Success = false, ErrorMessage = ex.Message };
                }
            });
        }

        /// <summary>
        /// Saves selected image files to the user's assets folder.
        /// </summary>
        public async Task<SaveResult> SaveSelectedImages(List<string> sourceImagePaths, string targetFolder, string fileName)
        {
            return await Task.Run(() =>
            {
                try
                {
                    if (sourceImagePaths == null || sourceImagePaths.Count == 0)
                        return new SaveResult { Success = false, ErrorMessage = "No images selected." };

                    EnsureDirectoryExists(targetFolder);
                    string safeFileName = MakeValidFileName(fileName);

                    int savedCount = 0;
                    string firstSavedPath = null;

                    foreach (var sourcePath in sourceImagePaths)
                    {
                        if (!File.Exists(sourcePath))
                            continue;

                        string extension = Path.GetExtension(sourcePath);
                        string uniqueName = GetUniqueFileName(targetFolder, safeFileName, extension);
                        string targetPath = Path.Combine(targetFolder, uniqueName);

                        try
                        {
                            File.Copy(sourcePath, targetPath, true);
                            if (firstSavedPath == null)
                                firstSavedPath = targetPath;
                            savedCount++;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Image copy failed: {ex.Message}");
                        }
                    }

                    if (savedCount == 0)
                        return new SaveResult { Success = false, ErrorMessage = "No images could be copied." };

                    return new SaveResult
                    {
                        Success = true,
                        SavedPath = firstSavedPath,
                        ItemCount = savedCount
                    };
                }
                catch (Exception ex)
                {
                    return new SaveResult { Success = false, ErrorMessage = ex.Message };
                }
            });
        }

        public string GetAssetKeywords(string filePath)
        {
            try
            {
                if (!File.Exists(filePath)) return "";

                string extension = Path.GetExtension(filePath).ToLower();
                if (extension == ".pptx")
                {
                    var ppManager = new PowerPointManager(Globals.ThisAddIn.Application);
                    var app = ppManager.GetApplication();
                    
                    Presentation pres = null;
                    try
                    {
                        pres = app.Presentations.Open(filePath, 
                            Microsoft.Office.Core.MsoTriState.msoTrue, 
                            Microsoft.Office.Core.MsoTriState.msoFalse, 
                            Microsoft.Office.Core.MsoTriState.msoFalse);
                        
                        var props = (Microsoft.Office.Core.DocumentProperties)pres.BuiltInDocumentProperties;
                        return props["Keywords"].Value?.ToString() ?? "";
                    }
                    catch
                    {
                        return "";
                    }
                    finally
                    {
                        pres?.Close();
                        if (pres != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(pres);
                    }
                }
                
                return "";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LocalAssetService.GetAssetKeywords error: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// Extracts a specific slide from a source file and saves it as a new asset.
        /// </summary>
        public async Task<SaveResult> ExtractAndSaveSlide(string sourcePath, int slideIndex, string targetFolder, string fileName, string keywords = null)
        {
            return await Task.Run(() =>
            {
                try
                {
                    var app = _ppManager.GetApplication();
                    if (app == null) return new SaveResult { Success = false, ErrorMessage = "PowerPoint not available." };

                    // Open source presentation invisibly
                    Presentation sourcePres = app.Presentations.Open(sourcePath, 
                        Microsoft.Office.Core.MsoTriState.msoTrue, // ReadOnly
                        Microsoft.Office.Core.MsoTriState.msoFalse, // Untitled
                        Microsoft.Office.Core.MsoTriState.msoFalse); // WithWindow

                    try
                    {
                        if (slideIndex < 1 || slideIndex > sourcePres.Slides.Count)
                            return new SaveResult { Success = false, ErrorMessage = "Invalid slide index." };

                        Slide sourceSlide = sourcePres.Slides[slideIndex];
                        
                        // Create new presentation
                        Presentation newPres = app.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoFalse);
                        try
                        {
                            sourceSlide.Copy();
                            newPres.Slides.Paste(1);

                            // Apply keywords
                            if (!string.IsNullOrEmpty(keywords))
                            {
                                try
                                {
                                    Microsoft.Office.Core.DocumentProperties props = (Microsoft.Office.Core.DocumentProperties)newPres.BuiltInDocumentProperties;
                                    props["Keywords"].Value = keywords;
                                }
                                catch { }
                            }

                            // Ensure target folder exists
                            EnsureDirectoryExists(targetFolder);
                            string safeFileName = MakeValidFileName(fileName);
                            string targetPath = GetUniqueFilePath(targetFolder, safeFileName, ".pptx");

                            newPres.SaveAs(targetPath);
                            return new SaveResult
                            {
                                Success = true,
                                SavedPath = targetPath,
                                ItemCount = 1
                            };
                        }
                        finally
                        {
                            newPres.Close();
                        }
                    }
                    finally
                    {
                        sourcePres.Close();
                    }
                }
                catch (Exception ex)
                {
                    return new SaveResult { Success = false, ErrorMessage = ex.Message };
                }
            });
        }

        /// <summary>
        /// Updates the name and keywords of an existing asset.
        /// </summary>
        public bool UpdateAssetDetails(string filePath, string newName, string newKeywords)
        {
            try
            {
                string directory = Path.GetDirectoryName(filePath);
                string extension = Path.GetExtension(filePath);
                string newPath = Path.Combine(directory, MakeValidFileName(newName) + extension);

                // 1. Handle Keywords based on type
                if (extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase))
                {
                    // Check if it's a shape (has .meta.json) or a full slide
                    string metaPath = filePath + ".meta.json";
                    if (File.Exists(metaPath))
                    {
                        // Shape asset - update meta.json
                        var metadata = JsonService.Load<List<ShapeMetadataItem>>(metaPath);
                        if (metadata != null)
                        {
                            foreach (var item in metadata)
                            {
                                item.Keywords = newKeywords;
                                item.Name = newName; // Also update name in metadata
                            }
                            JsonService.Save(metadata, metaPath);
                        }
                    }
                    else
                    {
                        // Full slide asset - update PPTX built-in properties
                        Presentation pres = null;
                        try
                        {
                            var app = _ppManager.GetApplication();
                            pres = app.Presentations.Open(filePath, 
                                Microsoft.Office.Core.MsoTriState.msoTrue, 
                                Microsoft.Office.Core.MsoTriState.msoFalse, 
                                Microsoft.Office.Core.MsoTriState.msoFalse);
                            
                            Microsoft.Office.Core.DocumentProperties props = (Microsoft.Office.Core.DocumentProperties)pres.BuiltInDocumentProperties;
                            props["Keywords"].Value = newKeywords;
                            pres.Save();
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"UpdateAssetDetails PPTX error: {ex.Message}");
                        }
                        finally
                        {
                            pres?.Close();
                            if (pres != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(pres);
                        }
                    }
                }

                // 2. Rename file if name changed
                if (!filePath.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                {
                    File.Move(filePath, newPath);
                    
                    // Also move meta.json if it exists
                    string oldMeta = filePath + ".meta.json";
                    string newMeta = newPath + ".meta.json";
                    if (File.Exists(oldMeta))
                    {
                        File.Move(oldMeta, newMeta);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LocalAssetService.UpdateAssetDetails error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Deletes an asset file and its metadata.
        /// </summary>
        public bool DeleteAsset(string path)
        {
            try
            {
                if (!path.StartsWith(_baseAssetsPath, StringComparison.OrdinalIgnoreCase))
                    return false; 

                if (File.Exists(path))
                {
                    File.Delete(path);
                    
                    // Also delete meta.json
                    string metaPath = path + ".meta.json";
                    if (File.Exists(metaPath)) File.Delete(metaPath);
                    
                    return true;
                }
                else if (Directory.Exists(path))
                {
                    if (!Directory.EnumerateFileSystemEntries(path).Any())
                    {
                        Directory.Delete(path);
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LocalAssetService.DeleteAsset error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Renames an asset file or folder.
        /// </summary>
        public bool RenameAsset(string oldPath, string newName)
        {
            try
            {
                if (!oldPath.StartsWith(_baseAssetsPath, StringComparison.OrdinalIgnoreCase))
                    return false; // Security check

                string safeName = MakeValidFileName(newName);
                string newPath = Path.Combine(Path.GetDirectoryName(oldPath), safeName);

                if (File.Exists(oldPath))
                {
                    File.Move(oldPath, newPath);
                    return true;
                }
                else if (Directory.Exists(oldPath))
                {
                    Directory.Move(oldPath, newPath);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LocalAssetService.RenameAsset error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Gets all image files in the images folder (recursively).
        /// </summary>
        public List<AssetFile> GetImageFiles(string folder = null)
        {
            string searchFolder = folder ?? GetAssetFolderPath(AssetType.Images);
            return GetAssetFilesInFolder(searchFolder, new[] { "*.png", "*.jpg", "*.jpeg", "*.svg", "*.gif", "*.bmp" });
        }

        /// <summary>
        /// Gets all slide files in the slides folder (recursively).
        /// </summary>
        public List<AssetFile> GetSlideFiles(string folder = null)
        {
            string searchFolder = folder ?? GetAssetFolderPath(AssetType.Slides);
            return GetAssetFilesInFolder(searchFolder, new[] { "*.pptx" });
        }

        /// <summary>
        /// Gets all shape files in the shapes folder (recursively).
        /// </summary>
        public List<AssetFile> GetShapeFiles(string folder = null)
        {
            string searchFolder = folder ?? GetAssetFolderPath(AssetType.Shapes);
            return GetAssetFilesInFolder(searchFolder, new[] { "*.png" });
        }

        #region Helpers

        private List<AssetFile> GetAssetFilesInFolder(string folder, string[] extensions)
        {
            var files = new List<AssetFile>();

            if (!Directory.Exists(folder))
                return files;

            try
            {
                foreach (var ext in extensions)
                {
                    var foundFiles = Directory.GetFiles(folder, ext, SearchOption.AllDirectories);
                    foreach (var filePath in foundFiles)
                    {
                        var fileInfo = new FileInfo(filePath);
                        string relativePath = GetRelativePath(_baseAssetsPath, filePath);

                        files.Add(new AssetFile
                        {
                            Name = Path.GetFileNameWithoutExtension(filePath),
                            FilePath = filePath,
                            Extension = fileInfo.Extension.ToLower(),
                            LastModified = fileInfo.LastWriteTimeUtc,
                            RelativePath = relativePath
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LocalAssetService.GetAssetFilesInFolder error: {ex.Message}");
            }

            return files.OrderBy(f => f.Name).ToList();
        }

        private static void EnsureDirectoryExists(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        /// <summary>
        /// Gets the relative path from basePath to targetPath.
        /// Replacement for Path.GetRelativePath (not available in .NET Framework).
        /// </summary>
        private static string GetRelativePath(string basePath, string targetPath)
        {
            // Ensure paths end with backslash for proper comparison
            if (!basePath.EndsWith("\\")) basePath += "\\";

            // Normalize paths
            string baseNormalized = basePath.Replace("\\\\", "\\").ToLower();
            string targetNormalized = targetPath.Replace("\\\\", "\\").ToLower();

            // If target starts with base, return the relative portion
            if (targetNormalized.StartsWith(baseNormalized))
            {
                return targetPath.Substring(basePath.Length);
            }

            // Fallback: return absolute path
            return targetPath;
        }

        private static string MakeValidFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return "unnamed";

            // Remove invalid characters
            char[] invalidChars = Path.GetInvalidFileNameChars();
            string validName = string.Concat(fileName.Split(invalidChars));

            // Trim and replace spaces with underscores for cleaner filenames
            validName = validName.Trim().Replace(" ", "_");

            return validName.Length > 100 ? validName.Substring(0, 100) : validName;
        }

        private string GetUniqueFilePath(string folder, string baseName, string extension)
        {
            string fileName = GetUniqueFileName(folder, baseName, extension);
            return Path.Combine(folder, fileName);
        }

        private string GetUniqueFileName(string folder, string baseName, string extension)
        {
            string fileName = $"{baseName}{extension}";
            string fullPath = Path.Combine(folder, fileName);

            if (!File.Exists(fullPath))
                return fileName;

            // Add number suffix until unique
            int counter = 1;
            while (true)
            {
                fileName = $"{baseName}_{counter}{extension}";
                fullPath = Path.Combine(folder, fileName);
                if (!File.Exists(fullPath))
                    return fileName;
                counter++;
            }
        }

        #endregion
    }

    /// <summary>
    /// Types of assets that can be saved.
    /// </summary>
    public enum AssetType
    {
        Slides,
        Images,
        Shapes
    }

    /// <summary>
    /// Represents a folder in the asset hierarchy.
    /// </summary>
    public class AssetFolder
    {
        public string Name { get; set; }
        public string FullPath { get; set; }
        public string ParentPath { get; set; }
        public bool IsRoot { get; set; }
    }

    /// <summary>
    /// Represents a file in the asset system.
    /// </summary>
    public class AssetFile
    {
        public string Name { get; set; }
        public string FilePath { get; set; }
        public string Extension { get; set; }
        public DateTime LastModified { get; set; }
        public string RelativePath { get; set; }
    }

    /// <summary>
    /// Result of a save operation.
    /// </summary>
    public class SaveResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string SavedPath { get; set; }
        public int ItemCount { get; set; }
    }

    public class ShapeMetadataItem
    {
        public int OriginalShapeId { get; set; }
        public string Name { get; set; }
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
        public float SlideWidth { get; set; }
        public float SlideHeight { get; set; }
        public string Keywords { get; set; }
    }
}
