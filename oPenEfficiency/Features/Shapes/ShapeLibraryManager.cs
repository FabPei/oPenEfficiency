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
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Slide = Microsoft.Office.Interop.PowerPoint.Slide;

namespace oPenEfficiency.Features
{
    public class ShapeLibraryManager
    {
        private readonly PowerPointManager _manager;
        private const string ConfigFile = "shape_library_folders.json";
        private readonly AssetCacheService _cacheService;

        public ShapeLibraryManager(PowerPointManager manager)
        {
            _manager = manager;
            _cacheService = new AssetCacheService();
        }

        // --- Step 1: Library Source Management ---

        public List<string> GetLibraryFolders()
        {
            var list = JsonService.Load<List<string>>(ConfigFile) ?? new List<string>();
            
            // Always add the user's personal "Favorites" folder
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string myAssetsPath = Path.Combine(appDataPath, "oPenEfficiency", "favorites", "shapes");
            if (!Directory.Exists(myAssetsPath)) Directory.CreateDirectory(myAssetsPath);
            
            if (!list.Contains(myAssetsPath, StringComparer.OrdinalIgnoreCase))
                list.Insert(0, myAssetsPath);

            return list;
        }

        public void AddLibraryFolder(string path)
        {
            var list = GetLibraryFolders();
            if (Directory.Exists(path) && !list.Contains(path))
            {
                list.Add(path);
                JsonService.Save(list, ConfigFile);
            }
        }

        public void RemoveLibraryFolder(string path)
        {
            var list = GetLibraryFolders();
            if (list.Contains(path))
            {
                list.Remove(path);
                JsonService.Save(list, ConfigFile);
            }
        }

        public async Task<List<ShapeFile>> ScanLibrariesAsync()
        {
            return await Task.Run(() =>
            {
                var scannedFiles = new List<ShapeFile>();
                var folders = GetLibraryFolders();

                // For shapes, we might support both PPTX (containers of shapes) and standalone EMF/SVG.
                string[] validExtensions = { ".pptx" };

                foreach (var folder in folders)
                {
                    if (Directory.Exists(folder))
                    {
                        try
                        {
                            var files = oPenEfficiency.Utils.CommonUtils.SafeGetFiles(folder, "*.*")
                                                 .Where(f => validExtensions.Contains(Path.GetExtension(f).ToLower()) && !Path.GetFileName(f).StartsWith("~$"))
                                                 .ToList();
                                                 
                            foreach (var file in files)
                            {
                                string fullPath = Path.GetFullPath(file);
                                if (scannedFiles.Any(f => f.FilePath.Equals(fullPath, StringComparison.OrdinalIgnoreCase))) continue;

                                // Compute relative path
                                string fileDir = Path.GetDirectoryName(fullPath);
                                string relativeSub = "";
                                if (fileDir.Length > folder.Length)
                                    relativeSub = fileDir.Substring(folder.Length).TrimStart(Path.DirectorySeparatorChar);

                                scannedFiles.Add(new ShapeFile
                                {
                                    FilePath = fullPath, 
                                    FileName = Path.GetFileNameWithoutExtension(fullPath),
                                    FolderName = new DirectoryInfo(folder).Name,
                                    RootFolder = folder,
                                    SubFolder = relativeSub
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"ShapeLibraryManager.ScanLibrariesAsync error: {ex.Message}");
                        }
                    }
                }
                return scannedFiles;
            });
        }

        public List<FolderTreeNode> BuildFolderTree(List<ShapeFile> files, bool foldersOnly = false)
        {
            var roots = new List<FolderTreeNode>();
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

                var withSub = rootGroup.Where(f => !string.IsNullOrEmpty(f.SubFolder)).ToList();
                BuildChildNodes(rootNode, withSub, rootPath, foldersOnly);

                if (!foldersOnly)
                {
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

                rootNode.TotalFileCount = rootGroup.Count();
                roots.Add(rootNode);
            }

            return roots;
        }

        private void BuildChildNodes(FolderTreeNode parent, List<ShapeFile> files, string rootPath, bool foldersOnly = false)
        {
            var groups = files.GroupBy(f =>
            {
                string rel = f.SubFolder;
                string parentRel = "";
                if (parent.FullPath.Length > rootPath.Length)
                    parentRel = parent.FullPath.Substring(rootPath.Length).TrimStart(Path.DirectorySeparatorChar);

                if (!string.IsNullOrEmpty(parentRel) && rel.StartsWith(parentRel + Path.DirectorySeparatorChar))
                    rel = rel.Substring(parentRel.Length + 1);
                else if (!string.IsNullOrEmpty(parentRel) && rel == parentRel)
                    return ""; 

                int sep = rel.IndexOf(Path.DirectorySeparatorChar);
                return sep >= 0 ? rel.Substring(0, sep) : rel;
            });

            foreach (var g in groups)
            {
                if (string.IsNullOrEmpty(g.Key)) continue;

                string childPath = Path.Combine(parent.FullPath, g.Key);
                var child = new FolderTreeNode
                {
                    Name = g.Key,
                    FullPath = childPath,
                    FileCount = g.Count(f => Path.GetDirectoryName(f.FilePath).Equals(childPath, StringComparison.OrdinalIgnoreCase)),
                    TotalFileCount = g.Count()
                };

                var deeper = g.Where(f => !Path.GetDirectoryName(f.FilePath).Equals(childPath, StringComparison.OrdinalIgnoreCase)).ToList();
                if (deeper.Count > 0)
                    BuildChildNodes(child, deeper, rootPath, foldersOnly);

                if (!foldersOnly)
                {
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

        // --- Step 2: Thumbnail Generation & Cache ---

        public string GetCacheFolder(string filePath)
        {
            return _cacheService.GetCacheFolder(filePath);
        }

        public bool IsCacheValid(string filePath)
        {
            return _cacheService.IsCacheValid(filePath);
        }

        public string GetCachedThumbnailPath(string filePath, int shapeId)
        {
            string path = Path.Combine(GetCacheFolder(filePath), $"shp_{shapeId}.png");
            return File.Exists(path) ? path : null;
        }

        public BitmapImage LoadCachedThumbnail(string filePath, int shapeId)
        {
            string thumbPath = GetCachedThumbnailPath(filePath, shapeId);
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
                System.Diagnostics.Debug.WriteLine($"ShapeLibraryManager.LoadCachedThumbnail error: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Reads counting of shapes from a manifest if it exists in the cache, or returns 0.
        /// </summary>
        public List<ShapeItem> GetCachedShapeItems(string filePath)
        {
            try
            {
                string cacheDir = GetCacheFolder(filePath);
                string manifest = Path.Combine(cacheDir, "manifest.json");
                if (File.Exists(manifest))
                {
                    return JsonService.Load<List<ShapeItem>>(manifest) ?? new List<ShapeItem>();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ShapeLibraryManager.GetCachedShapeItems error: {ex.Message}");
            }
            return new List<ShapeItem>();
        }

        /// <summary>
        /// Scans a PPTX for shapes, exports each as a PNG, and saves a manifest.
        /// </summary>
        public async Task GenerateThumbnailCache(string filePath, Action<ShapeItem, BitmapImage> onShapeReady = null)
        {
            // Ensure we are on the UI thread for COM operations
            if (!System.Windows.Application.Current.Dispatcher.CheckAccess())
            {
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(async () => await GenerateThumbnailCache(filePath, onShapeReady));
                return;
            }

            Presentation pres = null;
            try
            {
                string cacheDir = GetCacheFolder(filePath);
                Directory.CreateDirectory(cacheDir);

                var app  = _manager.GetApplication();
                pres = app.Presentations.Open(
                    filePath,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoFalse);

                var shapeItems = new List<ShapeItem>();
                int totalProcessed = 0;

                foreach (Slide slide in pres.Slides)
                {
                    try
                    {
                        foreach (Shape shape in slide.Shapes)
                        {
                            try
                            {
                                // Skip placeholders, backgrounds, empty text boxes just to be safe
                                if (shape.Type == MsoShapeType.msoPlaceholder) continue;

                                int shapeId = shape.Id;
                                string shapeName = string.IsNullOrWhiteSpace(shape.Name) ? $"Shape {shapeId}" : shape.Name;

                                string uniqueIdStr = $"{slide.SlideIndex}_{shapeId}";
                                int uniqueIdHash = Math.Abs(uniqueIdStr.GetHashCode());

                                string thumbPath = Path.Combine(cacheDir, $"shp_{uniqueIdHash}.png");

                                try
                                {
                                    shape.Export(thumbPath, PpShapeFormat.ppShapeFormatPNG, 300, 300, PpExportMode.ppRelativeToSlide);
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Shape export failed: {ex.Message}");
                                    continue; 
                                }

                                if (File.Exists(thumbPath))
                                {
                                    var item = new ShapeItem
                                    {
                                        IdStr = uniqueIdStr,
                                        UniqueId = uniqueIdHash,
                                        Title = shapeName,
                                        SlideIndex = slide.SlideIndex,
                                        OriginalShapeId = shapeId
                                    };
                                    shapeItems.Add(item);

                                    if (onShapeReady != null)
                                    {
                                        var img = new BitmapImage();
                                        img.BeginInit();
                                        img.CacheOption = BitmapCacheOption.OnLoad;
                                        img.UriSource   = new Uri(thumbPath);
                                        img.EndInit();
                                        img.Freeze();
                                        onShapeReady(item, img);
                                    }
                                }
                                
                                totalProcessed++;
                                // Yield periodically to keep UI responsive
                                if (totalProcessed % 5 == 0) await Task.Delay(1);
                            }
                            finally
                            {
                                if (shape != null)
                                {
                                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(shape); } catch { }
                                }
                            }
                        }
                    }
                    finally
                    {
                        if (slide != null)
                        {
                            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(slide); } catch { }
                        }
                    }
                }

                string manifest = Path.Combine(cacheDir, "manifest.json");
                JsonService.Save(shapeItems, manifest);

                _cacheService.RegisterSourcePath(filePath, cacheDir);
                _cacheService.WriteTimestamp(filePath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Shape GenerateThumbnailCache failed: " + ex.Message);
            }
            finally
            {
                if (pres != null)
                {
                    try { pres.Close(); } catch { }
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(pres); } catch { }
                }
            }
        }

        // --- Step 3: Insertion Logic ---

        public bool InsertShape(string sourcePath, int slideIndex, int shapeId)
        {
            // Ensure STA thread for COM insertion
            if (!System.Windows.Application.Current.Dispatcher.CheckAccess())
            {
                return System.Windows.Application.Current.Dispatcher.Invoke(() => InsertShape(sourcePath, slideIndex, shapeId));
            }

            Presentation sourcePres = null;
            try
            {
                var app = _manager.GetApplication();
                if (app.ActivePresentation == null) return false;

                try { app.StartNewUndoEntry(); } catch { }

                Slide activeSlide = null;
                try
                {
                    if (app.ActiveWindow != null && app.ActiveWindow.ViewType == PpViewType.ppViewNormal)
                    {
                        activeSlide = (Slide)app.ActiveWindow.View.Slide;
                    }
                }
                catch { }

                if (activeSlide == null) return false;

                sourcePres = app.Presentations.Open(
                    sourcePath,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoFalse);

                try
                {
                    Slide sourceSlide = sourcePres.Slides[slideIndex];
                    Shape shapeToCopy = null;

                    try
                    {
                        foreach(Shape s in sourceSlide.Shapes)
                        {
                            if (s.Id == shapeId)
                            {
                                shapeToCopy = s;
                                break;
                            }
                        }

                        if (shapeToCopy != null)
                        {
                            shapeToCopy.Copy();
                            var pastedRange = activeSlide.Shapes.Paste();
                            if (pastedRange.Count > 0)
                            {
                                Shape pastedShape = pastedRange[1];
                                pastedShape.Select(MsoTriState.msoTrue);
                                
                                try 
                                {
                                    string metaPath = sourcePath + ".meta.json";
                                    if (File.Exists(metaPath)) 
                                    {
                                        var metaList = JsonService.Load<List<ShapeMetadataItem>>(metaPath);
                                        if (metaList != null) 
                                        {
                                            var meta = metaList.FirstOrDefault(m => m.OriginalShapeId == shapeId);
                                            if (meta != null && meta.SlideWidth > 0 && meta.SlideHeight > 0) 
                                            {
                                                float relLeft = meta.Left / meta.SlideWidth;
                                                float relTop = meta.Top / meta.SlideHeight;
                                                
                                                pastedShape.Left = relLeft * app.ActivePresentation.PageSetup.SlideWidth;
                                                pastedShape.Top = relTop * app.ActivePresentation.PageSetup.SlideHeight;
                                            }
                                        }
                                    }
                                } 
                                catch { }
                            }
                            return true;
                        }
                        return false;
                    }
                    finally
                    {
                        if (shapeToCopy != null)
                        {
                            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(shapeToCopy); } catch { }
                        }
                    }
                }
                finally
                {
                    sourcePres.Close();
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(sourcePres); } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Insert Shape Error: " + ex.Message);
                return false;
            }
        }
    }
}