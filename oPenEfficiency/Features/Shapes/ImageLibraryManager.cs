using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Slide = Microsoft.Office.Interop.PowerPoint.Slide;

namespace oPenEfficiency.Features
{
    public class ImageLibraryManager
    {
        private readonly PowerPointManager _manager;
        private const string ConfigFile = "image_library_folders.json";

        public ImageLibraryManager(PowerPointManager manager)
        {
            _manager = manager;
        }

        // --- Step 1: Library Source Management ---

        public List<string> GetLibraryFolders()
        {
            var list = JsonService.Load<List<string>>(ConfigFile) ?? new List<string>();

            // Always add the user's personal "Favorites" folder
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string myAssetsPath = Path.Combine(appDataPath, "oPenEfficiency", "favorites", "images");
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

        public async Task<List<ImageFile>> ScanLibrariesAsync()
        {
            return await Task.Run(() =>
            {
                var scannedFiles = new List<ImageFile>();
                var folders = GetLibraryFolders();

                string[] validExtensions = { ".png", ".jpg", ".jpeg", ".svg" };

                foreach (var folder in folders)
                {
                    if (Directory.Exists(folder))
                    {
                        try
                        {
                            var files = oPenEfficiency.Utils.CommonUtils.SafeGetFiles(folder, "*.*")
                                                 .Where(f => validExtensions.Contains(Path.GetExtension(f).ToLower()))
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

                                                     scannedFiles.Add(new ImageFile 
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
                            System.Diagnostics.Debug.WriteLine($"ImageLibrary: Error scanning folder: {ex.Message}");
                            /* Access denied or other error */
                        }
                    }
                }
                return scannedFiles;
            });
        }

        public List<FolderTreeNode> BuildFolderTree(List<ImageFile> files, bool foldersOnly = false)
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
                    // Add files directly in the root folder as leaf nodes
                    var directFiles = rootGroup.Where(f => string.IsNullOrEmpty(f.SubFolder)).ToList();
                    foreach (var file in directFiles)
                    {
                        rootNode.Children.Add(new FolderTreeNode
                        {
                            Name = file.FileName,
                            FullPath = file.FilePath,
                            IsFile = true,
                            IsCached = true, // Images are the source, so they are always 'ready'
                            FilePath = file.FilePath
                        });
                    }
                }

                rootNode.TotalFileCount = rootGroup.Count();
                roots.Add(rootNode);
            }

            return roots;
        }

        private void BuildChildNodes(FolderTreeNode parent, List<ImageFile> files, string rootPath, bool foldersOnly = false)
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
                            IsCached = true,
                            FilePath = file.FilePath
                        });
                    }
                }

                parent.Children.Add(child);
            }
        }

        // --- Step 2: Thumbnail Generation ---
        
        public async Task<System.Windows.Media.ImageSource> GetThumbnailAsync(string filePath)
        {
            if (filePath.EndsWith(".svg", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var geom = oPenEfficiency.Utils.SvgParser.ParseLucideSvg(filePath, out bool isFilled);
                    if (geom != null)
                    {
                        var drawing = new System.Windows.Media.GeometryDrawing(
                            isFilled ? System.Windows.Media.Brushes.Gray : null,
                            isFilled ? null : new System.Windows.Media.Pen(System.Windows.Media.Brushes.Gray, 2),
                            geom);
                        var d = new System.Windows.Media.DrawingImage(drawing);
                        d.Freeze();
                        return (System.Windows.Media.ImageSource)d;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"ImageLibrary: Error loading SVG thumbnail: {ex.Message}");
                }
                return null;
            }

            return await Task.Run<System.Windows.Media.ImageSource>(() =>
            {
                try
                {
                    var image = new BitmapImage();
                    image.BeginInit();
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.DecodePixelWidth = 200; // Optimize memory
                    image.UriSource = new Uri(filePath);
                    image.EndInit();
                    image.Freeze(); // Allow cross thread
                    return image;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"ImageLibrary: Error loading thumbnail: {ex.Message}");
                }
                return null;
            });
        }

        // --- Step 3: Conditional Insertion Logic ---

        public bool InsertLibraryImage(string sourcePath, bool maintainAspectRatio)
        {
            try
            {
                var app = _manager.GetApplication();
                if (app.ActivePresentation == null) return false;

                try { app.StartNewUndoEntry(); }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"ImageLibrary: Error starting undo entry: {ex.Message}");
                }

                Slide activeSlide = null;
                try
                {
                    if (app.ActiveWindow != null && app.ActiveWindow.ViewType == PpViewType.ppViewNormal)
                    {
                        activeSlide = (Slide)app.ActiveWindow.View.Slide;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"ImageLibrary: Error getting active slide: {ex.Message}");
                }

                if (activeSlide == null) return false;

                // Insert in center of slide by default, or 0,0 if not provided
                Shape img = activeSlide.Shapes.AddPicture(sourcePath, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, -1, -1);
                
                if (img != null)
                {
                    if (maintainAspectRatio)
                    {
                        img.LockAspectRatio = MsoTriState.msoTrue;
                    }
                    
                    // Center the picture
                    float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                    float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
                    
                    // Don't let it be larger than the slide
                    if (img.Width > slideWidth) img.Width = slideWidth;
                    if (img.Height > slideHeight) img.Height = slideHeight;
                    
                    img.Left = (slideWidth - img.Width) / 2;
                    img.Top = (slideHeight - img.Height) / 2;
                    
                    img.Select(MsoTriState.msoTrue);
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Insert Error: " + ex.Message);
                return false;
            }
        }
    }
}