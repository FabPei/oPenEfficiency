using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Media;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Services
{
    public enum IconInsertMode
    {
        Svg,
        PowerPointShape
    }

    public class IconCatalogService
    {
        private const string ManifestRelativePath = @"Resources\IconLibrary\icon-manifest.json";
        private const string ArchiveRelativePath = @"Resources\IconLibrary\icon-library-assets.zip";
        private const string AssetsRelativePath = @"UI\Assets\IconLibrary";
        private const string FavoritesConfigFile = "icon_favorites.json";
        private const float DefaultIconSpacingRatio = 0.35f;
        private readonly string _assemblyDir;
        private readonly IIconCatalogDataSource _dataSource;
        private readonly IconManifestParser _manifestParser = new IconManifestParser();
        private readonly IconManifestCache _manifestCache = new IconManifestCache();

        public IconCatalogService()
        {
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            var uri = new UriBuilder(codeBase);
            _assemblyDir = Path.GetDirectoryName(Uri.UnescapeDataString(uri.Path));
            _dataSource = new FileIconCatalogDataSource(_assemblyDir, ManifestRelativePath);
        }

        public IconCollectionManifest LoadManifest()
        {
            try
            {
                string manifestPath = ResolveManifestPath();
                if (string.IsNullOrWhiteSpace(manifestPath) || !File.Exists(manifestPath))
                {
                    return new IconCollectionManifest();
                }

                DateTime writeTimeUtc = _dataSource.GetLastWriteTimeUtc(manifestPath);
                var manifest = BuildRuntimeManifest(_manifestCache.Get(
                    manifestPath,
                    writeTimeUtc,
                    path => _manifestParser.Parse(_dataSource.ReadAllText(path))));
                ApplyFavorites(manifest.Icons);
                return manifest;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.LoadManifest");
                return new IconCollectionManifest();
            }
        }

        private IconCollectionManifest BuildRuntimeManifest(IconCollectionManifest source)
        {
            var manifest = _manifestParser.Clone(source) ?? new IconCollectionManifest();
            MergeCustomIcons(manifest);
            return manifest;
        }

        public IReadOnlyList<IconCatalogItem> LoadIcons()
        {
            return LoadManifest().Icons;
        }

        public IReadOnlyList<IconCollectionInfo> LoadCollections()
        {
            return LoadManifest().Collections;
        }

        public IconCatalogItem GetIconById(string iconId)
        {
            return LoadIcons().FirstOrDefault(i => string.Equals(i.Id, iconId, StringComparison.OrdinalIgnoreCase));
        }

        public string ResolveIconPath(IconCatalogItem item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.RelativePath))
            {
                return null;
            }

            try
            {
                string runtimePath = Path.Combine(_assemblyDir, item.RelativePath);
                if (File.Exists(runtimePath))
                {
                    return runtimePath;
                }

                string devPath = Path.GetFullPath(Path.Combine(_assemblyDir, "..", "..", item.RelativePath));
                if (File.Exists(devPath))
                {
                    return devPath;
                }

                string extractedPath = ExtractIconFromArchive(item);
                if (File.Exists(extractedPath))
                {
                    return extractedPath;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.ResolveIconPath");
            }

            return null;
        }

        public IReadOnlyList<string> GetAvailableCollections(IEnumerable<IconCatalogItem> icons = null)
        {
            icons = icons ?? LoadIcons();
            return icons
                .Select(i => i.CollectionId)
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(v => v)
                .ToList();
        }

        public IReadOnlyList<string> GetAvailableThemes(IEnumerable<IconCatalogItem> icons = null)
        {
            icons = icons ?? LoadIcons();
            return icons
                .Select(i => i.Theme)
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(v => v)
                .ToList();
        }

        public IReadOnlyList<string> GetAvailableStyles(IEnumerable<IconCatalogItem> icons = null)
        {
            icons = icons ?? LoadIcons();
            return icons
                .Select(i => i.Style)
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(v => v)
                .ToList();
        }

        public IReadOnlyList<IconCatalogItem> FilterIcons(
            IEnumerable<IconCatalogItem> icons,
            string search = null,
            string collectionId = null,
            string theme = null,
            string style = null,
            bool favoritesOnly = false)
        {
            return new IconSearchIndex(icons ?? LoadIcons())
                .Search(search, collectionId, theme, style, favoritesOnly);
        }

        public void SetFavorite(string iconId, bool isFavorite)
        {
            if (string.IsNullOrWhiteSpace(iconId))
            {
                return;
            }

            try
            {
                var config = JsonService.Load<IconFavoritesConfig>(FavoritesConfigFile) ?? new IconFavoritesConfig();

                if (isFavorite)
                {
                    if (!config.FavoriteIconIds.Contains(iconId, StringComparer.OrdinalIgnoreCase))
                    {
                        config.FavoriteIconIds.Add(iconId);
                    }
                }
                else
                {
                    config.FavoriteIconIds = config.FavoriteIconIds
                        .Where(id => !string.Equals(id, iconId, StringComparison.OrdinalIgnoreCase))
                        .ToList();
                }

                JsonService.Save(config, FavoritesConfigFile);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.SetFavorite");
            }
        }

        private void ApplyFavorites(IEnumerable<IconCatalogItem> icons)
        {
            var favorites = GetFavoriteSet();
            foreach (var icon in icons ?? Enumerable.Empty<IconCatalogItem>())
            {
                icon.IsFavorite = favorites.Contains(icon.Id ?? string.Empty);
            }
        }

        private HashSet<string> GetFavoriteSet()
        {
            try
            {
                var config = JsonService.Load<IconFavoritesConfig>(FavoritesConfigFile) ?? new IconFavoritesConfig();
                return new HashSet<string>(config.FavoriteIconIds ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.GetFavoriteSet");
                return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        private string ResolveManifestPath()
        {
            return _dataSource.FindManifest();
        }

        public string GetAssetsRootPath()
        {
            string runtimePath = Path.Combine(_assemblyDir, AssetsRelativePath);
            if (Directory.Exists(runtimePath))
            {
                return runtimePath;
            }

            string devPath = Path.GetFullPath(Path.Combine(_assemblyDir, "..", "..", AssetsRelativePath));
            if (Directory.Exists(devPath))
            {
                return devPath;
            }

            return null;
        }

        private void MergeCustomIcons(IconCollectionManifest manifest)
        {
            if (manifest == null)
            {
                return;
            }

            try
            {
                string customRoot = GetCustomIconsRootPath();
                if (string.IsNullOrWhiteSpace(customRoot) || !Directory.Exists(customRoot))
                {
                    return;
                }

                var customIcons = DiscoverCustomIcons(customRoot);
                if (customIcons.Count == 0)
                {
                    return;
                }

                manifest.Collections = manifest.Collections ?? new List<IconCollectionInfo>();
                manifest.Icons = manifest.Icons ?? new List<IconCatalogItem>();

                manifest.Icons = manifest.Icons
                    .Where(icon => !string.Equals(icon.CollectionId, "custom", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                var existingCollection = manifest.Collections
                    .FirstOrDefault(c => string.Equals(c.Id, "custom", StringComparison.OrdinalIgnoreCase));

                var styles = customIcons
                    .Select(icon => icon.Style)
                    .Where(style => !string.IsNullOrWhiteSpace(style))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(style => style)
                    .ToList();

                if (existingCollection == null)
                {
                    manifest.Collections.Add(new IconCollectionInfo
                    {
                        Id = "custom",
                        Name = "Custom Icons",
                        Description = "Local project icons from the custom folder.",
                        License = "Project Local",
                        LicenseUrl = string.Empty,
                        SourceUrl = string.Empty,
                        Styles = styles
                    });
                }
                else
                {
                    existingCollection.Name = "Custom Icons";
                    existingCollection.Description = "Local project icons from the custom folder.";
                    existingCollection.License = "Project Local";
                    existingCollection.Styles = styles;
                }

                manifest.Icons.AddRange(customIcons);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.MergeCustomIcons");
            }
        }

        private string GetCustomIconsRootPath()
        {
            string assetsRoot = GetAssetsRootPath();
            if (string.IsNullOrWhiteSpace(assetsRoot))
            {
                return null;
            }

            return Path.Combine(assetsRoot, "custom");
        }

        private List<IconCatalogItem> DiscoverCustomIcons(string customRoot)
        {
            var results = new List<IconCatalogItem>();
            foreach (string filePath in Directory.EnumerateFiles(customRoot, "*.svg", SearchOption.AllDirectories))
            {
                try
                {
                    var icon = CreateCustomIconItem(customRoot, filePath);
                    if (icon != null)
                    {
                        results.Add(icon);
                    }
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "IconCatalogService.DiscoverCustomIcons");
                }
            }

            return results
                .OrderBy(icon => icon.Name)
                .ThenBy(icon => icon.Style)
                .ToList();
        }

        private static IconCatalogItem CreateCustomIconItem(string customRoot, string filePath)
        {
            if (string.IsNullOrWhiteSpace(customRoot) || string.IsNullOrWhiteSpace(filePath))
            {
                return null;
            }

            string relativeWithinCustom = GetRelativePathSafe(customRoot, filePath);
            if (relativeWithinCustom.StartsWith("..", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            string normalizedRelative = relativeWithinCustom.Replace(Path.DirectorySeparatorChar, '/');
            string[] segments = normalizedRelative.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string fileName = Path.GetFileNameWithoutExtension(filePath);

            string style = DetectCustomStyle(fileName);
            string theme = DetectCustomTheme(segments);
            string displayName = BuildCustomDisplayName(fileName);
            var tags = BuildCustomTags(fileName, segments, style, theme);
            string idSlug = SanitizeSlug(normalizedRelative.Replace(".svg", string.Empty));

            return new IconCatalogItem
            {
                Id = $"custom.{idSlug}",
                Name = displayName,
                CollectionId = "custom",
                Theme = theme,
                Style = style,
                RelativePath = $@"UI\Assets\IconLibrary\custom\{relativeWithinCustom.Replace('/', '\\')}",
                License = "Project Local",
                SourceUrl = string.Empty,
                Categories = new List<string> { theme },
                Tags = tags
            };
        }

        private static string DetectCustomStyle(string fileName)
        {
            string normalized = (fileName ?? string.Empty).ToLowerInvariant();
            if (normalized.Contains("outline"))
            {
                return "outline";
            }

            if (normalized.Contains("filled") || normalized.Contains("fill"))
            {
                return "filled";
            }

            if (normalized.Contains("regular"))
            {
                return "regular";
            }

            return "custom";
        }

        private static string DetectCustomTheme(string[] segments)
        {
            if (segments == null || segments.Length <= 1)
            {
                return "custom";
            }

            return ToTagValue(Path.GetFileNameWithoutExtension(segments[0]));
        }

        private static string BuildCustomDisplayName(string fileName)
        {
            string cleaned = fileName ?? string.Empty;
            cleaned = Regex.Replace(cleaned, @"(?i)[_-](outline|filled|fill|regular)\b", string.Empty);
            cleaned = Regex.Replace(cleaned, @"(?i)[_-]\d+x\d+\b", string.Empty);
            cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim(' ', '_', '-');
            return ToDisplayLabelSafe(cleaned);
        }

        private static List<string> BuildCustomTags(string fileName, string[] segments, string style, string theme)
        {
            var tags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (string segment in segments ?? Array.Empty<string>())
            {
                string withoutExtension = Path.GetFileNameWithoutExtension(segment);
                foreach (string token in Regex.Split(withoutExtension.Replace('-', ' ').Replace('_', ' '), @"\s+"))
                {
                    string tag = ToTagValue(token);
                    if (!string.IsNullOrWhiteSpace(tag))
                    {
                        tags.Add(tag);
                    }
                }
            }

            foreach (string token in Regex.Split((fileName ?? string.Empty).Replace('-', ' ').Replace('_', ' '), @"\s+"))
            {
                string tag = ToTagValue(token);
                if (!string.IsNullOrWhiteSpace(tag))
                {
                    tags.Add(tag);
                }
            }

            if (!string.IsNullOrWhiteSpace(style))
            {
                tags.Add(style);
            }

            if (!string.IsNullOrWhiteSpace(theme))
            {
                tags.Add(theme);
            }

            tags.Remove("svg");
            return tags.OrderBy(tag => tag).ToList();
        }

        private static string ToTagValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            string normalized = Regex.Replace(value.Trim().ToLowerInvariant(), @"[^a-z0-9äöüß]+", " ").Trim();
            return string.IsNullOrWhiteSpace(normalized) ? null : normalized;
        }

        private static string SanitizeSlug(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "icon";
            }

            string slug = value.Replace('\\', '.').Replace('/', '.').Replace(' ', '-').ToLowerInvariant();
            slug = Regex.Replace(slug, @"[^a-z0-9\.-]+", "-");
            slug = Regex.Replace(slug, @"-+", "-").Trim('-', '.');
            return string.IsNullOrWhiteSpace(slug) ? "icon" : slug;
        }

        private static string GetRelativePathSafe(string rootPath, string fullPath)
        {
            if (string.IsNullOrWhiteSpace(rootPath) || string.IsNullOrWhiteSpace(fullPath))
            {
                return string.Empty;
            }

            string normalizedRoot = Path.GetFullPath(rootPath)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            string normalizedFull = Path.GetFullPath(fullPath);

            if (!normalizedFull.StartsWith(normalizedRoot, StringComparison.OrdinalIgnoreCase))
            {
                return fullPath;
            }

            return normalizedFull.Substring(normalizedRoot.Length);
        }

        private static string ToDisplayLabelSafe(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "Unknown";
            }

            var parts = value.Replace("-", " ").Replace("_", " ")
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(" ", parts.Select(p => char.ToUpperInvariant(p[0]) + p.Substring(1)));
        }

        private string ExtractIconFromArchive(IconCatalogItem item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.RelativePath))
            {
                return null;
            }

            try
            {
                string archivePath = ResolveArchivePath();
                if (string.IsNullOrWhiteSpace(archivePath) || !File.Exists(archivePath))
                {
                    return null;
                }

                string relativeAssetPath = item.RelativePath.Replace('\\', '/');
                const string archivePrefix = "UI/Assets/IconLibrary/";
                if (!relativeAssetPath.StartsWith(archivePrefix, StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                string archiveEntryName = relativeAssetPath.Substring(archivePrefix.Length);
                string extractRoot = Path.Combine(Path.GetTempPath(), "oPenEfficiency", "IconLibraryCache");
                string extractedPath = Path.Combine(extractRoot, archiveEntryName.Replace('/', Path.DirectorySeparatorChar));
                string extractedDir = Path.GetDirectoryName(extractedPath);
                if (!string.IsNullOrWhiteSpace(extractedDir))
                {
                    Directory.CreateDirectory(extractedDir);
                }

                DateTime archiveWriteUtc = File.GetLastWriteTimeUtc(archivePath);
                if (File.Exists(extractedPath) && File.GetLastWriteTimeUtc(extractedPath) >= archiveWriteUtc)
                {
                    return extractedPath;
                }

                using (var archive = ZipFile.OpenRead(archivePath))
                {
                    var entry = archive.GetEntry(archiveEntryName);
                    if (entry == null)
                    {
                        return null;
                    }

                    entry.ExtractToFile(extractedPath, true);
                }

                File.SetLastWriteTimeUtc(extractedPath, archiveWriteUtc);
                return extractedPath;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.ExtractIconFromArchive");
                return null;
            }
        }

        public bool InsertIconIntoActiveSlide(
            PowerPointManager manager,
            IconCatalogItem item,
            string hexColor,
            IconInsertMode insertMode = IconInsertMode.Svg,
            float desiredSize = 48f)
        {
            return InsertIconsIntoActiveSlide(manager, new[] { item }, hexColor, insertMode, desiredSize);
        }

        public bool InsertIconsIntoActiveSlide(
            PowerPointManager manager,
            IEnumerable<IconCatalogItem> items,
            string hexColor,
            IconInsertMode insertMode = IconInsertMode.Svg,
            float desiredSize = 48f)
        {
            if (manager == null || items == null)
            {
                return false;
            }

            try
            {
                var app = manager.GetApplication();
                if (app?.ActivePresentation == null || app.ActiveWindow == null)
                {
                    return false;
                }

                PowerPoint.Slide activeSlide = null;
                try
                {
                    if (app.ActiveWindow.ViewType == PowerPoint.PpViewType.ppViewNormal)
                    {
                        activeSlide = (PowerPoint.Slide)app.ActiveWindow.View.Slide;
                    }
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "IconCatalogService.InsertIconIntoActiveSlide.GetSlide");
                }

                if (activeSlide == null)
                {
                    return false;
                }

                var iconItems = (items ?? Enumerable.Empty<IconCatalogItem>())
                    .Where(i => i != null)
                    .ToList();
                if (iconItems.Count == 0)
                {
                    return false;
                }

                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
                int columns = (int)Math.Ceiling(Math.Sqrt(iconItems.Count));
                int rows = (int)Math.Ceiling((double)iconItems.Count / columns);
                float spacing = Math.Max(10f, desiredSize * DefaultIconSpacingRatio);
                float blockWidth = (columns * desiredSize) + ((columns - 1) * spacing);
                float blockHeight = (rows * desiredSize) + ((rows - 1) * spacing);
                float startLeft = (slideWidth - blockWidth) / 2f;
                float startTop = (slideHeight - blockHeight) / 2f;

                var insertedShapeIds = new List<int>();
                var tempFiles = new List<string>();
                string normalizedColor = NormalizeHexColor(hexColor);

                try
                {
                    for (int index = 0; index < iconItems.Count; index++)
                    {
                        var iconItem = iconItems[index];
                        string sourcePath = ResolveIconPath(iconItem);
                        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
                        {
                            continue;
                        }

                        int column = index % columns;
                        int row = index / columns;
                        float left = startLeft + (column * (desiredSize + spacing));
                        float top = startTop + (row * (desiredSize + spacing));

                        PowerPoint.Shape shape = null;
                        if (insertMode == IconInsertMode.PowerPointShape)
                        {
                            bool shouldFallbackToSvg = ShouldFallbackToSvgInsert(iconItem, sourcePath);
                            if (!shouldFallbackToSvg)
                            {
                                shape = CreateNativeIconShape(activeSlide, iconItem, sourcePath, normalizedColor, left, top, desiredSize);
                            }

                            if (shape == null)
                            {
                                string tempFile = CreateColorizedTempSvg(sourcePath, normalizedColor);
                                if (string.IsNullOrWhiteSpace(tempFile) || !File.Exists(tempFile))
                                {
                                    continue;
                                }

                                tempFiles.Add(tempFile);

                                shape = InsertSvgAsPowerPointShape(app, activeSlide, tempFile, left, top, desiredSize)
                                    ?? activeSlide.Shapes.AddPicture(
                                        tempFile,
                                        Office.MsoTriState.msoFalse,
                                        Office.MsoTriState.msoTrue,
                                        left,
                                        top,
                                        desiredSize,
                                        desiredSize);
                            }
                        }
                        else
                        {
                            string tempFile = CreateColorizedTempSvg(sourcePath, normalizedColor);
                            if (string.IsNullOrWhiteSpace(tempFile) || !File.Exists(tempFile))
                            {
                                continue;
                            }

                            tempFiles.Add(tempFile);

                            shape = activeSlide.Shapes.AddPicture(
                                tempFile,
                                Office.MsoTriState.msoFalse,
                                Office.MsoTriState.msoTrue,
                                left,
                                top,
                                desiredSize,
                                desiredSize);
                        }

                        if (shape != null)
                        {
                            insertedShapeIds.Add(shape.Id);
                        }
                    }
                }
                finally
                {
                    foreach (string tempFile in tempFiles)
                    {
                        try
                        {
                            if (!string.IsNullOrWhiteSpace(tempFile) && File.Exists(tempFile))
                            {
                                File.Delete(tempFile);
                            }
                        }
                        catch (Exception ex)
                        {
                            ExceptionLogger.Log(ex, "IconCatalogService.InsertIconsIntoActiveSlide.DeleteTemp");
                        }
                    }
                }

                if (insertedShapeIds.Count == 0)
                {
                    return false;
                }

                TrySelectInsertedShapesById(activeSlide, insertedShapeIds);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.InsertIconsIntoActiveSlide");
                return false;
            }
        }

        private PowerPoint.Shape CreateNativeIconShape(
            PowerPoint.Slide slide,
            IconCatalogItem iconItem,
            string sourcePath,
            string hexColor,
            float left,
            float top,
            float desiredSize)
        {
            var createdShapes = new List<PowerPoint.Shape>();
            try
            {
                var geometry = SvgParser.ParseLucideSvg(sourcePath, out bool isFilled);
                if (geometry == null)
                {
                    return null;
                }

                var flattened = geometry.GetFlattenedPathGeometry();
                if (flattened == null || flattened.Figures.Count == 0)
                {
                    return null;
                }

                Rect bounds = flattened.Bounds;
                if (bounds.Width <= 0 || bounds.Height <= 0)
                {
                    return null;
                }

                float scale = (float)Math.Min(desiredSize / bounds.Width, desiredSize / bounds.Height);
                float targetWidth = (float)(bounds.Width * scale);
                float targetHeight = (float)(bounds.Height * scale);
                float offsetX = left + ((desiredSize - targetWidth) / 2f);
                float offsetY = top + ((desiredSize - targetHeight) / 2f);
                bool useGroupedFilledFigures = IsPhosphorRegular(iconItem);
                bool renderAsFilled = isFilled || useGroupedFilledFigures;
                bool preferPolyline = !renderAsFilled;

                foreach (PathFigure figure in flattened.Figures)
                {
                    var createdShape = CreateShapeForFigure(slide, figure, bounds, scale, offsetX, offsetY, preferPolyline);
                    if (createdShape == null)
                    {
                        continue;
                    }

                    ApplyNativeShapeFormatting(createdShape, renderAsFilled, hexColor, desiredSize);
                    createdShapes.Add(createdShape);
                }

                if (createdShapes.Count == 0)
                {
                    return null;
                }

                if (createdShapes.Count == 1)
                {
                    return createdShapes[0];
                }

                var mergedShape = (renderAsFilled && !useGroupedFilledFigures)
                    ? TryCombineShapes(slide, createdShapes)
                    : TryGroupShapes(slide, createdShapes);

                if (mergedShape != null)
                {
                    return mergedShape;
                }

                return createdShapes[0];
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.CreateNativeIconShape");
                CleanupShapes(createdShapes);
                return null;
            }
        }

        private static PowerPoint.Shape InsertSvgAsPowerPointShape(
            PowerPoint.Application app,
            PowerPoint.Slide slide,
            string svgPath,
            float left,
            float top,
            float desiredSize)
        {
            PowerPoint.Shape importedShape = null;
            try
            {
                importedShape = slide.Shapes.AddPicture(
                    svgPath,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue,
                    left,
                    top,
                    desiredSize,
                    desiredSize);

                return TryConvertImportedSvgViaOffice(app, slide, importedShape)
                    ?? TryExpandImportedSvg(importedShape, slide)
                    ?? importedShape;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.InsertSvgAsPowerPointShape");
                return importedShape;
            }
        }

        private static PowerPoint.Shape TryConvertImportedSvgViaOffice(
            PowerPoint.Application app,
            PowerPoint.Slide slide,
            PowerPoint.Shape importedShape)
        {
            if (app == null || slide == null || importedShape == null)
            {
                return null;
            }

            try
            {
                bool canConvert;
                try
                {
                    canConvert = app.CommandBars.GetEnabledMso("SVGEdit");
                }
                catch
                {
                    canConvert = true;
                }

                if (!canConvert)
                {
                    ExceptionLogger.Log("SVGEdit is disabled for the current selection.", "IconCatalogService.TryConvertImportedSvgViaOffice");
                    return null;
                }

                var beforeIds = GetShapeIds(slide);
                ExceptionLogger.Log(
                    $"Before SVGEdit: shapeCount={slide.Shapes.Count}, importedShapeId={importedShape.Id}, importedType={importedShape.Type}, importedName={importedShape.Name}",
                    "IconCatalogService.TryConvertImportedSvgViaOffice");
                importedShape.Select();
                app.CommandBars.ExecuteMso("SVGEdit");

                var selection = app.ActiveWindow?.Selection;
                string selectionType = selection == null ? "null" : selection.Type.ToString();
                int selectionCount = 0;
                try
                {
                    if (selection?.ShapeRange != null)
                    {
                        selectionCount = selection.ShapeRange.Count;
                    }
                }
                catch (Exception ex)
                {
                    ExceptionPolicy.ReportOptionalComFailure(ex, "IconCatalogService.TryConvertImportedSvgViaOffice.SelectionCount");
                }

                ExceptionLogger.Log(
                    $"After SVGEdit: shapeCount={slide.Shapes.Count}, selectionType={selectionType}, selectionCount={selectionCount}",
                    "IconCatalogService.TryConvertImportedSvgViaOffice");

                if (selection?.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange != null && selection.ShapeRange.Count > 0)
                {
                    var shapeNames = new List<string>();
                    for (int i = 1; i <= selection.ShapeRange.Count; i++)
                    {
                        if (selection.ShapeRange[i] != null)
                        {
                            shapeNames.Add(selection.ShapeRange[i].Name);
                            ExceptionLogger.Log(
                                $"Selected shape after SVGEdit: name={selection.ShapeRange[i].Name}, id={selection.ShapeRange[i].Id}, type={selection.ShapeRange[i].Type}",
                                "IconCatalogService.TryConvertImportedSvgViaOffice");
                        }
                    }

                    if (shapeNames.Count == 1)
                    {
                        return slide.Shapes.Range(new[] { shapeNames[0] })[1];
                    }

                    if (shapeNames.Count > 1)
                    {
                        return slide.Shapes.Range(shapeNames.ToArray()).Group();
                    }
                }

                var newestShape = FindNewestShape(slide, beforeIds);
                if (newestShape != null)
                {
                    ExceptionLogger.Log(
                        $"Newest shape fallback: name={newestShape.Name}, id={newestShape.Id}, type={newestShape.Type}",
                        "IconCatalogService.TryConvertImportedSvgViaOffice");
                }

                return newestShape;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.TryConvertImportedSvgViaOffice");
                return null;
            }
        }

        private static PowerPoint.Shape TryExpandImportedSvg(PowerPoint.Shape importedShape, PowerPoint.Slide slide)
        {
            if (importedShape == null || slide == null)
            {
                return null;
            }

            try
            {
                var currentRange = importedShape.Ungroup();
                if (currentRange == null || currentRange.Count == 0)
                {
                    return null;
                }

                for (int pass = 0; pass < 2; pass++)
                {
                    bool expandedNestedGroup = false;
                    var nestedNames = new List<string>();
                    for (int i = 1; i <= currentRange.Count; i++)
                    {
                        var currentShape = currentRange[i];
                        if (currentShape == null)
                        {
                            continue;
                        }

                        if (currentShape.Type == Office.MsoShapeType.msoGroup)
                        {
                            var nestedRange = currentShape.Ungroup();
                            if (nestedRange != null && nestedRange.Count > 0)
                            {
                                for (int j = 1; j <= nestedRange.Count; j++)
                                {
                                    if (nestedRange[j] != null)
                                    {
                                        nestedNames.Add(nestedRange[j].Name);
                                    }
                                }

                                expandedNestedGroup = true;
                            }
                        }
                        else
                        {
                            nestedNames.Add(currentShape.Name);
                        }
                    }

                    if (!expandedNestedGroup)
                    {
                        break;
                    }

                    if (nestedNames.Count == 0)
                    {
                        return null;
                    }

                    currentRange = slide.Shapes.Range(nestedNames.ToArray());
                }

                var shapeNames = new List<string>();
                for (int i = 1; i <= currentRange.Count; i++)
                {
                    if (currentRange[i] != null)
                    {
                        shapeNames.Add(currentRange[i].Name);
                    }
                }

                if (shapeNames.Count == 0)
                {
                    return null;
                }

                if (shapeNames.Count == 1)
                {
                    return slide.Shapes.Range(new[] { shapeNames[0] })[1];
                }

                return slide.Shapes.Range(shapeNames.ToArray()).Group();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.TryExpandImportedSvg");
                return null;
            }
        }

        private static PowerPoint.Shape CreateShapeForFigure(
            PowerPoint.Slide slide,
            PathFigure figure,
            Rect bounds,
            float scale,
            float offsetX,
            float offsetY,
            bool preferPolyline)
        {
            try
            {
                var points = new List<System.Windows.Point> { figure.StartPoint };
                foreach (PathSegment segment in figure.Segments)
                {
                    if (segment is PolyLineSegment polyLineSegment)
                    {
                        points.AddRange(polyLineSegment.Points);
                    }
                    else if (segment is LineSegment lineSegment)
                    {
                        points.Add(lineSegment.Point);
                    }
                }

                if (points.Count < 2)
                {
                    return null;
                }

                if (preferPolyline || !figure.IsClosed)
                {
                    return CreatePolylineShape(slide, points, bounds, scale, offsetX, offsetY);
                }

                float startX = TransformCoordinate(points[0].X, bounds.Left, scale, offsetX);
                float startY = TransformCoordinate(points[0].Y, bounds.Top, scale, offsetY);
                var builder = slide.Shapes.BuildFreeform(Office.MsoEditingType.msoEditingCorner, startX, startY);

                for (int i = 1; i < points.Count; i++)
                {
                    float x = TransformCoordinate(points[i].X, bounds.Left, scale, offsetX);
                    float y = TransformCoordinate(points[i].Y, bounds.Top, scale, offsetY);
                    builder.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, x, y);
                }

                if (figure.IsClosed)
                {
                    var lastPoint = points[points.Count - 1];
                    if (!ArePointsClose(lastPoint, points[0]))
                    {
                        builder.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, startX, startY);
                    }
                }

                return builder.ConvertToShape();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.CreateShapeForFigure");
                return null;
            }
        }

        private static PowerPoint.Shape CreatePolylineShape(
            PowerPoint.Slide slide,
            IReadOnlyList<System.Windows.Point> points,
            Rect bounds,
            float scale,
            float offsetX,
            float offsetY)
        {
            try
            {
                var safeArray = new float[points.Count, 2];
                for (int i = 0; i < points.Count; i++)
                {
                    safeArray[i, 0] = TransformCoordinate(points[i].X, bounds.Left, scale, offsetX);
                    safeArray[i, 1] = TransformCoordinate(points[i].Y, bounds.Top, scale, offsetY);
                }

                return slide.Shapes.AddPolyline(safeArray);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.CreatePolylineShape");
                return null;
            }
        }

        private static PowerPoint.Shape TryCombineShapes(PowerPoint.Slide slide, IReadOnlyList<PowerPoint.Shape> shapes)
        {
            try
            {
                var idsBeforeMerge = GetShapeIds(slide);
                var names = shapes
                    .Where(shape => shape != null)
                    .Select(shape => shape.Name)
                    .ToArray();

                if (names.Length == 0)
                {
                    return null;
                }

                if (names.Length == 1)
                {
                    return slide.Shapes.Range(new[] { names[0] })[1];
                }

                var range = slide.Shapes.Range(names);
                var primaryShape = slide.Shapes.Range(new[] { names[0] })[1];
                int primaryId = primaryShape.Id;
                range.MergeShapes(Office.MsoMergeCmd.msoMergeCombine, primaryShape);

                var mergedShape = FindShapeById(slide, primaryId);
                if (mergedShape != null)
                {
                    return mergedShape;
                }

                return FindNewestShape(slide, idsBeforeMerge);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.TryCombineShapes");
                return null;
            }
        }

        private static PowerPoint.Shape TryGroupShapes(PowerPoint.Slide slide, IReadOnlyList<PowerPoint.Shape> shapes)
        {
            try
            {
                var names = shapes
                    .Where(shape => shape != null)
                    .Select(shape => shape.Name)
                    .ToArray();

                if (names.Length == 0)
                {
                    return null;
                }

                if (names.Length == 1)
                {
                    return slide.Shapes.Range(new[] { names[0] })[1];
                }

                return slide.Shapes.Range(names).Group();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.TryGroupShapes");
                return null;
            }
        }

        private static void CleanupShapes(IEnumerable<PowerPoint.Shape> shapes)
        {
            foreach (var shape in shapes ?? Enumerable.Empty<PowerPoint.Shape>())
            {
                if (shape == null)
                {
                    continue;
                }

                try
                {
                    shape.Delete();
                }
                catch (Exception ex)
                {
                    ExceptionPolicy.ReportOptionalComFailure(ex, "IconCatalogService.CleanupShapes.Delete");
                }
            }
        }

        private static void ApplyNativeShapeFormatting(PowerPoint.Shape shape, bool isFilled, string hexColor, float desiredSize)
        {
            if (shape == null)
            {
                return;
            }

            int rgb = HexToOfficeRgb(hexColor);
            if (isFilled)
            {
                shape.Fill.Visible = Office.MsoTriState.msoTrue;
                shape.Fill.ForeColor.RGB = rgb;
                shape.Line.Visible = Office.MsoTriState.msoFalse;
            }
            else
            {
                shape.Fill.Visible = Office.MsoTriState.msoFalse;
                shape.Line.Visible = Office.MsoTriState.msoTrue;
                shape.Line.ForeColor.RGB = rgb;
                shape.Line.Weight = Math.Max(1.25f, desiredSize / 18f);
            }
        }

        private static HashSet<int> GetShapeIds(PowerPoint.Slide slide)
        {
            var ids = new HashSet<int>();
            if (slide?.Shapes == null)
            {
                return ids;
            }

            for (int i = 1; i <= slide.Shapes.Count; i++)
            {
                try
                {
                    ids.Add(slide.Shapes[i].Id);
                }
                catch (Exception ex)
                {
                    ExceptionPolicy.ReportOptionalComFailure(ex, "IconCatalogService.GetShapeIds");
                }
            }

            return ids;
        }

        private static PowerPoint.Shape FindNewestShape(PowerPoint.Slide slide, HashSet<int> previousIds)
        {
            if (slide?.Shapes == null)
            {
                return null;
            }

            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                try
                {
                    var shape = slide.Shapes[i];
                    if (!previousIds.Contains(shape.Id))
                    {
                        return shape;
                    }
                }
                catch (Exception ex)
                {
                    ExceptionPolicy.ReportOptionalComFailure(ex, "IconCatalogService.FindNewestShape");
                }
            }

            return null;
        }

        private static PowerPoint.Shape FindShapeById(PowerPoint.Slide slide, int shapeId)
        {
            if (slide?.Shapes == null)
            {
                return null;
            }

            for (int i = 1; i <= slide.Shapes.Count; i++)
            {
                try
                {
                    var shape = slide.Shapes[i];
                    if (shape.Id == shapeId)
                    {
                        return shape;
                    }
                }
                catch (Exception ex)
                {
                    ExceptionPolicy.ReportOptionalComFailure(ex, "IconCatalogService.FindShapeById");
                }
            }

            return null;
        }

        private static bool ShouldFallbackToSvgInsert(IconCatalogItem iconItem, string sourcePath)
        {
            if (iconItem == null || string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
            {
                return false;
            }

            string collectionId = iconItem.CollectionId ?? string.Empty;
            string style = iconItem.Style ?? string.Empty;

            if (IsPhosphorRegular(iconItem))
            {
                return false;
            }

            if (string.Equals(collectionId, "phosphor", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(collectionId, "heroicons", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(collectionId, "iconoir", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(style, "solid", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(collectionId, "tabler", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(style, "filled", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            try
            {
                string svg = File.ReadAllText(sourcePath);
                if (svg.IndexOf("fill-rule=", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    svg.IndexOf("clip-rule=", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    svg.IndexOf("<defs", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    svg.IndexOf("<clipPath", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    svg.IndexOf("<mask", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    svg.IndexOf("<linearGradient", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    svg.IndexOf("<radialGradient", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    svg.IndexOf("<pattern", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    svg.IndexOf("<image", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    svg.IndexOf("<ellipse", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                var compoundFilledPathMatches = Regex.Matches(
                    svg,
                    "<path[^>]*\\sd\\s*=\\s*\"([^\"]+)\"[^>]*>",
                    RegexOptions.IgnoreCase);

                foreach (Match match in compoundFilledPathMatches)
                {
                    string pathData = match.Groups[1].Value;
                    int moveCount = Regex.Matches(pathData, "[Mm]").Count;
                    if (moveCount > 1)
                    {
                        bool isExplicitOutline = Regex.IsMatch(match.Value, "\\bstroke\\s*=\\s*\"[^\"]+\"", RegexOptions.IgnoreCase) &&
                            !Regex.IsMatch(match.Value, "\\bfill\\s*=\\s*\"(?!none)[^\"]+\"", RegexOptions.IgnoreCase);

                        if (!isExplicitOutline)
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.ShouldFallbackToSvgInsert");
            }

            return false;
        }

        private static bool IsPhosphorRegular(IconCatalogItem iconItem)
        {
            if (iconItem == null)
            {
                return false;
            }

            return string.Equals(iconItem.CollectionId, "phosphor", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(iconItem.Style, "regular", StringComparison.OrdinalIgnoreCase);
        }

        private string CreateColorizedTempSvg(string sourcePath, string hexColor)
        {
            try
            {
                string svg = File.ReadAllText(sourcePath);
                string safeColor = NormalizeHexColor(hexColor);

                svg = Regex.Replace(svg, "currentColor", safeColor, RegexOptions.IgnoreCase);

                if (!svg.Contains(safeColor))
                {
                    svg = svg.Replace("<svg ", $"<svg color=\"{safeColor}\" ");
                }

                string tempDir = Path.Combine(Path.GetTempPath(), "oPenEfficiency", "IconBrowser");
                Directory.CreateDirectory(tempDir);

                string tempFile = Path.Combine(tempDir, $"{Guid.NewGuid():N}.svg");
                File.WriteAllText(tempFile, svg);
                return tempFile;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.CreateColorizedTempSvg");
                return null;
            }
        }

        private string ResolveArchivePath()
        {
            string runtimePath = Path.Combine(_assemblyDir, ArchiveRelativePath);
            if (File.Exists(runtimePath))
            {
                return runtimePath;
            }

            string devPath = Path.GetFullPath(Path.Combine(_assemblyDir, "..", "..", ArchiveRelativePath));
            if (File.Exists(devPath))
            {
                return devPath;
            }

            return null;
        }

        private static string NormalizeHexColor(string hexColor)
        {
            if (string.IsNullOrWhiteSpace(hexColor))
            {
                return "#111111";
            }

            string value = hexColor.Trim();
            if (!value.StartsWith("#", StringComparison.Ordinal))
            {
                value = "#" + value;
            }

            return value.Length == 7 ? value : "#111111";
        }

        private static int HexToOfficeRgb(string hexColor)
        {
            string normalized = NormalizeHexColor(hexColor);
            int red = Convert.ToInt32(normalized.Substring(1, 2), 16);
            int green = Convert.ToInt32(normalized.Substring(3, 2), 16);
            int blue = Convert.ToInt32(normalized.Substring(5, 2), 16);
            return red | (green << 8) | (blue << 16);
        }

        private static float TransformCoordinate(double value, double origin, float scale, float offset)
        {
            return offset + (float)((value - origin) * scale);
        }

        private static bool ArePointsClose(System.Windows.Point first, System.Windows.Point second)
        {
            return Math.Abs(first.X - second.X) < 0.01 && Math.Abs(first.Y - second.Y) < 0.01;
        }

        private static void TrySelectInsertedShapesById(PowerPoint.Slide slide, IReadOnlyList<int> shapeIds)
        {
            if (slide == null || shapeIds == null || shapeIds.Count == 0)
            {
                return;
            }

            try
            {
                var shapeNames = new List<string>();
                foreach (int shapeId in shapeIds)
                {
                    var shape = FindShapeById(slide, shapeId);
                    if (shape != null)
                    {
                        shapeNames.Add(shape.Name);
                    }
                }

                if (shapeNames.Count == 0)
                {
                    return;
                }

                if (shapeNames.Count == 1)
                {
                    slide.Shapes.Range(new[] { shapeNames[0] }).Select();
                    return;
                }

                slide.Shapes.Range(shapeNames.ToArray()).Select();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "IconCatalogService.TrySelectInsertedShapesById");
            }
        }
    }
}
