using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Threading;
using Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Loads bundled, presentation-optimized map geometry and renders it as native
    /// PowerPoint freeforms. The source GIS data is simplified at build time so no
    /// network or GIS runtime is required on the user's machine.
    /// </summary>
    public static class MapAssetService
    {
        public const string MapIdTag = "OPE_MAP_ID";
        public const string RegionIdTag = "OPE_MAP_REGION_ID";
        public const string RegionAliasesTag = "OPE_MAP_REGION_ALIASES";
        public const string SourceTag = "OPE_MAP_SOURCE";
        public const string ScopeTag = "OPE_MAP_SCOPE";
        public const string DetailLevelTag = "OPE_MAP_DETAIL_LEVEL";
        public const string CountryIdTag = "OPE_MAP_COUNTRY_ID";
        public const string ProjectionTag = "OPE_MAP_PROJECTION";
        public const string CompactInsetsTag = "OPE_MAP_COMPACT_INSETS";
        public const string RemoteRegionModeTag = "OPE_MAP_REMOTE_REGIONS";

        private const string GermanyResourceSuffix = "Resources.Maps.de_adm1_naturalearth_5_1_1.json";
        private const string CountriesResourceSuffix = "Resources.Maps.countries_admin0_naturalearth_5_1_1.json";
        private const string CountriesOverviewResourceSuffix = "Resources.Maps.countries_overview_admin0_naturalearth_5_1_1.json";
        private const string AdministrativeRegionsResourceSuffix = "Resources.Maps.states_provinces_admin1_naturalearth_5_1_1.json.gz";
        private const string AdministrativeRegionsCatalogResourceSuffix = "Resources.Maps.states_provinces_admin1_catalog_naturalearth_5_1_1.tsv";
        private const string GermanyAdministrativeDistrictsResourceSuffix = "Resources.Maps.de_adm2_government_districts_geoboundaries_2021.json.gz";
        private const string GermanyDistrictsResourceSuffix = "Resources.Maps.de_adm3_districts_geoboundaries_2021.json.gz";
        private const string GlobalAdministrativeRegionsResourceSuffix = "Resources.Maps.administrative_regions_adm2_geoboundaries_current.json.gz";
        private const string GlobalThirdLevelAdministrativeRegionsResourceSuffix = "Resources.Maps.administrative_regions_adm3_geoboundaries_current.json.gz";
        private const string PrebuiltOverviewResourceSuffix = "Resources.Maps.map_overviews_naturalearth_5_1_1.pptx";
        private const string MercatorPrebuiltOverviewResourceSuffix = "Resources.Maps.map_overviews_mercator_naturalearth_5_1_1.pptx";
        private const float DefaultMapHeightRatio = 0.72f;
        private static readonly object AssetCacheLock = new object();
        private static readonly object PrebuiltAssetLock = new object();
        private static readonly Dictionary<string, MapAssetDefinition> AssetCache = new Dictionary<string, MapAssetDefinition>(StringComparer.OrdinalIgnoreCase);
        private const int GlobalAdministrativeRegionsIdleCacheMilliseconds = 5 * 60 * 1000;
        private static Timer GlobalAdministrativeRegionsCacheTimer;
        private static Timer GlobalThirdLevelAdministrativeRegionsCacheTimer;
        private static Dictionary<string, int> AdministrativeRegionCounts;
        private static AdministrativeMetadata AdministrativeMetadataCache;
        private const string AdministrativeMetadataResourceSuffix = "Resources.Maps.administrative_region_metadata.json";

        [DataContract]
        private sealed class AdministrativeMetadata
        {
            [DataMember] public List<MapAssetRegion> Regions { get; set; }
            [DataMember] public List<AdministrativeCount> Counts { get; set; }
        }

        [DataContract]
        private sealed class AdministrativeCount
        {
            [DataMember] public string Level { get; set; }
            [DataMember] public string Country { get; set; }
            [DataMember] public string Parent { get; set; }
            [DataMember] public int Count { get; set; }
        }

        // Selection controls must not deserialize global polygon datasets on the UI thread.
        // The generated metadata is derived from those exact bundled datasets; insertion
        // still loads the original geometry through LoadBundledAsset.
        private static AdministrativeMetadata GetAdministrativeMetadata()
        {
            lock (AssetCacheLock)
            {
                if (AdministrativeMetadataCache != null) return AdministrativeMetadataCache;
                Assembly assembly = Assembly.GetExecutingAssembly();
                string resource = assembly.GetManifestResourceNames().Single(name =>
                    name.EndsWith(AdministrativeMetadataResourceSuffix, StringComparison.OrdinalIgnoreCase));
                using (Stream stream = assembly.GetManifestResourceStream(resource))
                {
                    var serializer = new DataContractJsonSerializer(typeof(AdministrativeMetadata));
                    var metadata = (AdministrativeMetadata)serializer.ReadObject(stream);
                    if (metadata?.Regions == null || metadata.Counts == null)
                        throw new InvalidDataException("Bundled administrative map metadata is incomplete.");
                    AdministrativeMetadataCache = metadata;
                    return metadata;
                }
            }
        }

        private static int GetAdministrativeMetadataCount(string level, string countryId, string parentRegionId)
        {
            if (string.IsNullOrWhiteSpace(countryId)) return 0;
            return GetAdministrativeMetadata().Counts.Where(item =>
                string.Equals(item.Level, level, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(item.Country, countryId, StringComparison.OrdinalIgnoreCase) &&
                (string.IsNullOrWhiteSpace(parentRegionId) ||
                 string.Equals(item.Parent, parentRegionId, StringComparison.OrdinalIgnoreCase)))
                .Sum(item => item.Count);
        }

        public static IReadOnlyList<MapCatalogEntry> GetCountryCatalog()
        {
            MapAssetDefinition asset = LoadBundledAsset(CountriesResourceSuffix);
            if (asset?.Regions == null)
            {
                return new List<MapCatalogEntry>();
            }

            return asset.Regions
                .Select(region => new MapCatalogEntry
                {
                    Id = region.Id,
                    DisplayName = BuildCatalogDisplayName(region),
                    LocalName = region.Name,
                    Group = region.Group,
                    EntityType = region.EntityType
                })
                .OrderBy(entry => entry.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        public static IReadOnlyList<string> GetContinentCatalog()
        {
            return new[] { "Africa", "Asia", "Europe", "North America", "South America", "Oceania", "Antarctica" };
        }

        public static PowerPoint.Shape InsertGermanyFederalStates(PowerPointManager manager, MapProjection projection = MapProjection.Mercator)
        {
            return InsertCountryAdministrativeRegions(manager, "DEU", projection) ??
                   InsertBundledMap(manager, GermanyResourceSuffix, region => true, false, "DE-ADM1", "Germany – Federal States", "Country", "ADM1", "DEU", projection);
        }

        public static int GetAdministrativeRegionCount(string countryId)
        {
            if (string.IsNullOrWhiteSpace(countryId))
            {
                return 0;
            }

            EnsureAdministrativeRegionCatalog();
            lock (AssetCacheLock)
            {
                int count;
                return AdministrativeRegionCounts != null && AdministrativeRegionCounts.TryGetValue(countryId, out count)
                    ? count
                    : 0;
            }
        }

        public static IReadOnlyList<MapCatalogEntry> GetGermanyStateCatalog()
        {
            return GetCountryAdministrativeRegionCatalog("DEU");
        }

        public static IReadOnlyList<MapCatalogEntry> GetCountryAdministrativeRegionCatalog(string countryId)
        {
            return GetAdministrativeMetadata().Regions
                .Where(region => string.Equals(region.Group, countryId, StringComparison.OrdinalIgnoreCase))
                .Select(region => new MapCatalogEntry
                {
                    Id = region.Id,
                    DisplayName = region.EnglishName ?? region.Name ?? region.Id,
                    LocalName = region.Name,
                    Group = region.Group,
                    EntityType = region.EntityType
                })
                .OrderBy(entry => entry.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                .ToList() ?? new List<MapCatalogEntry>();
        }

        public static IReadOnlyList<MapCatalogEntry> GetLowerLevelParentRegionCatalog(string countryId, string detailLevel)
        {
            string level = string.Equals(detailLevel, "ADM3", StringComparison.OrdinalIgnoreCase) ? "ADM3" : "ADM2";
            var parentIds = new HashSet<string>(GetAdministrativeMetadata().Counts
                .Where(item => string.Equals(item.Level, level, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(item.Country, countryId, StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(item.Parent))
                .Select(item => item.Parent), StringComparer.OrdinalIgnoreCase);
            return GetCountryAdministrativeRegionCatalog(countryId)
                .Where(region => parentIds.Contains(region.Id))
                .ToList();
        }

        public static int GetGermanyAdministrativeDistrictCount(string stateId = null)
        {
            MapAssetDefinition asset = LoadBundledAsset(GermanyAdministrativeDistrictsResourceSuffix);
            return asset?.Regions?.Count(region =>
                string.IsNullOrWhiteSpace(stateId) ||
                string.Equals(region.ParentGroup, stateId, StringComparison.OrdinalIgnoreCase)) ?? 0;
        }

        public static int GetSecondLevelAdministrativeRegionCount(string countryId, string parentRegionId = null)
        {
            if (string.IsNullOrWhiteSpace(countryId))
            {
                return 0;
            }
            return GetAdministrativeMetadataCount("ADM2", countryId, parentRegionId);
        }

        public static PowerPoint.Shape InsertCountrySecondLevelAdministrativeRegions(
            PowerPointManager manager,
            string countryId,
            string parentRegionId = null,
            MapProjection projection = MapProjection.Mercator)
        {
            if (GetSecondLevelAdministrativeRegionCount(countryId, parentRegionId) == 0)
            {
                return null;
            }
            string countryName = GetCountryCatalog()
                .FirstOrDefault(country => string.Equals(country.Id, countryId, StringComparison.OrdinalIgnoreCase))?.DisplayName ?? countryId;
            string parentName = GetCountryAdministrativeRegionCatalog(countryId)
                .FirstOrDefault(region => string.Equals(region.Id, parentRegionId, StringComparison.OrdinalIgnoreCase))?.DisplayName;
            return InsertBundledMap(
                manager,
                GlobalAdministrativeRegionsResourceSuffix,
                region => string.Equals(region.Group, countryId, StringComparison.OrdinalIgnoreCase) &&
                    (string.IsNullOrWhiteSpace(parentRegionId) || string.Equals(region.ParentGroup, parentRegionId, StringComparison.OrdinalIgnoreCase)),
                false,
                string.IsNullOrWhiteSpace(parentRegionId) ? $"COUNTRY-{countryId.ToUpperInvariant()}-ADM2" : $"COUNTRY-{countryId.ToUpperInvariant()}-{parentRegionId}-ADM2",
                string.IsNullOrWhiteSpace(parentName) ? $"{countryName} – Second-Level Regions" : $"{countryName} – {parentName} Second-Level Regions",
                "Country",
                "ADM2",
                countryId.ToUpperInvariant(),
                projection);
        }

        public static int GetThirdLevelAdministrativeRegionCount(string countryId, string parentRegionId = null)
        {
            if (string.IsNullOrWhiteSpace(countryId))
            {
                return 0;
            }
            return GetAdministrativeMetadataCount("ADM3", countryId, parentRegionId);
        }

        public static PowerPoint.Shape InsertCountryThirdLevelAdministrativeRegions(
            PowerPointManager manager,
            string countryId,
            string parentRegionId = null,
            MapProjection projection = MapProjection.Mercator)
        {
            if (GetThirdLevelAdministrativeRegionCount(countryId, parentRegionId) == 0)
            {
                return null;
            }
            string countryName = GetCountryCatalog()
                .FirstOrDefault(country => string.Equals(country.Id, countryId, StringComparison.OrdinalIgnoreCase))?.DisplayName ?? countryId;
            string parentName = GetCountryAdministrativeRegionCatalog(countryId)
                .FirstOrDefault(region => string.Equals(region.Id, parentRegionId, StringComparison.OrdinalIgnoreCase))?.DisplayName;
            return InsertBundledMap(
                manager,
                GlobalThirdLevelAdministrativeRegionsResourceSuffix,
                region => string.Equals(region.Group, countryId, StringComparison.OrdinalIgnoreCase) &&
                    (string.IsNullOrWhiteSpace(parentRegionId) || string.Equals(region.ParentGroup, parentRegionId, StringComparison.OrdinalIgnoreCase)),
                false,
                string.IsNullOrWhiteSpace(parentRegionId) ? $"COUNTRY-{countryId.ToUpperInvariant()}-ADM3" : $"COUNTRY-{countryId.ToUpperInvariant()}-{parentRegionId}-ADM3",
                string.IsNullOrWhiteSpace(parentName) ? $"{countryName} – Third-Level Regions" : $"{countryName} – {parentName} Third-Level Regions",
                "Country",
                "ADM3",
                countryId.ToUpperInvariant(),
                projection);
        }

        public static PowerPoint.Shape InsertGermanyAdministrativeDistricts(
            PowerPointManager manager,
            string stateId,
            MapProjection projection = MapProjection.Mercator)
        {
            bool isFiltered = !string.IsNullOrWhiteSpace(stateId);
            string stateName = GetGermanyStateCatalog()
                .FirstOrDefault(state => string.Equals(state.Id, stateId, StringComparison.OrdinalIgnoreCase))?.DisplayName;
            return InsertBundledMap(
                manager,
                GermanyAdministrativeDistrictsResourceSuffix,
                region => !isFiltered || string.Equals(region.ParentGroup, stateId, StringComparison.OrdinalIgnoreCase),
                false,
                isFiltered ? $"DEU-{stateId}-ADM2" : "DEU-ADM2",
                isFiltered ? $"Germany – {stateName} Government Districts" : "Germany – Government Districts",
                "Country",
                "ADM2",
                "DEU",
                projection);
        }

        public static int GetGermanyDistrictCount(string stateId)
        {
            if (string.IsNullOrWhiteSpace(stateId))
            {
                return 0;
            }
            MapAssetDefinition asset = LoadBundledAsset(GermanyDistrictsResourceSuffix);
            return asset?.Regions?.Count(region => string.Equals(region.ParentGroup, stateId, StringComparison.OrdinalIgnoreCase)) ?? 0;
        }

        public static PowerPoint.Shape InsertGermanyDistricts(PowerPointManager manager, string stateId, MapProjection projection = MapProjection.Mercator)
        {
            if (GetGermanyDistrictCount(stateId) == 0)
            {
                return null;
            }
            string stateName = GetGermanyStateCatalog()
                .FirstOrDefault(state => string.Equals(state.Id, stateId, StringComparison.OrdinalIgnoreCase))?.DisplayName;
            return InsertBundledMap(
                manager,
                GermanyDistrictsResourceSuffix,
                region => string.Equals(region.ParentGroup, stateId, StringComparison.OrdinalIgnoreCase),
                false,
                $"DEU-{stateId}-ADM3",
                $"Germany – {stateName} Districts and Independent Cities",
                "Country",
                "ADM3",
                "DEU",
                projection);
        }

        public static PowerPoint.Shape InsertCountryAdministrativeRegions(
            PowerPointManager manager,
            string countryId,
            MapProjection projection = MapProjection.Mercator,
            MapRemoteRegionMode remoteRegionMode = MapRemoteRegionMode.Geographic)
        {
            if (string.IsNullOrWhiteSpace(countryId) || GetAdministrativeRegionCount(countryId) == 0)
            {
                return null;
            }

            MapCatalogEntry country = GetCountryCatalog()
                .FirstOrDefault(entry => string.Equals(entry.Id, countryId, StringComparison.OrdinalIgnoreCase));
            string countryName = country?.DisplayName ?? countryId;
            return InsertBundledMap(
                manager,
                AdministrativeRegionsResourceSuffix,
                region => string.Equals(region.Group, countryId, StringComparison.OrdinalIgnoreCase),
                false,
                $"COUNTRY-{countryId.ToUpperInvariant()}-ADM1",
                $"{countryName} – States and Provinces",
                "Country",
                "ADM1",
                countryId.ToUpperInvariant(),
                projection,
                remoteRegionMode);
        }

        public static PowerPoint.Shape InsertCountryAdministrativeRegions(
            PowerPointManager manager,
            string countryId,
            MapProjection projection,
            bool compactInsets)
        {
            return InsertCountryAdministrativeRegions(
                manager,
                countryId,
                projection,
                compactInsets ? MapRemoteRegionMode.CompactInsets : MapRemoteRegionMode.Geographic);
        }

        public static PowerPoint.Shape InsertEuropeCountries(PowerPointManager manager, MapProjection projection = MapProjection.Mercator)
        {
            return InsertContinentCountries(manager, "Europe", projection);
        }

        public static PowerPoint.Shape InsertWorldCountries(PowerPointManager manager, MapProjection projection = MapProjection.Mercator)
        {
            return InsertPrebuiltOverview(manager, "WORLD-ADM0", projection) ?? InsertWorldCountriesFromGeometry(manager, projection);
        }

        public static PowerPoint.Shape InsertWorldCountriesFromGeometry(PowerPointManager manager, MapProjection projection = MapProjection.Mercator)
        {
            return InsertBundledMap(
                manager,
                CountriesOverviewResourceSuffix,
                region => !string.Equals(region.Group, "Antarctica", StringComparison.OrdinalIgnoreCase),
                false,
                "WORLD-ADM0",
                "World – Countries and Territories",
                "World",
                "ADM0",
                null,
                projection);
        }

        public static PowerPoint.Shape InsertContinentCountries(PowerPointManager manager, string continent, MapProjection projection = MapProjection.Mercator)
        {
            if (string.IsNullOrWhiteSpace(continent))
            {
                return null;
            }

            string mapId = $"CONTINENT-{continent.ToUpperInvariant().Replace(' ', '-')}";
            if (string.Equals(continent, "Antarctica", StringComparison.OrdinalIgnoreCase))
            {
                projection = MapProjection.Natural;
            }
            return InsertPrebuiltOverview(manager, mapId, projection) ?? InsertContinentCountriesFromGeometry(manager, continent, projection);
        }

        public static PowerPoint.Shape InsertContinentCountriesFromGeometry(PowerPointManager manager, string continent, MapProjection projection = MapProjection.Mercator)
        {
            if (string.IsNullOrWhiteSpace(continent))
            {
                return null;
            }

            return InsertBundledMap(
                manager,
                CountriesOverviewResourceSuffix,
                region => string.Equals(region.Group, continent, StringComparison.OrdinalIgnoreCase),
                false,
                $"CONTINENT-{continent.ToUpperInvariant().Replace(' ', '-')}",
                $"{continent} – Countries and Territories",
                "Continent",
                "ADM0",
                null,
                projection);
        }

        public static PowerPoint.Shape InsertCountry(
            PowerPointManager manager,
            string countryId,
            bool largestLandmassOnly = false,
            MapProjection projection = MapProjection.Mercator)
        {
            if (string.IsNullOrWhiteSpace(countryId))
            {
                return null;
            }

            MapAssetDefinition asset = LoadBundledAsset(CountriesResourceSuffix);
            MapAssetRegion country = asset?.Regions?.FirstOrDefault(region => string.Equals(region.Id, countryId, StringComparison.OrdinalIgnoreCase));
            if (country == null)
            {
                return null;
            }

            return InsertBundledMap(
                manager,
                CountriesResourceSuffix,
                region => string.Equals(region.Id, countryId, StringComparison.OrdinalIgnoreCase),
                largestLandmassOnly,
                $"COUNTRY-{country.Id}",
                country.EnglishName ?? country.Name ?? country.Id,
                "Country",
                "ADM0",
                country.Id,
                projection);
        }

        private static PowerPoint.Shape InsertPrebuiltOverview(PowerPointManager manager, string mapId, MapProjection projection)
        {
            if (manager == null || string.IsNullOrWhiteSpace(mapId))
            {
                return null;
            }

            PowerPoint.Presentation sourcePresentation = null;
            try
            {
                var app = manager.GetApplication();
                var targetSlide = app?.ActiveWindow?.View?.Slide as PowerPoint.Slide;
                if (targetSlide == null || app.ActivePresentation == null)
                {
                    return null;
                }

                string assetPath = EnsurePrebuiltOverviewAsset(projection);
                if (string.IsNullOrWhiteSpace(assetPath) || !File.Exists(assetPath))
                {
                    return null;
                }

                sourcePresentation = app.Presentations.Open(
                    assetPath,
                    MsoTriState.msoTrue,
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse);

                PowerPoint.Shape sourceShape = null;
                for (int slideIndex = 1; slideIndex <= sourcePresentation.Slides.Count && sourceShape == null; slideIndex++)
                {
                    PowerPoint.Slide sourceSlide = sourcePresentation.Slides[slideIndex];
                    for (int shapeIndex = 1; shapeIndex <= sourceSlide.Shapes.Count; shapeIndex++)
                    {
                        PowerPoint.Shape candidate = sourceSlide.Shapes[shapeIndex];
                        if (string.Equals(ShapeMetadataService.GetTag(candidate, MapIdTag), mapId, StringComparison.OrdinalIgnoreCase))
                        {
                            sourceShape = candidate;
                            break;
                        }
                    }
                }

                if (sourceShape == null)
                {
                    return null;
                }

                PowerPoint.ShapeRange pasted = CopyAndPasteShape(sourceShape, targetSlide);
                PowerPoint.Shape map = pasted.Count > 0 ? pasted[1] : null;
                if (map == null)
                {
                    return null;
                }

                FitMapToSlide(map, app.ActivePresentation.PageSetup.SlideWidth, app.ActivePresentation.PageSetup.SlideHeight);
                ShapeMetadataService.SetTag(map, ProjectionTag, projection == MapProjection.Mercator ? "Mercator" : "Equirectangular");
                sourcePresentation.Close();
                sourcePresentation = null;
                targetSlide.Select();
                map.Select();
                return map;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MapAssetService.InsertPrebuiltOverview");
                return null;
            }
            finally
            {
                if (sourcePresentation != null)
                {
                    try { sourcePresentation.Close(); } catch { /* Temporary source presentation may already be closed after an import failure. */ }
                }
            }
        }

        private static PowerPoint.ShapeRange CopyAndPasteShape(PowerPoint.Shape sourceShape, PowerPoint.Slide targetSlide)
        {
            Exception lastError = null;
            for (int attempt = 1; attempt <= 5; attempt++)
            {
                try
                {
                    sourceShape.Copy();
                    Thread.Sleep(60 * attempt);
                    return targetSlide.Shapes.Paste();
                }
                catch (Exception ex)
                {
                    lastError = ex;
                    Thread.Sleep(80 * attempt);
                }
            }

            throw new InvalidOperationException("PowerPoint did not make the prebuilt map available on the clipboard.", lastError);
        }

        private static string EnsurePrebuiltOverviewAsset(MapProjection projection)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string resourceSuffix = projection == MapProjection.Mercator
                ? MercatorPrebuiltOverviewResourceSuffix
                : PrebuiltOverviewResourceSuffix;
            string resourceName = assembly.GetManifestResourceNames()
                .FirstOrDefault(name => name.EndsWith(resourceSuffix, StringComparison.OrdinalIgnoreCase));
            if (resourceName == null)
            {
                return null;
            }

            string version = assembly.GetName().Version?.ToString() ?? "current";
            string cacheDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "oPenEfficiency",
                "Maps");
            string projectionKey = projection == MapProjection.Mercator ? "mercator" : "natural";
            string cachePath = Path.Combine(cacheDirectory, $"map_overviews_{projectionKey}_5_1_1_{version}.pptx");

            lock (PrebuiltAssetLock)
            {
                using (Stream resource = assembly.GetManifestResourceStream(resourceName))
                {
                    if (resource == null)
                    {
                        return null;
                    }

                    if (File.Exists(cachePath) && new FileInfo(cachePath).Length == resource.Length)
                    {
                        return cachePath;
                    }

                    Directory.CreateDirectory(cacheDirectory);
                    using (var output = new FileStream(cachePath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        resource.CopyTo(output);
                    }
                }
            }
            return cachePath;
        }

        private static void FitMapToSlide(PowerPoint.Shape map, float slideWidth, float slideHeight)
        {
            float maxWidth = slideWidth * 0.82f;
            float maxHeight = slideHeight * DefaultMapHeightRatio;
            float scale = Math.Min(maxWidth / Math.Max(1f, map.Width), maxHeight / Math.Max(1f, map.Height));
            map.LockAspectRatio = MsoTriState.msoTrue;
            map.Width *= scale;
            map.Left = (slideWidth - map.Width) / 2f;
            map.Top = (slideHeight - map.Height) / 2f;
        }

        private static PowerPoint.Shape InsertBundledMap(
            PowerPointManager manager,
            string resourceSuffix,
            Func<MapAssetRegion, bool> regionFilter,
            bool largestLandmassOnly,
            string mapId,
            string displayName,
            string scope,
            string detailLevel,
            string countryId,
            MapProjection projection,
            MapRemoteRegionMode remoteRegionMode = MapRemoteRegionMode.Geographic)
        {
            if (manager == null)
            {
                return null;
            }

            PowerPoint.Shape mapGroup = null;
            var regionRoots = new List<PowerPoint.Shape>();
            var currentParts = new List<PowerPoint.Shape>();

            try
            {
                var app = manager.GetApplication();
                var slide = app?.ActiveWindow?.View?.Slide as PowerPoint.Slide;
                if (slide == null || app.ActivePresentation == null)
                {
                    return null;
                }

                MapAssetDefinition asset = LoadBundledAsset(resourceSuffix);
                if (asset == null || asset.Regions == null || asset.Regions.Count == 0)
                {
                    throw new InvalidOperationException("The bundled map asset could not be loaded.");
                }

                var renderRegions = asset.Regions
                    .Where(regionFilter ?? (region => true))
                    .Select(region => new RenderRegion(region, SelectParts(region, largestLandmassOnly)))
                    .Where(region => region.Parts.Count > 0)
                    .ToList();
                if (string.Equals(detailLevel, "ADM1", StringComparison.OrdinalIgnoreCase))
                {
                    if (remoteRegionMode == MapRemoteRegionMode.CompactInsets)
                    {
                        renderRegions = ApplyCompactInsets(renderRegions, countryId);
                    }
                    else if (remoteRegionMode == MapRemoteRegionMode.Exclude)
                    {
                        renderRegions = ExcludeRemoteRegions(renderRegions, countryId);
                    }
                }
                if (renderRegions.Count == 0)
                {
                    throw new InvalidOperationException("The selected map did not contain any renderable regions.");
                }

                List<MapAssetPoint> renderPoints = renderRegions
                    .SelectMany(region => region.Parts)
                    .SelectMany(part => part.Points ?? Enumerable.Empty<MapAssetPoint>())
                    .ToList();
                double longitudePeriod = Math.Max(0.00001, asset.AspectRatio);
                double longitudeWrapStart = GetLongitudeWrapStart(scope, renderPoints, longitudePeriod);
                double rawSourceLeft = renderPoints.Min(point => UnwrapHorizontalCoordinate(point.X, longitudeWrapStart, longitudePeriod));
                double rawSourceRight = renderPoints.Max(point => UnwrapHorizontalCoordinate(point.X, longitudeWrapStart, longitudePeriod));
                VerticalProjection verticalProjection = CreateVerticalProjection(asset, resourceSuffix, projection);
                double sourceTop = renderPoints.Min(point => verticalProjection.Transform(point.Y));
                double horizontalOrigin = (rawSourceLeft + rawSourceRight) / 2d;
                double horizontalScale = GetLocalHorizontalScale(scope, renderPoints, projection);
                double sourceLeft = TransformHorizontalCoordinate(rawSourceLeft, horizontalOrigin, horizontalScale);
                double sourceWidth = Math.Max(
                    0.00001,
                    TransformHorizontalCoordinate(rawSourceRight, horizontalOrigin, horizontalScale) - sourceLeft);
                double sourceHeight = Math.Max(0.00001, renderPoints.Max(point => verticalProjection.Transform(point.Y)) - sourceTop);
                double renderedAspectRatio = sourceWidth / sourceHeight;

                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
                float mapHeight = slideHeight * DefaultMapHeightRatio;
                float mapWidth = mapHeight * (float)renderedAspectRatio;
                if (mapWidth > slideWidth * 0.82f)
                {
                    mapWidth = slideWidth * 0.82f;
                    mapHeight = mapWidth / (float)renderedAspectRatio;
                }

                float left = (slideWidth - mapWidth) / 2f;
                float top = (slideHeight - mapHeight) / 2f;
                float scale = mapHeight / (float)sourceHeight;

                foreach (RenderRegion renderRegion in renderRegions)
                {
                    MapAssetRegion region = renderRegion.Region;
                    currentParts.Clear();
                    string regionAliases = BuildAliases(region);
                    int partIndex = 0;
                    foreach (MapAssetPolygon polygon in renderRegion.Parts)
                    {
                        PowerPoint.Shape part = CreatePolygon(
                            slide,
                            polygon,
                            left,
                            top,
                            sourceLeft,
                            sourceTop,
                            scale,
                            horizontalOrigin,
                            horizontalScale,
                            longitudeWrapStart,
                            longitudePeriod,
                            verticalProjection);
                        if (part == null)
                        {
                            continue;
                        }

                        ApplyDefaultFormatting(part);
                        partIndex++;
                        part.Name = $"oPE_MapPart_{SanitizeName(region.Id)}_{partIndex}_{part.Id}";
                        part.AlternativeText = region.Name;
                        ShapeMetadataService.SetTag(part, RegionIdTag, region.Id);
                        ShapeMetadataService.SetTag(part, RegionAliasesTag, regionAliases);
                        currentParts.Add(part);
                    }

                    if (currentParts.Count == 0)
                    {
                        continue;
                    }

                    // PowerPoint flattens nested groups when the complete map is grouped.
                    // Avoid an expensive intermediate COM Group call for multi-region maps;
                    // every part already carries the logical region ID and aliases.
                    if (renderRegions.Count > 1)
                    {
                        regionRoots.AddRange(currentParts);
                        currentParts.Clear();
                        continue;
                    }

                    PowerPoint.Shape regionRoot;
                    if (currentParts.Count == 1)
                    {
                        regionRoot = currentParts[0];
                    }
                    else
                    {
                        regionRoot = slide.Shapes.Range(currentParts.Select(shape => shape.Name).ToArray()).Group();
                    }

                    regionRoot.Name = $"oPE_MapRegion_{SanitizeName(region.Id)}_{regionRoot.Id}";
                    regionRoot.AlternativeText = region.Name;
                    ShapeMetadataService.SetTag(regionRoot, RegionIdTag, region.Id);
                    ShapeMetadataService.SetTag(regionRoot, RegionAliasesTag, regionAliases);
                    regionRoots.Add(regionRoot);
                    currentParts.Clear();
                }

                if (regionRoots.Count == 0)
                {
                    throw new InvalidOperationException("The selected map did not contain any renderable regions.");
                }

                mapGroup = regionRoots.Count == 1
                    ? regionRoots[0]
                    : slide.Shapes.Range(regionRoots.Select(shape => shape.Name).ToArray()).Group();
                mapGroup.Name = $"oPE_Map_{SanitizeName(mapId)}_{mapGroup.Id}";
                mapGroup.AlternativeText = displayName;
                ShapeMetadataService.SetTag(mapGroup, MapIdTag, mapId);
                ShapeMetadataService.SetTag(mapGroup, SourceTag, $"{asset.Source} {asset.SourceVersion}|{asset.SourceUrl}");
                ShapeMetadataService.SetTag(mapGroup, ScopeTag, scope);
                ShapeMetadataService.SetTag(mapGroup, DetailLevelTag, detailLevel);
                ShapeMetadataService.SetTag(
                    mapGroup,
                    ProjectionTag,
                    projection == MapProjection.Mercator
                        ? "Mercator"
                        : string.Equals(scope, "Country", StringComparison.OrdinalIgnoreCase)
                            ? "Local Equirectangular"
                            : "Equirectangular");
                ShapeMetadataService.SetTag(mapGroup, CompactInsetsTag, remoteRegionMode == MapRemoteRegionMode.CompactInsets ? "True" : "False");
                ShapeMetadataService.SetTag(mapGroup, RemoteRegionModeTag, remoteRegionMode.ToString());
                if (!string.IsNullOrWhiteSpace(countryId))
                {
                    ShapeMetadataService.SetTag(mapGroup, CountryIdTag, countryId);
                }
                mapGroup.Select();
                return mapGroup;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MapAssetService.InsertBundledMap");
                TryDelete(mapGroup);
                foreach (PowerPoint.Shape shape in regionRoots)
                {
                    TryDelete(shape);
                }
                foreach (PowerPoint.Shape shape in currentParts)
                {
                    TryDelete(shape);
                }
                return null;
            }
        }

        private static MapAssetDefinition LoadBundledAsset(string resourceSuffix)
        {
            lock (AssetCacheLock)
            {
                MapAssetDefinition cached;
                if (AssetCache.TryGetValue(resourceSuffix, out cached))
                {
                    ResetGlobalAdministrativeRegionsIdleTimer(resourceSuffix);
                    return cached;
                }
            }

            Assembly assembly = Assembly.GetExecutingAssembly();
            string resourceName = assembly.GetManifestResourceNames()
                .FirstOrDefault(name => name.EndsWith(resourceSuffix, StringComparison.OrdinalIgnoreCase));
            if (resourceName == null)
            {
                return null;
            }

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    return null;
                }

                var serializer = new DataContractJsonSerializer(typeof(MapAssetDefinition));
                MapAssetDefinition asset;
                if (resourceSuffix.EndsWith(".gz", StringComparison.OrdinalIgnoreCase))
                {
                    using (var gzip = new GZipStream(stream, CompressionMode.Decompress, false))
                    {
                        asset = serializer.ReadObject(gzip) as MapAssetDefinition;
                    }
                }
                else
                {
                    asset = serializer.ReadObject(stream) as MapAssetDefinition;
                }
                if (asset != null)
                {
                    lock (AssetCacheLock)
                    {
                        AssetCache[resourceSuffix] = asset;
                    }
                    ResetGlobalAdministrativeRegionsIdleTimer(resourceSuffix);
                }
                return asset;
            }
        }

        private static void ResetGlobalAdministrativeRegionsIdleTimer(string resourceSuffix)
        {
            if (!string.Equals(resourceSuffix, GlobalAdministrativeRegionsResourceSuffix, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(resourceSuffix, GlobalThirdLevelAdministrativeRegionsResourceSuffix, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            lock (AssetCacheLock)
            {
                if (string.Equals(resourceSuffix, GlobalAdministrativeRegionsResourceSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    if (GlobalAdministrativeRegionsCacheTimer == null)
                    {
                        GlobalAdministrativeRegionsCacheTimer = new Timer(
                            _ => EvictGlobalAdministrativeRegionsCache(), null, Timeout.Infinite, Timeout.Infinite);
                    }
                    GlobalAdministrativeRegionsCacheTimer.Change(GlobalAdministrativeRegionsIdleCacheMilliseconds, Timeout.Infinite);
                }
                else
                {
                    if (GlobalThirdLevelAdministrativeRegionsCacheTimer == null)
                    {
                        GlobalThirdLevelAdministrativeRegionsCacheTimer = new Timer(
                            _ => EvictGlobalThirdLevelAdministrativeRegionsCache(), null, Timeout.Infinite, Timeout.Infinite);
                    }
                    GlobalThirdLevelAdministrativeRegionsCacheTimer.Change(GlobalAdministrativeRegionsIdleCacheMilliseconds, Timeout.Infinite);
                }
            }
        }

        private static void EvictGlobalAdministrativeRegionsCache()
        {
            lock (AssetCacheLock)
            {
                AssetCache.Remove(GlobalAdministrativeRegionsResourceSuffix);
            }
        }

        private static void EvictGlobalThirdLevelAdministrativeRegionsCache()
        {
            lock (AssetCacheLock)
            {
                AssetCache.Remove(GlobalThirdLevelAdministrativeRegionsResourceSuffix);
            }
        }

        private static void EnsureAdministrativeRegionCatalog()
        {
            lock (AssetCacheLock)
            {
                if (AdministrativeRegionCounts != null)
                {
                    return;
                }

                var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                Assembly assembly = Assembly.GetExecutingAssembly();
                string resourceName = assembly.GetManifestResourceNames()
                    .FirstOrDefault(name => name.EndsWith(AdministrativeRegionsCatalogResourceSuffix, StringComparison.OrdinalIgnoreCase));
                if (resourceName != null)
                {
                    using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                    using (var reader = stream == null ? null : new StreamReader(stream))
                    {
                        string line;
                        while (reader != null && (line = reader.ReadLine()) != null)
                        {
                            string[] parts = line.Split('\t');
                            int count;
                            if (parts.Length == 2 && int.TryParse(parts[1], out count) && count > 0)
                            {
                                counts[parts[0]] = count;
                            }
                        }
                    }
                }

                AdministrativeRegionCounts = counts;
            }
        }

        private static PowerPoint.Shape CreatePolygon(
            PowerPoint.Slide slide,
            MapAssetPolygon polygon,
            float left,
            float top,
            double sourceLeft,
            double sourceTop,
            float scale,
            double horizontalOrigin,
            double horizontalScale,
            double longitudeWrapStart,
            double longitudePeriod,
            VerticalProjection verticalProjection)
        {
            if (polygon?.Points == null || polygon.Points.Count < 3)
            {
                return null;
            }

            MapAssetPoint first = polygon.Points[0];
            float startX = left + ((float)(TransformHorizontalCoordinate(
                UnwrapHorizontalCoordinate(first.X, longitudeWrapStart, longitudePeriod),
                horizontalOrigin,
                horizontalScale) - sourceLeft) * scale);
            float startY = top + ((float)(verticalProjection.Transform(first.Y) - sourceTop) * scale);
            PowerPoint.FreeformBuilder builder = slide.Shapes.BuildFreeform(
                MsoEditingType.msoEditingCorner,
                startX,
                startY);

            for (int index = 1; index < polygon.Points.Count; index++)
            {
                MapAssetPoint point = polygon.Points[index];
                builder.AddNodes(
                    MsoSegmentType.msoSegmentLine,
                    MsoEditingType.msoEditingCorner,
                    left + ((float)(TransformHorizontalCoordinate(
                        UnwrapHorizontalCoordinate(point.X, longitudeWrapStart, longitudePeriod),
                        horizontalOrigin,
                        horizontalScale) - sourceLeft) * scale),
                    top + ((float)(verticalProjection.Transform(point.Y) - sourceTop) * scale));
            }

            builder.AddNodes(
                MsoSegmentType.msoSegmentLine,
                MsoEditingType.msoEditingCorner,
                startX,
                startY);
            return builder.ConvertToShape();
        }

        /// <summary>
        /// Corrects the horizontal scale of a standalone country at its mean
        /// latitude. Overview maps retain their shared projection so neighboring
        /// countries continue to align exactly.
        /// </summary>
        private static double GetLocalHorizontalScale(
            string scope,
            IReadOnlyList<MapAssetPoint> points,
            MapProjection projection)
        {
            if (projection == MapProjection.Mercator ||
                !string.Equals(scope, "Country", StringComparison.OrdinalIgnoreCase) ||
                points == null ||
                points.Count == 0)
            {
                return 1d;
            }

            double centerY = (points.Min(point => point.Y) + points.Max(point => point.Y)) / 2d;
            double centerLatitudeRadians = (0.5d - centerY) * Math.PI;
            return Math.Max(0.2d, Math.Cos(centerLatitudeRadians));
        }

        /// <summary>
        /// Finds the empty longitudinal interval where a standalone country map can
        /// be cut. This keeps countries crossing the antimeridian, such as New Zealand
        /// and the United States, together instead of stretching them across the world.
        /// </summary>
        private static double GetLongitudeWrapStart(string scope, IReadOnlyList<MapAssetPoint> points, double period)
        {
            if (!string.Equals(scope, "Country", StringComparison.OrdinalIgnoreCase) || points == null || points.Count < 2)
            {
                return double.NaN;
            }

            List<double> coordinates = points
                .Select(point => point.X)
                .Distinct()
                .OrderBy(value => value)
                .ToList();
            if (coordinates.Count < 2)
            {
                return coordinates[0];
            }

            double largestGap = (coordinates[0] + period) - coordinates[coordinates.Count - 1];
            double wrapStart = coordinates[0];
            for (int index = 0; index < coordinates.Count - 1; index++)
            {
                double gap = coordinates[index + 1] - coordinates[index];
                if (gap > largestGap)
                {
                    largestGap = gap;
                    wrapStart = coordinates[index + 1];
                }
            }

            return wrapStart;
        }

        private static double UnwrapHorizontalCoordinate(double x, double wrapStart, double period)
        {
            if (double.IsNaN(wrapStart))
            {
                return x;
            }

            return x < wrapStart ? x + period : x;
        }

        private static double TransformHorizontalCoordinate(double x, double origin, double horizontalScale)
        {
            return origin + ((x - origin) * horizontalScale);
        }

        private static VerticalProjection CreateVerticalProjection(
            MapAssetDefinition asset,
            string resourceSuffix,
            MapProjection projection)
        {
            if (projection != MapProjection.Mercator)
            {
                return VerticalProjection.Identity;
            }

            bool localCoordinates = string.Equals(asset.CoordinateSpace, "LocalEquirectangular", StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(resourceSuffix, GermanyResourceSuffix, StringComparison.OrdinalIgnoreCase);
            double minLatitude = asset.MinLatitude ?? (localCoordinates ? 47.2701d : -90d);
            double maxLatitude = asset.MaxLatitude ?? (localCoordinates ? 55.0581d : 90d);
            return new VerticalProjection(true, localCoordinates, minLatitude, maxLatitude);
        }

        private sealed class VerticalProjection
        {
            private const double MercatorLatitudeLimit = 85.05112878d;
            public static readonly VerticalProjection Identity = new VerticalProjection(false, false, -90d, 90d);
            private readonly bool _mercator;
            private readonly bool _normalizeLocally;
            private readonly double _minLatitude;
            private readonly double _maxLatitude;
            private readonly double _localTop;
            private readonly double _localHeight;

            public VerticalProjection(bool mercator, bool normalizeLocally, double minLatitude, double maxLatitude)
            {
                _mercator = mercator;
                _normalizeLocally = normalizeLocally;
                _minLatitude = minLatitude;
                _maxLatitude = maxLatitude;
                _localTop = MercatorCoordinate(maxLatitude);
                _localHeight = Math.Max(0.00001d, MercatorCoordinate(minLatitude) - _localTop);
            }

            public double Transform(double sourceY)
            {
                if (!_mercator)
                {
                    return sourceY;
                }

                double latitude = _normalizeLocally
                    ? _maxLatitude - (sourceY * (_maxLatitude - _minLatitude))
                    : 90d - (sourceY * 180d);
                double mercatorY = MercatorCoordinate(latitude);
                return _normalizeLocally ? (mercatorY - _localTop) / _localHeight : mercatorY;
            }

            private static double MercatorCoordinate(double latitude)
            {
                double clampedLatitude = Math.Max(-MercatorLatitudeLimit, Math.Min(MercatorLatitudeLimit, latitude));
                double radians = clampedLatitude * Math.PI / 180d;
                return 1d - (Math.Log(Math.Tan((Math.PI / 4d) + (radians / 2d))) / Math.PI);
            }
        }

        private static void ApplyDefaultFormatting(PowerPoint.Shape shape)
        {
            shape.Fill.Visible = MsoTriState.msoTrue;
            shape.Fill.Solid();
            shape.Fill.ForeColor.RGB = ColorConstants.FromHex("#D9E2F3");
            shape.Fill.Transparency = 0f;
            shape.Line.Visible = MsoTriState.msoTrue;
            shape.Line.ForeColor.RGB = ColorConstants.FromHex("#FFFFFF");
            shape.Line.Weight = 0.75f;
        }

        private static string BuildAliases(MapAssetRegion region)
        {
            var aliases = new List<string> { region.Id, region.Name, region.EnglishName };
            if (region.Aliases != null)
            {
                aliases.AddRange(region.Aliases);
            }

            return string.Join("|", aliases
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase));
        }

        private static string BuildCatalogDisplayName(MapAssetRegion region)
        {
            string englishName = region.EnglishName ?? region.Name ?? region.Id;
            string localSuffix = !string.IsNullOrWhiteSpace(region.Name) &&
                                 !string.Equals(region.Name, englishName, StringComparison.OrdinalIgnoreCase)
                ? $" / {region.Name}"
                : string.Empty;
            return $"{englishName}{localSuffix} ({region.Id})";
        }

        private static List<RenderRegion> ApplyCompactInsets(List<RenderRegion> regions, string countryId)
        {
            AdministrativeRegionSplit split = ClassifyAdministrativeRegions(regions, countryId);
            if (split == null || split.MainlandRegions.Count == 0 || split.RemoteRegions.Count == 0)
            {
                return regions;
            }

            MapAssetDefinition countries = LoadBundledAsset(CountriesResourceSuffix);
            MapBounds mainBounds = GetBounds(split.MainlandRegions.SelectMany(region => region.Parts));
            int columns = Math.Min(5, split.RemoteRegions.Count);
            int rows = (int)Math.Ceiling(split.RemoteRegions.Count / (double)columns);
            double gap = Math.Max(mainBounds.Height * 0.035d, 0.002d);
            double insetAreaHeight = mainBounds.Height * Math.Min(0.38d, 0.20d + (0.06d * rows));
            double rowHeight = insetAreaHeight / rows;
            double cellWidth = mainBounds.Width / columns;
            double insetTop = mainBounds.Bottom + gap;
            var replacements = new Dictionary<RenderRegion, RenderRegion>();

            List<RenderRegion> orderedRemoteRegions = split.RemoteRegions
                .OrderByDescending(region => region.Parts.Sum(GetPolygonArea))
                .ToList();
            for (int index = 0; index < orderedRemoteRegions.Count; index++)
            {
                RenderRegion region = orderedRemoteRegions[index];
                List<MapAssetPolygon> unwrappedParts = UnwrapParts(region.Parts, countries.AspectRatio);
                MapBounds sourceBounds = GetBounds(unwrappedParts);
                int row = index / columns;
                int column = index % columns;
                double cellLeft = mainBounds.Left + (column * cellWidth);
                double cellTop = insetTop + (row * rowHeight);
                double targetWidth = cellWidth * 0.82d;
                double targetHeight = rowHeight * 0.76d;
                double scale = Math.Min(
                    targetWidth / Math.Max(0.00001d, sourceBounds.Width),
                    targetHeight / Math.Max(0.00001d, sourceBounds.Height));
                double targetLeft = cellLeft + ((cellWidth - (sourceBounds.Width * scale)) / 2d);
                double targetTop = cellTop + ((rowHeight - (sourceBounds.Height * scale)) / 2d);
                replacements[region] = new RenderRegion(
                    region.Region,
                    TransformParts(unwrappedParts, sourceBounds, targetLeft, targetTop, scale));
            }

            return regions.Select(region => replacements.ContainsKey(region) ? replacements[region] : region).ToList();
        }

        private static List<RenderRegion> ExcludeRemoteRegions(List<RenderRegion> regions, string countryId)
        {
            AdministrativeRegionSplit split = ClassifyAdministrativeRegions(regions, countryId);
            return split != null && split.MainlandRegions.Count > 0 && split.RemoteRegions.Count > 0
                ? split.MainlandRegions
                : regions;
        }

        private static AdministrativeRegionSplit ClassifyAdministrativeRegions(List<RenderRegion> regions, string countryId)
        {
            if (regions == null || regions.Count < 2 || string.IsNullOrWhiteSpace(countryId))
            {
                return null;
            }

            MapAssetDefinition countries = LoadBundledAsset(CountriesResourceSuffix);
            MapAssetRegion country = countries?.Regions?.FirstOrDefault(region =>
                string.Equals(region.Id, countryId, StringComparison.OrdinalIgnoreCase));
            MapAssetPolygon mainland = country?.Parts?
                .Where(part => part?.Points != null && part.Points.Count >= 3)
                .OrderByDescending(GetPolygonArea)
                .FirstOrDefault();
            if (mainland == null)
            {
                return null;
            }

            MapBounds referenceBounds = GetBounds(new[] { mainland });
            double marginX = referenceBounds.Width * 0.08d;
            double marginY = referenceBounds.Height * 0.08d;
            var split = new AdministrativeRegionSplit();
            foreach (RenderRegion region in regions)
            {
                MapBounds bounds = GetBounds(region.Parts);
                double centerX = bounds.Left + (bounds.Width / 2d);
                double centerY = bounds.Top + (bounds.Height / 2d);
                bool belongsToMainland = centerX >= referenceBounds.Left - marginX &&
                                         centerX <= referenceBounds.Right + marginX &&
                                         centerY >= referenceBounds.Top - marginY &&
                                         centerY <= referenceBounds.Bottom + marginY;
                (belongsToMainland ? split.MainlandRegions : split.RemoteRegions).Add(region);
            }
            return split;
        }

        private static List<MapAssetPolygon> UnwrapParts(List<MapAssetPolygon> parts, double period)
        {
            List<MapAssetPoint> points = parts.SelectMany(part => part.Points).ToList();
            double wrapStart = GetLongitudeWrapStart("Country", points, period);
            return parts.Select(part => new MapAssetPolygon
            {
                Points = part.Points.Select(point => new MapAssetPoint
                {
                    X = UnwrapHorizontalCoordinate(point.X, wrapStart, period),
                    Y = point.Y
                }).ToList()
            }).ToList();
        }

        private static List<MapAssetPolygon> TransformParts(
            List<MapAssetPolygon> parts,
            MapBounds sourceBounds,
            double targetLeft,
            double targetTop,
            double scale)
        {
            return parts.Select(part => new MapAssetPolygon
            {
                Points = part.Points.Select(point => new MapAssetPoint
                {
                    X = targetLeft + ((point.X - sourceBounds.Left) * scale),
                    Y = targetTop + ((point.Y - sourceBounds.Top) * scale)
                }).ToList()
            }).ToList();
        }

        private static MapBounds GetBounds(IEnumerable<MapAssetPolygon> parts)
        {
            List<MapAssetPoint> points = parts
                .Where(part => part?.Points != null)
                .SelectMany(part => part.Points)
                .ToList();
            return new MapBounds(
                points.Min(point => point.X),
                points.Min(point => point.Y),
                points.Max(point => point.X),
                points.Max(point => point.Y));
        }

        private static List<MapAssetPolygon> SelectParts(MapAssetRegion region, bool largestLandmassOnly)
        {
            var parts = (region.Parts ?? new List<MapAssetPolygon>())
                .Where(part => part?.Points != null && part.Points.Count >= 3)
                .ToList();
            if (!largestLandmassOnly || parts.Count <= 1)
            {
                return parts;
            }

            MapAssetPolygon largest = parts.OrderByDescending(GetPolygonArea).First();
            return new List<MapAssetPolygon> { largest };
        }

        private static double GetPolygonArea(MapAssetPolygon polygon)
        {
            double area = 0;
            for (int index = 0; index < polygon.Points.Count; index++)
            {
                MapAssetPoint current = polygon.Points[index];
                MapAssetPoint next = polygon.Points[(index + 1) % polygon.Points.Count];
                area += (current.X * next.Y) - (next.X * current.Y);
            }
            return Math.Abs(area / 2d);
        }

        private static string SanitizeName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "Map";
            }

            return new string(value.Select(character => char.IsLetterOrDigit(character) ? character : '_').ToArray());
        }

        private static void TryDelete(PowerPoint.Shape shape)
        {
            try
            {
                shape?.Delete();
            }
            catch
            {
                // Rollback is best-effort because the generated shape may already have been deleted.
            }
        }

        private sealed class RenderRegion
        {
            public RenderRegion(MapAssetRegion region, List<MapAssetPolygon> parts)
            {
                Region = region;
                Parts = parts;
            }

            public MapAssetRegion Region { get; }
            public List<MapAssetPolygon> Parts { get; }
        }

        private sealed class AdministrativeRegionSplit
        {
            public List<RenderRegion> MainlandRegions { get; } = new List<RenderRegion>();
            public List<RenderRegion> RemoteRegions { get; } = new List<RenderRegion>();
        }

        private sealed class MapBounds
        {
            public MapBounds(double left, double top, double right, double bottom)
            {
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
            }

            public double Left { get; }
            public double Top { get; }
            public double Right { get; }
            public double Bottom { get; }
            public double Width => Math.Max(0.00001d, Right - Left);
            public double Height => Math.Max(0.00001d, Bottom - Top);
        }
    }
}
