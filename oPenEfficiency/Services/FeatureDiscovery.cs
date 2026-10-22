using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using oPenEfficiency.UI;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Auto-discovers and manages features using reflection.
    /// Scans the assembly for classes with [FeatureMetadata] attribute.
    /// </summary>
    public static class FeatureDiscovery
    {
        private static List<FeatureWrapper> _discoveredFeatures;
        private static bool _initialized;
        private static readonly object _lock = new object();

        /// <summary>
        /// Gets all discovered features. Initializes discovery on first access.
        /// </summary>
        public static List<FeatureWrapper> AllFeatures
        {
            get
            {
                if (!_initialized)
                {
                    Initialize();
                }
                return _discoveredFeatures;
            }
        }

        /// <summary>
        /// Initializes feature discovery by scanning the assembly.
        /// Called automatically on first access, but can be called manually.
        /// </summary>
        public static void Initialize()
        {
            lock (_lock)
            {
                if (_initialized) return;

                _discoveredFeatures = new List<FeatureWrapper>();

                try
                {
                    var assembly = Assembly.GetExecutingAssembly();
                    var featureTypes = GetLoadableTypes(assembly)
                        .Where(t => t.IsClass &&
                                    t.IsAbstract && // static classes are abstract in IL
                                    t.GetCustomAttribute<FeatureMetadataAttribute>() != null)
                        .OrderBy(t => t.Name)
                        .ToList();

                    var duplicateIds = new HashSet<string>(
                        featureTypes
                            .GroupBy(
                                type => type.GetCustomAttribute<FeatureMetadataAttribute>().Id,
                                StringComparer.Ordinal)
                            .Where(group => group.Count() > 1)
                            .Select(group => group.Key),
                        StringComparer.Ordinal);

                    foreach (var duplicateId in duplicateIds)
                    {
                        string typeNames = string.Join(", ", featureTypes
                            .Where(type => string.Equals(
                                type.GetCustomAttribute<FeatureMetadataAttribute>().Id,
                                duplicateId,
                                StringComparison.Ordinal))
                            .Select(type => type.FullName));
                        ExceptionLogger.Log(
                            new InvalidOperationException($"Duplicate feature ID '{duplicateId}': {typeNames}"),
                            "FeatureDiscovery.Initialize.DuplicateId");
                    }

                    foreach (var type in featureTypes)
                    {
                        var metadata = type.GetCustomAttribute<FeatureMetadataAttribute>();
                        if (duplicateIds.Contains(metadata.Id)) continue;

                        try
                        {
                            var wrapper = new FeatureWrapper(type);
                            _discoveredFeatures.Add(wrapper);
                        }
                        catch (Exception ex)
                        {
                            ExceptionLogger.Log(ex, $"FeatureDiscovery.Initialize.{type.Name}");
                        }
                    }

                    _initialized = true;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "FeatureDiscovery.Initialize");
                    _discoveredFeatures = new List<FeatureWrapper>();
                }
            }
        }

        /// <summary>
        /// Gets feature metadata by ID.
        /// </summary>
        public static FeatureDisplayInfo? GetFeatureInfo(string featureId)
        {
            if (!_initialized) Initialize();

            var feature = _discoveredFeatures.FirstOrDefault(f => f.Id == featureId);
            return feature?.DisplayInfo;
        }

        /// <summary>
        /// Gets the FeatureWrapper by ID.
        /// </summary>
        public static FeatureWrapper GetFeature(string featureId)
        {
            if (!_initialized) Initialize();

            return _discoveredFeatures.FirstOrDefault(f => f.Id == featureId);
        }

        /// <summary>
        /// Checks if a feature ID is registered.
        /// </summary>
        public static bool IsRegistered(string featureId)
        {
            if (!_initialized) Initialize();

            return _discoveredFeatures.Any(f => f.Id == featureId);
        }

        /// <summary>
        /// Resets discovery (useful for testing or hot-reload scenarios).
        /// </summary>
        public static void Reset()
        {
            lock (_lock)
            {
                _discoveredFeatures = null;
                _initialized = false;
            }
        }

        private static IEnumerable<Type> GetLoadableTypes(Assembly assembly)
        {
            try
            {
                return assembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                // Windows Runtime metadata can contain types that the legacy VSTO loader
                // cannot materialize. Keep discovering the types that did load instead of
                // disabling every feature in the sidebar.
                ExceptionLogger.Log(ex, "FeatureDiscovery.GetLoadableTypes");
                return ex.Types.Where(type => type != null);
            }
        }
    }
}
