using System;
using System.IO;
using System.Reflection;
using System.Windows.Media;
using oPenEfficiency.UI;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    public static class IconService
    {
        private static string _assemblyDir;

        static IconService()
        {
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            _assemblyDir = Path.GetDirectoryName(Uri.UnescapeDataString(uri.Path));
        }

        /// <summary>
        /// Role-aware icon geometry loader. Returns a <see cref="ParsedIconGeometry"/> with
        /// separate base and accent layers parsed from <c>data-oe-role</c> SVG attributes.
        /// Returns null if the icon cannot be found or parsed.
        /// </summary>
        public static ParsedIconGeometry GetIconGeometryWithRoles(string id)
        {
            string svgFilePath = ResolveIconPath(id);
            if (svgFilePath == null) return null;

            var parsed = SvgParser.ParseSvgWithRoles(svgFilePath);
            return parsed;
        }

        /// <summary>
        /// Legacy single-layer icon geometry loader. Used as fallback when the role-aware parser
        /// returns no result, or for consumers that don't need two-layer rendering.
        /// </summary>
        public static Geometry GetIconGeometryLegacy(string id, out bool isFilled)
        {
            isFilled = true; // Default for legacy path data

            string svgFilePath = ResolveIconPath(id);
            if (svgFilePath != null)
            {
                var svgGeom = SvgParser.ParseLucideSvg(svgFilePath, out isFilled);
                if (svgGeom != null && svgGeom.Children.Count > 0)
                {
                    return svgGeom;
                }
            }

            // Fallback to FeatureLibrary path data
            var info = UI.FeatureLibrary.GetFeatureInfo(id);
            if (!string.IsNullOrEmpty(info.IconData))
            {
                try
                {
                    isFilled = true; // Legacy icons are always filled
                    return Geometry.Parse(info.IconData);
                }
                catch
                {
                    return null;
                }
            }

            return null;
        }

        /// <summary>
        /// Backward-compatible overload for consumers still using the legacy API.
        /// </summary>
        public static Geometry GetIconGeometry(string id, out bool isFilled)
        {
            return GetIconGeometryLegacy(id, out isFilled);
        }

        private static string ResolveIconPath(string id)
        {
            // 1. Try to find SVG file (no hover variants)
            string svgFilePath = Path.Combine(_assemblyDir, "UI", "Assets", "Icons", id + ".svg");

            // Development path fallback
            if (!File.Exists(svgFilePath))
            {
                string devPath = Path.GetFullPath(Path.Combine(_assemblyDir, "..", "..", "UI", "Assets", "Icons", id + ".svg"));
                if (File.Exists(devPath)) svgFilePath = devPath;
            }

            return File.Exists(svgFilePath) ? svgFilePath : null;
        }
    }
}
