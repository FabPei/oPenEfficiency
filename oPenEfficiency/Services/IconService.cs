using System;
using System.IO;
using System.Reflection;
using System.Windows.Media;
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

        public static Geometry GetIconGeometry(string id, out bool isFilled)
        {
            return GetIconGeometry(id, false, out isFilled);
        }

        public static Geometry GetIconGeometry(string id, bool hover, out bool isFilled)
        {
            isFilled = true; // Default for legacy path data

            string fileName = id + (hover ? "_hover" : "");
            
            // 1. Try to find SVG file
            string svgFilePath = Path.Combine(_assemblyDir, "UI", "Assets", "Icons", fileName + ".svg");

            // Development path fallback
            if (!File.Exists(svgFilePath))
            {
                string devPath = Path.GetFullPath(Path.Combine(_assemblyDir, "..", "..", "UI", "Assets", "Icons", fileName + ".svg"));
                if (File.Exists(devPath)) svgFilePath = devPath;
            }

            if (File.Exists(svgFilePath))
            {
                var svgGeom = SvgParser.ParseLucideSvg(svgFilePath, out isFilled);
                if (svgGeom != null && svgGeom.Children.Count > 0)
                {
                    return svgGeom;
                }
            }

            // If we're looking for a hover icon and didn't find one, return null (no fallback to default icon here)
            if (hover) return null;

            // 2. Fallback to FeatureLibrary path data (only for non-hover)
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
    }
}
