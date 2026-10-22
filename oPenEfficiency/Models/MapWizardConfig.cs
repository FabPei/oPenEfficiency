using System.Collections.Generic;

namespace oPenEfficiency.Models
{
    public enum MapColorMode
    {
        Continuous, // Heatmap interpolation between min and max colors
        Categorical // Specific color mapping per discrete label
    }

    public enum MapProjection
    {
        Mercator,
        Natural
    }

    public enum MapRemoteRegionMode
    {
        Geographic,
        CompactInsets,
        Exclude
    }

    public class MapDataRow
    {
        public string Region { get; set; }
        public string RawValue { get; set; }
        public double NumericValue { get; set; }
        public bool IsMatched { get; set; } // True if the region name matches a shape in the group
    }

    public class MapWizardConfig
    {
        public MapColorMode Mode { get; set; } = MapColorMode.Continuous;
        public MapProjection Projection { get; set; } = MapProjection.Mercator;
        public MapRemoteRegionMode RemoteRegionMode { get; set; } = MapRemoteRegionMode.Geographic;
        
        // Colors for Continuous Mode
        public string LowColorHex { get; set; } = "#E3F2FD"; // Light Blue
        public string HighColorHex { get; set; } = "#0D47A1"; // Dark Blue
        public string DefaultColorHex { get; set; } = "#D3D3D3"; // Gray for missing data

        // Colors for Categorical Mode
        public Dictionary<string, string> CategoryColors { get; set; } = new Dictionary<string, string>();

        public List<MapDataRow> Data { get; set; } = new List<MapDataRow>();
    }
}
