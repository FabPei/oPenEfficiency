using System.Collections.Generic;
using System.Runtime.Serialization;

namespace oPenEfficiency.Models
{
    [DataContract]
    public sealed class MapAssetDefinition
    {
        [DataMember] public string MapId { get; set; }
        [DataMember] public string DisplayName { get; set; }
        [DataMember] public string Source { get; set; }
        [DataMember] public string SourceVersion { get; set; }
        [DataMember] public string SourceUrl { get; set; }
        [DataMember] public double AspectRatio { get; set; }
        [DataMember] public string CoordinateSpace { get; set; }
        [DataMember] public double? MinLatitude { get; set; }
        [DataMember] public double? MaxLatitude { get; set; }
        [DataMember] public List<MapAssetRegion> Regions { get; set; } = new List<MapAssetRegion>();
    }

    [DataContract]
    public sealed class MapAssetRegion
    {
        [DataMember] public string Id { get; set; }
        [DataMember] public string Name { get; set; }
        [DataMember] public string EnglishName { get; set; }
        [DataMember] public string Group { get; set; }
        [DataMember] public string ParentGroup { get; set; }
        [DataMember] public string EntityType { get; set; }
        [DataMember] public List<string> Aliases { get; set; } = new List<string>();
        [DataMember] public List<MapAssetPolygon> Parts { get; set; } = new List<MapAssetPolygon>();
    }

    public sealed class MapCatalogEntry
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string LocalName { get; set; }
        public string Group { get; set; }
        public string EntityType { get; set; }
    }

    [DataContract]
    public sealed class MapAssetPolygon
    {
        [DataMember] public List<MapAssetPoint> Points { get; set; } = new List<MapAssetPoint>();
    }

    [DataContract]
    public sealed class MapAssetPoint
    {
        [DataMember] public double X { get; set; }
        [DataMember] public double Y { get; set; }
    }
}
