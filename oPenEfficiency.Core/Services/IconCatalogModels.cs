using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace oPenEfficiency.Models
{
    [DataContract]
    public class IconCollectionManifest
    {
        [DataMember]
        public string Version { get; set; } = "1.0";

        [DataMember]
        public string GeneratedAtUtc { get; set; } = DateTime.UtcNow.ToString("o");

        [DataMember]
        public List<IconCollectionInfo> Collections { get; set; } = new List<IconCollectionInfo>();

        [DataMember]
        public List<IconCatalogItem> Icons { get; set; } = new List<IconCatalogItem>();
    }

    [DataContract]
    public class IconCollectionInfo
    {
        [DataMember]
        public string Id { get; set; }

        [DataMember]
        public string Name { get; set; }

        [DataMember]
        public string Description { get; set; }

        [DataMember]
        public string License { get; set; }

        [DataMember]
        public string LicenseUrl { get; set; }

        [DataMember]
        public string SourceUrl { get; set; }

        [DataMember]
        public string SourceRevision { get; set; }

        [DataMember]
        public List<string> Styles { get; set; } = new List<string>();
    }

    [DataContract]
    public class IconCatalogItem
    {
        [DataMember]
        public string Id { get; set; }

        [DataMember]
        public string Name { get; set; }

        [DataMember]
        public string CollectionId { get; set; }

        [DataMember]
        public string Theme { get; set; }

        [DataMember]
        public string Style { get; set; }

        [DataMember]
        public string RelativePath { get; set; }

        [DataMember]
        public string License { get; set; }

        [DataMember]
        public string SourceUrl { get; set; }

        [DataMember]
        public List<string> Categories { get; set; } = new List<string>();

        [DataMember]
        public List<string> Tags { get; set; } = new List<string>();

        [IgnoreDataMember]
        public bool IsFavorite { get; set; }

        [IgnoreDataMember]
        public string SearchText
        {
            get
            {
                var parts = new List<string>();

                if (!string.IsNullOrWhiteSpace(Id)) parts.Add(Id);
                if (!string.IsNullOrWhiteSpace(Name)) parts.Add(Name);
                if (!string.IsNullOrWhiteSpace(CollectionId)) parts.Add(CollectionId);
                if (!string.IsNullOrWhiteSpace(Theme)) parts.Add(Theme);
                if (!string.IsNullOrWhiteSpace(Style)) parts.Add(Style);
                if (Categories != null && Categories.Count > 0) parts.AddRange(Categories);
                if (Tags != null && Tags.Count > 0) parts.AddRange(Tags);

                return string.Join(" ", parts).ToLowerInvariant();
            }
        }
    }

    [DataContract]
    public class IconFavoritesConfig
    {
        [DataMember]
        public List<string> FavoriteIconIds { get; set; } = new List<string>();
    }
}
