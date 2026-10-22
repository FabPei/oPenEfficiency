using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace oPenEfficiency.Models
{
    /// <summary>
    /// Hand-authored metadata for the bundled illustration packs. One entry per source collection so
    /// license, attribution and bundling permissions stay data instead of hard-coded assumptions.
    /// The item list itself is derived from the SVG files by <c>IllustrationCatalog</c>.
    /// </summary>
    [DataContract]
    public class IllustrationManifest
    {
        [DataMember]
        public string Version { get; set; } = "1.0";

        [DataMember]
        public List<IllustrationCollectionInfo> Collections { get; set; } = new List<IllustrationCollectionInfo>();
    }

    [DataContract]
    public class IllustrationCollectionInfo
    {
        [DataMember]
        public string Id { get; set; }

        [DataMember]
        public string Name { get; set; }

        [DataMember]
        public string Description { get; set; }

        /// <summary>Short license name as stated by the source, e.g. "CC0 1.0 Universal".</summary>
        [DataMember]
        public string License { get; set; }

        [DataMember]
        public string LicenseUrl { get; set; }

        [DataMember]
        public string SourceUrl { get; set; }

        /// <summary>Revision/date marker of the downloaded bundle so asset updates stay deliberate.</summary>
        [DataMember]
        public string SourceRevision { get; set; }

        /// <summary>True when the license obliges visible credit on every use.</summary>
        [DataMember]
        public bool AttributionRequired { get; set; }

        /// <summary>Credit text to insert automatically when <see cref="AttributionRequired"/> is true.</summary>
        [DataMember]
        public string AttributionText { get; set; }

        /// <summary>True when the license allows shipping the assets inside the add-in.</summary>
        [DataMember]
        public bool AllowsBundling { get; set; }

        /// <summary>Relative folder below Resources\Illustrations that holds this collection's SVGs.</summary>
        [DataMember]
        public string RelativeFolder { get; set; }

        /// <summary>True when the pack's SVGs carry named parts that can be split and recoloured individually.</summary>
        [DataMember]
        public bool SupportsPartDecomposition { get; set; }
    }

    [DataContract]
    public class IllustrationPart
    {
        /// <summary>Full group id from the SVG, e.g. "pose/sitting/closed_legs-2".</summary>
        [DataMember]
        public string Id { get; set; }

        /// <summary>Display name of the leaf segment, e.g. "Closed Legs 2".</summary>
        [DataMember]
        public string Name { get; set; }

        /// <summary>Part family, e.g. Head, Hair, Face, FacialHair, Accessory, Pose, Container, Other.</summary>
        [DataMember]
        public string Role { get; set; }

        [DataMember]
        public int Depth { get; set; }

        /// <summary>True for "*-None" placeholder groups that carry no artwork.</summary>
        [DataMember]
        public bool IsPlaceholder { get; set; }

        [DataMember]
        public List<string> Colors { get; set; } = new List<string>();
    }

    [DataContract]
    public class IllustrationItem
    {
        [DataMember]
        public string Id { get; set; }

        [DataMember]
        public string Name { get; set; }

        [DataMember]
        public string CollectionId { get; set; }

        /// <summary>Path relative to Resources\Illustrations, e.g. "Peeps/open-peeps/peep-4.svg".</summary>
        [DataMember]
        public string RelativePath { get; set; }

        [DataMember]
        public List<string> Parts { get; set; } = new List<string>();

        [DataMember]
        public List<string> Tags { get; set; } = new List<string>();

        [DataMember]
        public List<string> Palette { get; set; } = new List<string>();

        /// <summary>True when the item can be split into its named parts after insertion.</summary>
        [DataMember]
        public bool Decomposable { get; set; }

        [IgnoreDataMember]
        public string SearchText
        {
            get
            {
                var parts = new List<string>();

                if (!string.IsNullOrWhiteSpace(Id)) parts.Add(Id);
                if (!string.IsNullOrWhiteSpace(Name)) parts.Add(Name);
                if (!string.IsNullOrWhiteSpace(CollectionId)) parts.Add(CollectionId);
                if (Tags != null && Tags.Count > 0) parts.AddRange(Tags);
                if (Parts != null && Parts.Count > 0) parts.AddRange(Parts);

                return string.Join(" ", parts).ToLowerInvariant();
            }
        }
    }
}
