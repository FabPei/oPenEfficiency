using System.Collections.Generic;
using System.Runtime.Serialization;

namespace oPenEfficiency.Models
{
    [DataContract]
    public class TemplateItem
    {
        [DataMember]
        public string Name { get; set; }
        [DataMember]
        public string Path { get; set; }
    }

    [DataContract]
    public class TemplateConfig
    {
        [DataMember]
        public List<TemplateItem> CustomTemplates { get; set; } = new List<TemplateItem>();
        [DataMember]
        public List<string> WatchedDirectories { get; set; } = new List<string>();
        [DataMember]
        public bool HideDefaultLocations { get; set; } = false;
        [DataMember]
        public List<string> BlacklistedPaths { get; set; } = new List<string>();
    }
}
