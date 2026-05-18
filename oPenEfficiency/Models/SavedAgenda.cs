using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace oPenEfficiency.Models
{
    [DataContract]
    public class SavedAgenda
    {
        [DataMember(Name = "name")]
        public string Name { get; set; }

        [DataMember(Name = "title")]
        public string Title { get; set; }

        [DataMember(Name = "subtitle")]
        public string Subtitle { get; set; }

        [DataMember(Name = "createdDate")]
        public DateTime CreatedDate { get; set; } = DateTime.Now;

        [DataMember(Name = "items")]
        public List<SavedAgendaItem> Items { get; set; } = new List<SavedAgendaItem>();
    }

    [DataContract]
    public class SavedAgendaItem
    {
        [DataMember(Name = "level")]
        public int Level { get; set; }

        [DataMember(Name = "topic")]
        public string Topic { get; set; }

        [DataMember(Name = "duration")]
        public int Duration { get; set; }

        [DataMember(Name = "format")]
        public ShapeFormatConfig Format { get; set; }

        [DataMember(Name = "shouldGenerate")]
        public bool ShouldGenerate { get; set; } = true;

        [DataMember(Name = "showMinutes")]
        public bool ShowMinutes { get; set; } = true;

        [DataMember(Name = "showResponsible")]
        public bool ShowResponsible { get; set; } = true;
    }
}
