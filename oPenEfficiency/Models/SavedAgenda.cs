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

        [DataMember(Name = "startTime")]
        public string StartTime { get; set; } = "09:00";

        [DataMember(Name = "targetEndTime")]
        public string TargetEndTime { get; set; } = "10:30";

        [DataMember(Name = "startNumber")]
        public int StartNumber { get; set; } = 1;

        [DataMember(Name = "timeFormat")]
        public string TimeFormat { get; set; } = "HH:mm";

        [DataMember(Name = "items")]
        public List<SavedAgendaItem> Items { get; set; } = new List<SavedAgendaItem>();
    }

    [DataContract]
    public class SavedAgendaItem
    {
        [DataMember(Name = "id")]
        public string Id { get; set; } = Guid.NewGuid().ToString("D");

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

        [DataMember(Name = "responsible")]
        public string Responsible { get; set; }

        [DataMember(Name = "isBreak")]
        public bool IsBreak { get; set; }
    }
}
