using System;
using System.Collections.Generic;
using System.Linq;

namespace oPenEfficiency.Services
{
    public sealed class AgendaPlanItem
    {
        public string Topic { get; set; } public int Level { get; set; } public int Duration { get; set; }
        public bool IsBreak { get; set; } public bool ShouldGenerate { get; set; }
        public string DisplayIndex { get; set; } public string TimeSlot { get; set; }
    }
    public sealed class AgendaPreviewItem
    {
        public string Topic { get; set; } public string DisplayIndex { get; set; } public string TimeSlot { get; set; }
        public int Level { get; set; } public int Duration { get; set; }
    }
    public sealed class AgendaPlanResult
    {
        public AgendaPlanResult() { Preview = new List<AgendaPreviewItem>(); Issues = new List<string>(); }
        public List<AgendaPreviewItem> Preview { get; private set; } public List<string> Issues { get; private set; }
        public int TotalMinutes { get; set; }
    }
    public sealed class AgendaPlanningService
    {
        public bool TryFitDurations(
            IList<AgendaPlanItem> items,
            TimeSpan start,
            TimeSpan targetEnd)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));

            if (targetEnd <= start) targetEnd = targetEnd.Add(TimeSpan.FromDays(1));
            int budget = (int)Math.Round((targetEnd - start).TotalMinutes);
            var adjustableIndexes = items
                .Select((item, index) => new { item, index })
                .Where(entry => entry.item.ShouldGenerate && !entry.item.IsBreak)
                .Select(entry => entry.index)
                .ToList();
            int fixedMinutes = items
                .Where(item => item.ShouldGenerate && item.IsBreak)
                .Sum(item => Math.Max(0, item.Duration));
            int available = budget - fixedMinutes;

            if (adjustableIndexes.Count == 0 || available < adjustableIndexes.Count)
                return false;

            int currentAdjustableTotal = adjustableIndexes
                .Sum(index => Math.Max(1, items[index].Duration));
            var fittedDurations = new int[adjustableIndexes.Count];
            int assigned = 0;

            for (int fitIndex = 0; fitIndex < adjustableIndexes.Count; fitIndex++)
            {
                int remainingItems = adjustableIndexes.Count - fitIndex - 1;
                int itemIndex = adjustableIndexes[fitIndex];
                int proposed = (int)Math.Round(
                    available * (Math.Max(1, items[itemIndex].Duration) / (double)currentAdjustableTotal));
                int duration = fitIndex == adjustableIndexes.Count - 1
                    ? available - assigned
                    : Math.Max(1, Math.Min(proposed, available - assigned - remainingItems));
                fittedDurations[fitIndex] = duration;
                assigned += duration;
            }

            for (int fitIndex = 0; fitIndex < adjustableIndexes.Count; fitIndex++)
                items[adjustableIndexes[fitIndex]].Duration = fittedDurations[fitIndex];

            return true;
        }

        public AgendaPlanResult Build(IList<AgendaPlanItem> items, TimeSpan start, int startNumber, string timeFormat, TimeSpan? targetEnd)
        {
            if (items == null) throw new ArgumentNullException("items"); var result=new AgendaPlanResult(); int[] counters=new int[20]; counters[0]=startNumber-1; TimeSpan cursor=start;
            for(int i=0;i<items.Count;i++) { var item=items[i]; int level=Math.Max(0,Math.Min(19,item.Level));
                if(!item.IsBreak){ counters[level]++; for(int j=level+1;j<counters.Length;j++) counters[j]=0; item.DisplayIndex=level==0?counters[0].ToString():string.Join(".",counters.Take(level+1).Where(v=>v>0)); } else item.DisplayIndex="";
                var end=cursor+TimeSpan.FromMinutes(item.Duration); item.TimeSlot=DateTime.Today.Add(cursor).ToString(timeFormat)+" - "+DateTime.Today.Add(end).ToString(timeFormat); cursor=end;
                if(item.ShouldGenerate){ result.TotalMinutes+=Math.Max(0,item.Duration); result.Preview.Add(new AgendaPreviewItem{Topic=item.Topic,DisplayIndex=item.DisplayIndex,TimeSlot=item.TimeSlot,Level=item.Level,Duration=item.Duration}); }
                if(string.IsNullOrWhiteSpace(item.Topic)) result.Issues.Add("Topic "+(i+1)+" has no name."); if(item.ShouldGenerate&&item.Duration<=0) result.Issues.Add("'"+item.Topic+"' has no positive duration."); if(i==0&&item.Level>0) result.Issues.Add("The first topic cannot be a subtopic."); if(i>0&&item.Level>items[i-1].Level+1) result.Issues.Add("'"+item.Topic+"' skips a hierarchy level."); }
            foreach(var d in items.Where(x=>!string.IsNullOrWhiteSpace(x.Topic)).GroupBy(x=>x.Topic.Trim(),StringComparer.OrdinalIgnoreCase).Where(g=>g.Count()>1)) result.Issues.Add("The topic '"+d.Key+"' is duplicated.");
            if(targetEnd.HasValue){ var target=targetEnd.Value; if(target<=start) target=target.Add(TimeSpan.FromDays(1)); int budget=(int)Math.Round((target-start).TotalMinutes); if(result.TotalMinutes>budget) result.Issues.Add("The agenda exceeds its target end time by "+(result.TotalMinutes-budget)+" minutes."); }
            return result;
        }
    }
}
