using System;
using System.Collections.Generic;

namespace oPenEfficiency.Services
{
    public sealed class HtmlLayoutRect
    {
        public HtmlLayoutRect(float left, float top, float width, float height) { Left = left; Top = top; Width = width; Height = height; }
        public float Left { get; private set; } public float Top { get; private set; } public float Width { get; private set; } public float Height { get; private set; }
    }
    public sealed class HtmlSlideScale
    {
        public HtmlSlideScale(float x, float y, float offsetX, float offsetY) { X=x; Y=y; OffsetX=offsetX; OffsetY=offsetY; Average=(x+y)/2f; }
        public float X { get; private set; } public float Y { get; private set; } public float OffsetX { get; private set; } public float OffsetY { get; private set; } public float Average { get; private set; }
    }
    public sealed class HtmlLayoutEngine
    {
        public HtmlSlideScale Fit(float sourceWidth, float sourceHeight, float targetWidth, float targetHeight)
        {
            if (sourceWidth <= 0 || sourceHeight <= 0) throw new ArgumentOutOfRangeException();
            float ratio = Math.Min(targetWidth/sourceWidth, targetHeight/sourceHeight);
            return new HtmlSlideScale(ratio, ratio, (targetWidth-sourceWidth*ratio)/2f, (targetHeight-sourceHeight*ratio)/2f);
        }
        public IReadOnlyList<HtmlLayoutRect> Columns(HtmlLayoutRect area, int count, float gap)
        {
            var result = new List<HtmlLayoutRect>(); if (count <= 0) return result;
            float width = Math.Max(0, (area.Width-gap*(count-1))/count);
            for (int i=0;i<count;i++) result.Add(new HtmlLayoutRect(area.Left+i*(width+gap), area.Top, width, area.Height));
            return result;
        }
        public IReadOnlyList<HtmlLayoutRect> Grid(HtmlLayoutRect area, int itemCount, int columns, float columnGap, float rowGap)
        {
            var result = new List<HtmlLayoutRect>(); if (itemCount<=0 || columns<=0) return result;
            int rows=(int)Math.Ceiling(itemCount/(double)columns); float w=Math.Max(0,(area.Width-columnGap*(columns-1))/columns); float h=Math.Max(0,(area.Height-rowGap*(rows-1))/rows);
            for(int i=0;i<itemCount;i++) result.Add(new HtmlLayoutRect(area.Left+(i%columns)*(w+columnGap),area.Top+(i/columns)*(h+rowGap),w,h));
            return result;
        }
    }
}
