using System;

namespace oPenEfficiency.Services
{
    public enum HtmlSlideRendererKind { GptDeck, GenericDom, Structured }
    public enum HtmlNodeRendererKind { Text, List, Image, Svg, Media, Table, Rule, Grid, Flex, Container }

    public sealed class HtmlRenderRouter
    {
        public HtmlSlideRendererKind GetSlideRenderer(HtmlImportSlideModel slide)
        {
            if (slide == null) throw new ArgumentNullException("slide");
            return slide.IsGptDeck ? HtmlSlideRendererKind.GptDeck : slide.IsGenericDom ? HtmlSlideRendererKind.GenericDom : HtmlSlideRendererKind.Structured;
        }

        public HtmlNodeRendererKind GetNodeRenderer(string tagName, string display)
        {
            string tag=(tagName??"").ToLowerInvariant(); string layout=(display??"").ToLowerInvariant();
            if(layout=="grid"||layout=="inline-grid") return HtmlNodeRendererKind.Grid;
            if(layout=="flex"||layout=="inline-flex") return HtmlNodeRendererKind.Flex;
            if(tag=="h1"||tag=="h2"||tag=="h3"||tag=="h4"||tag=="h5"||tag=="h6"||tag=="p"||tag=="blockquote") return HtmlNodeRendererKind.Text;
            if(tag=="ul"||tag=="ol") return HtmlNodeRendererKind.List;
            if(tag=="img") return HtmlNodeRendererKind.Image;
            if(tag=="svg") return HtmlNodeRendererKind.Svg;
            if(tag=="video"||tag=="audio"||tag=="iframe"||tag=="canvas") return HtmlNodeRendererKind.Media;
            if(tag=="table") return HtmlNodeRendererKind.Table;
            if(tag=="hr") return HtmlNodeRendererKind.Rule;
            return HtmlNodeRendererKind.Container;
        }
    }
}
