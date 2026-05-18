using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace oPenEfficiency.Models
{
    [DataContract]
    public class AgendaLayoutConfig
    {
        [DataMember(Name = "layoutName")]
        public string LayoutName { get; set; }

        // Slide level shapes
        [DataMember(Name = "slideShape")]
        public ShapeFormatConfig SlideShape { get; set; }

        [DataMember(Name = "headlineShape")]
        public ShapeFormatConfig HeadlineShape { get; set; }

        [DataMember(Name = "subHeadlineShape")]
        public ShapeFormatConfig SubHeadlineShape { get; set; }

        // Main item shapes
        [DataMember(Name = "itemNumberingShape")]
        public ShapeFormatConfig ItemNumberingShape { get; set; }

        [DataMember(Name = "itemPageNumberShape")]
        public ShapeFormatConfig ItemPageNumberShape { get; set; }

        [DataMember(Name = "itemTextShape")]
        public ShapeFormatConfig ItemTextShape { get; set; }

        [DataMember(Name = "itemBackgroundShape")]
        public ShapeFormatConfig ItemBackgroundShape { get; set; }

        [DataMember(Name = "itemTimeShape")]
        public ShapeFormatConfig ItemTimeShape { get; set; }

        [DataMember(Name = "itemDurationShape")]
        public ShapeFormatConfig ItemDurationShape { get; set; }

        [DataMember(Name = "itemResponsibleShape")]
        public ShapeFormatConfig ItemResponsibleShape { get; set; }

        // Sub-item shapes (optional)
        [DataMember(Name = "subItemNumberingShape")]
        public ShapeFormatConfig SubItemNumberingShape { get; set; }

        [DataMember(Name = "subItemPageNumberShape")]
        public ShapeFormatConfig SubItemPageNumberShape { get; set; }

        [DataMember(Name = "subItemTextShape")]
        public ShapeFormatConfig SubItemTextShape { get; set; }

        [DataMember(Name = "subItemBackgroundShape")]
        public ShapeFormatConfig SubItemBackgroundShape { get; set; }

        [DataMember(Name = "subItemTimeShape")]
        public ShapeFormatConfig SubItemTimeShape { get; set; }

        [DataMember(Name = "subItemDurationShape")]
        public ShapeFormatConfig SubItemDurationShape { get; set; }

        [DataMember(Name = "subItemResponsibleShape")]
        public ShapeFormatConfig SubItemResponsibleShape { get; set; }

        // Settings
        [DataMember(Name = "listAlignment")]
        public string ListAlignment { get; set; } = "Left"; // "Left", "Center", "Right"

        [DataMember(Name = "listOrientation")]
        public string ListOrientation { get; set; } = "Vertical"; // "Vertical", "Horizontal"

        [DataMember(Name = "itemSpacing")]
        public float ItemSpacing { get; set; } = 15f;

        // Align agenda item and info middle
        [DataMember(Name = "alignItemsMiddle")]
        public bool AlignItemsMiddle { get; set; } = false;

        // Align info shapes and numbering with item text (vertical orientation only)
        [DataMember(Name = "alignInfoShapesWithItem")]
        public bool AlignInfoShapesWithItem { get; set; } = false;

        // Highlight formatting for current agenda item
        [DataMember(Name = "highlightBold")]
        public bool HighlightBold { get; set; } = false;

        [DataMember(Name = "highlightItalic")]
        public bool HighlightItalic { get; set; } = false;

        [DataMember(Name = "highlightUnderline")]
        public bool HighlightUnderline { get; set; }

        [DataMember(Name = "highlightFontColor")]
        public bool HighlightFontColor { get; set; }

        [DataMember(Name = "highlightFontColorRgb")]
        public int HighlightFontColorRgb { get; set; }

        [DataMember(Name = "highlightFillColor")]
        public bool HighlightFillColor { get; set; }

        [DataMember(Name = "highlightFillColorRgb")]
        public int HighlightFillColorRgb { get; set; }
    }

    [DataContract]
    public class ShapeFormatConfig
    {
        [DataMember(Name = "shapeId")]
        public int ShapeId { get; set; }

        // Dimensions
        [DataMember(Name = "top")]
        public float Top { get; set; }

        [DataMember(Name = "left")]
        public float Left { get; set; }

        [DataMember(Name = "width")]
        public float Width { get; set; }

        [DataMember(Name = "height")]
        public float Height { get; set; }

        // Background
        [DataMember(Name = "hasFill")]
        public bool HasFill { get; set; }

        [DataMember(Name = "fillColorRgb")]
        public int FillColorRgb { get; set; }

        // Border
        [DataMember(Name = "hasLine")]
        public bool HasLine { get; set; }

        [DataMember(Name = "lineColorRgb")]
        public int LineColorRgb { get; set; }

        [DataMember(Name = "lineWeight")]
        public float LineWeight { get; set; }

        // Text & Font
        [DataMember(Name = "fontName")]
        public string FontName { get; set; }

        [DataMember(Name = "fontSize")]
        public float FontSize { get; set; }

        [DataMember(Name = "fontColorRgb")]
        public int FontColorRgb { get; set; }

        [DataMember(Name = "fontBold")]
        public bool FontBold { get; set; }

        [DataMember(Name = "fontItalic")]
        public bool FontItalic { get; set; }

        [DataMember(Name = "textAlignment")]
        public int TextAlignment { get; set; } // PpParagraphAlignment

        // Text Box Margins
        [DataMember(Name = "marginLeft")]
        public float MarginLeft { get; set; }

        [DataMember(Name = "marginRight")]
        public float MarginRight { get; set; }

        [DataMember(Name = "marginTop")]
        public float MarginTop { get; set; }

        [DataMember(Name = "marginBottom")]
        public float MarginBottom { get; set; }

        [DataMember(Name = "orientation")]
        public int Orientation { get; set; } // MsoTextOrientation
    }
}
