using System;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Horizontal destination alignment for a transferred linked-table cell. Host-neutral so the
    /// policy can live in <c>oPenEfficiency.Core</c> and be shared by the Excel-side creation and
    /// the PowerPoint-side refresh (FEAT-005 DL-FORMAT-NOFILL-001 / DL-FORMAT-CREATE-002).
    /// </summary>
    public enum DataLinkAlignment
    {
        Left,
        Center,
        Right,
        Justify
    }

    /// <summary>
    /// Vertical destination anchor for a transferred linked-table cell (host-neutral).
    /// </summary>
    public enum DataLinkVerticalAnchor
    {
        Top,
        Middle,
        Bottom
    }

    /// <summary>
    /// Pure part of the Excel-to-PowerPoint cell-format transfer: it turns the Excel formatting
    /// values (plain <see cref="int"/> enum values and boxed OLE colors) into the decisions the
    /// host-side transfer applies. No Office interop type appears here, so both add-ins and the
    /// unit suite can use it and a value that transfers differently on creation than on refresh
    /// cannot be explained by two different implementations of the mapping.
    ///
    /// The Excel enum values are reproduced as documented constants (verified against
    /// the Excel 15.0 interop assemblies) because a host-neutral assembly may not
    /// reference the interop assemblies.
    /// </summary>
    public static class DataLinkFormatPolicy
    {
        // Excel XlHAlign
        public const int XlHAlignGeneral = 1;
        public const int XlHAlignLeft = -4131;
        public const int XlHAlignCenter = -4108;
        public const int XlHAlignCenterAcrossSelection = 7;
        public const int XlHAlignRight = -4152;
        public const int XlHAlignJustify = -4130;
        public const int XlHAlignDistributed = -4117;

        // Excel XlVAlign
        public const int XlVAlignTop = -4160;
        public const int XlVAlignCenter = -4108;
        public const int XlVAlignBottom = -4107;

        // Excel XlBorderWeight
        public const int XlBorderWeightHairline = 1;
        public const int XlBorderWeightThin = 2;
        public const int XlBorderWeightMedium = -4138;
        public const int XlBorderWeightThick = 4;

        // Excel XlLineStyle.xlLineStyleNone and Constants.xlNone.
        public const int XlLineStyleNone = -4142;

        /// <summary>True when Excel holds a solid interior for the cell (ColorIndex != xlNone).</summary>
        public static bool HasExplicitFill(int colorIndex) => colorIndex != XlLineStyleNone;

        /// <summary>True when the Excel border draws a line (LineStyle != xlLineStyleNone).</summary>
        public static bool HasVisibleBorder(int lineStyle) => lineStyle != XlLineStyleNone;

        /// <summary>
        /// Excel horizontal alignment to destination alignment. <c>xlHAlignGeneral</c> means
        /// "Excel decides": numbers are right-aligned and everything else left-aligned, so the
        /// caller passes what Excel shows for the cell.
        /// </summary>
        public static DataLinkAlignment MapAlignment(int xlHAlign, bool generalIsNumeric)
        {
            switch (xlHAlign)
            {
                case XlHAlignCenter:
                case XlHAlignCenterAcrossSelection:
                    return DataLinkAlignment.Center;
                case XlHAlignRight:
                    return DataLinkAlignment.Right;
                case XlHAlignJustify:
                case XlHAlignDistributed:
                    return DataLinkAlignment.Justify;
                case XlHAlignLeft:
                    return DataLinkAlignment.Left;
                default:
                    return generalIsNumeric ? DataLinkAlignment.Right : DataLinkAlignment.Left;
            }
        }

        /// <summary>Excel vertical alignment to destination anchor (unknown values keep middle).</summary>
        public static DataLinkVerticalAnchor MapVerticalAnchor(int xlVAlign)
        {
            switch (xlVAlign)
            {
                case XlVAlignTop: return DataLinkVerticalAnchor.Top;
                case XlVAlignBottom: return DataLinkVerticalAnchor.Bottom;
                default: return DataLinkVerticalAnchor.Middle;
            }
        }

        /// <summary>Excel border weight to an approximate line weight in points.</summary>
        public static float MapBorderWeight(int xlBorderWeight)
        {
            switch (xlBorderWeight)
            {
                case XlBorderWeightHairline: return 0.5f;
                case XlBorderWeightThin: return 1.0f;
                case XlBorderWeightMedium: return 2.0f;
                case XlBorderWeightThick: return 3.0f;
                default: return 1.0f;
            }
        }

        /// <summary>
        /// Normalises a boxed Excel OLE color (both Excel and PowerPoint use a BGR-packed
        /// COLORREF) to <see cref="int"/>. Excel returns the color as a double, so a direct cast
        /// would be wrong; non-convertible values resolve to 0.
        /// </summary>
        public static int ToBgr(object oleColor)
        {
            try { return unchecked((int)Convert.ToUInt32(oleColor)); }
            catch (InvalidCastException) { return 0; }
            catch (FormatException) { return 0; }
            catch (OverflowException) { return 0; }
        }

        /// <summary>
        /// Excel returns Bold/Italic as a boxed bool and null for a mixed range; anything that is
        /// not clearly true stays false so a mixed cell does not turn bold.
        /// </summary>
        public static bool ToBoolean(object value)
        {
            try { return Convert.ToBoolean(value); }
            catch (InvalidCastException) { return false; }
            catch (FormatException) { return false; }
        }
    }
}
