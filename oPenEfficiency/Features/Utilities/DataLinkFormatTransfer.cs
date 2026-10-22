// Shared transfer file. It is compiled into the PowerPoint add-in (oPenEfficiency.csproj, the
// refresh side that owns the transfer) and linked into the Excel add-in (oPenEfficiency.Xl.csproj,
// the creation side). Both product routes therefore run ONE cell-transfer implementation instead of
// two - values, font, fill, alignment and borders travel by the same code on creation
// (LinkRangeToPowerPointTableFeature.ExecuteCore with formatting) and on refresh
// (TableLinkFeature.RefreshTable). FEAT-005 DL-FORMAT-CREATE-002 pins that contract; the pure
// mapping decisions live in oPenEfficiency.Services.DataLinkFormatPolicy.
using System;
using System.Runtime.InteropServices;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using Xl = global::Microsoft.Office.Interop.Excel;
using Ppt = global::Microsoft.Office.Interop.PowerPoint;
using OfficeCore = global::Microsoft.Office.Core;

namespace oPenEfficiency.Services.DataLink
{
    /// <summary>
    /// One cell's transferable formatting, read from the Excel source cell and applied to the
    /// destination table cell. Defaults describe an unformatted cell, so a property Excel cannot
    /// report keeps the destination's default instead of writing a wrong value.
    /// </summary>
    public sealed class DataLinkCellFormat
    {
        public string FontName;
        public float FontSize;
        public bool Bold;
        public bool Italic;
        public bool Underline;
        public int FontColor;
        public bool HasFill;
        public int FillColor;
        public int HorizontalAlignment = DataLinkFormatPolicy.XlHAlignLeft;
        public int VerticalAlignment = DataLinkFormatPolicy.XlVAlignCenter;

        /// <summary>True when Excel's General alignment applies and the cell holds a number.</summary>
        public bool GeneralIsNumeric;

        public bool[] BorderVisible = new bool[4];
        public int[] BorderColor = new int[4];
        public float[] BorderWeight = new float[4];
    }

    /// <summary>
    /// The single Excel-to-PowerPoint cell transfer used by link creation and by link refresh.
    /// </summary>
    public static class DataLinkFormatTransfer
    {
        private static readonly Xl.XlBordersIndex[] BorderEdges =
        {
            Xl.XlBordersIndex.xlEdgeLeft,
            Xl.XlBordersIndex.xlEdgeRight,
            Xl.XlBordersIndex.xlEdgeTop,
            Xl.XlBordersIndex.xlEdgeBottom
        };

        private static readonly Ppt.PpBorderType[] BorderTargets =
        {
            Ppt.PpBorderType.ppBorderLeft,
            Ppt.PpBorderType.ppBorderRight,
            Ppt.PpBorderType.ppBorderTop,
            Ppt.PpBorderType.ppBorderBottom
        };

        /// <summary>The displayed (calculated) text of a source cell; empty when it is unavailable.</summary>
        public static string ReadDisplayedText(Xl.Range cell)
        {
            if (cell == null) return string.Empty;
            try { return Convert.ToString(cell.Text) ?? string.Empty; }
            catch (COMException) { return string.Empty; }
            catch (InvalidComObjectException) { return string.Empty; }
        }

        /// <summary>Reads everything the transfer can reproduce from a source cell.</summary>
        public static DataLinkCellFormat ReadCellFormat(Xl.Range cell)
        {
            var data = new DataLinkCellFormat();
            if (cell == null) return data;

            TryReadOptionalFormat(() => data.FontName = Convert.ToString(cell.Font.Name));
            TryReadOptionalFormat(() => data.FontSize = Convert.ToSingle(cell.Font.Size));
            TryReadOptionalFormat(() => data.Bold = DataLinkFormatPolicy.ToBoolean(cell.Font.Bold));
            TryReadOptionalFormat(() => data.Italic = DataLinkFormatPolicy.ToBoolean(cell.Font.Italic));
            TryReadOptionalFormat(() => data.Underline =
                Convert.ToInt32(cell.Font.Underline) != DataLinkFormatPolicy.XlLineStyleNone);
            TryReadOptionalFormat(() => data.FontColor = DataLinkFormatPolicy.ToBgr(cell.Font.Color));
            TryReadOptionalFormat(() => data.HasFill =
                DataLinkFormatPolicy.HasExplicitFill(Convert.ToInt32(cell.Interior.ColorIndex)));
            TryReadOptionalFormat(() => data.FillColor = DataLinkFormatPolicy.ToBgr(cell.Interior.Color));
            TryReadOptionalFormat(() => data.HorizontalAlignment = Convert.ToInt32(cell.HorizontalAlignment));
            TryReadOptionalFormat(() => data.VerticalAlignment = Convert.ToInt32(cell.VerticalAlignment));
            if (data.HorizontalAlignment == DataLinkFormatPolicy.XlHAlignGeneral)
            {
                // Excel's General alignment right-aligns numbers; the destination needs the same.
                TryReadOptionalFormat(() =>
                {
                    object value = cell.Value2;
                    data.GeneralIsNumeric = value is double || value is int;
                });
            }

            for (int i = 0; i < BorderEdges.Length; i++)
            {
                int edge = i;
                TryReadOptionalFormat(() =>
                {
                    Xl.Border border = cell.Borders[BorderEdges[edge]];
                    data.BorderVisible[edge] =
                        DataLinkFormatPolicy.HasVisibleBorder(Convert.ToInt32(border.LineStyle));
                    data.BorderColor[edge] = DataLinkFormatPolicy.ToBgr(border.Color);
                    data.BorderWeight[edge] = DataLinkFormatPolicy.MapBorderWeight(Convert.ToInt32(border.Weight));
                });
            }
            return data;
        }

        /// <summary>
        /// Writes one source cell into a destination table cell: the displayed text always, the
        /// formatting only for a link that was created with formatting. Returns false when the
        /// destination cell could not be written.
        /// </summary>
        public static bool ApplyCell(Xl.Range sourceCell, Ppt.Cell destinationCell, bool withFormatting)
        {
            if (sourceCell == null || destinationCell == null) return false;
            try
            {
                destinationCell.Shape.TextFrame.TextRange.Text = ReadDisplayedText(sourceCell);
                if (withFormatting) ApplyCellFormatting(ReadCellFormat(sourceCell), destinationCell);
                return true;
            }
            catch (COMException)
            {
                // A cell that cannot be written stays at its previous content.
                return false;
            }
            catch (InvalidComObjectException)
            {
                // A disconnected cell proxy is treated as an unwritable cell.
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DataLinkFormatTransfer.ApplyCell");
                return false;
            }
        }

        /// <summary>
        /// Transfers the whole source range into an existing table: every overlapping cell gets its
        /// displayed text, plus the Excel formatting when <paramref name="withFormatting"/> is set.
        /// Column widths are deliberately not part of the transfer - the table keeps its
        /// PowerPoint-side column layout. Returns the number of cells that were written.
        /// </summary>
        public static int ApplyTable(Xl.Range source, Ppt.Table table, bool withFormatting)
        {
            if (source == null || table == null) return 0;

            int sourceRows = source.Rows.Count;
            int sourceColumns = source.Columns.Count;
            int rows = Math.Min(table.Rows.Count, sourceRows);
            int columns = Math.Min(table.Columns.Count, sourceColumns);

            int written = 0;
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= columns; c++)
                {
                    try
                    {
                        if (ApplyCell((Xl.Range)source.Cells[r, c], table.Cell(r, c), withFormatting)) written++;
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "DataLinkFormatTransfer.ApplyTable.Cell");
                    }
                }
            }
            return written;
        }

        /// <summary>
        /// Applies one cell's formatting payload. Each aspect is wrapped individually so a single
        /// unsupported COM property never aborts the remaining formatting of that cell.
        /// </summary>
        public static void ApplyCellFormatting(DataLinkCellFormat data, Ppt.Cell cell)
        {
            if (data == null || cell == null) return;

            TryApplyOptionalFormat(() =>
            {
                Ppt.Font font = cell.Shape.TextFrame.TextRange.Font;
                if (!string.IsNullOrEmpty(data.FontName)) font.Name = data.FontName;
                if (data.FontSize > 0) font.Size = data.FontSize;
                font.Bold = data.Bold ? OfficeCore.MsoTriState.msoTrue : OfficeCore.MsoTriState.msoFalse;
                font.Italic = data.Italic ? OfficeCore.MsoTriState.msoTrue : OfficeCore.MsoTriState.msoFalse;
                font.Underline = data.Underline ? OfficeCore.MsoTriState.msoTrue : OfficeCore.MsoTriState.msoFalse;
                font.Color.RGB = data.FontColor;
            }, "DataLinkFormatTransfer.ApplyCellFormatting.Font");

            TryApplyOptionalFormat(
                () => ApplyFill(cell, DataLinkFillPolicy.ResolveFillBgr(data.HasFill, data.FillColor)),
                "DataLinkFormatTransfer.ApplyCellFormatting.Fill");

            TryApplyOptionalFormat(() =>
            {
                cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment =
                    ToParagraphAlignment(DataLinkFormatPolicy.MapAlignment(
                        data.HorizontalAlignment, data.GeneralIsNumeric));
            }, "DataLinkFormatTransfer.ApplyCellFormatting.HorizontalAlignment");

            TryApplyOptionalFormat(() =>
            {
                cell.Shape.TextFrame.VerticalAnchor =
                    ToVerticalAnchor(DataLinkFormatPolicy.MapVerticalAnchor(data.VerticalAlignment));
            }, "DataLinkFormatTransfer.ApplyCellFormatting.VerticalAlignment");

            for (int i = 0; i < BorderTargets.Length; i++)
            {
                int edge = i;
                TryApplyOptionalFormat(() =>
                {
                    Ppt.LineFormat border = cell.Borders[BorderTargets[edge]];
                    border.Visible = data.BorderVisible[edge]
                        ? OfficeCore.MsoTriState.msoTrue : OfficeCore.MsoTriState.msoFalse;
                    if (data.BorderVisible[edge])
                    {
                        border.ForeColor.RGB = data.BorderColor[edge];
                        border.Weight = data.BorderWeight[edge];
                    }
                }, "DataLinkFormatTransfer.ApplyCellFormatting.Border");
            }
        }

        private const int FillVerificationAttempts = 2;

        /// <summary>
        /// Writes one cell's fill and reads it back. The read-back is the point: a fill that does not
        /// stick used to be invisible, because a failing COM write is swallowed. The owner's deck
        /// showed the table style's own fill (rgb=15592167) on every cell while nothing complained -
        /// the same measurement the pre-fix creation route produced for its fixture cells. One bounded
        /// retry covers a transient COM state; a fill that still does not match is logged, so the
        /// acceptance runner's "no new error-log entries" assertion fails instead of the formatting
        /// loss passing unnoticed (FEAT-005 DL-FORMAT-CREATE-002).
        /// </summary>
        private static void ApplyFill(Ppt.Cell cell, int fillBgr)
        {
            for (int attempt = 1; attempt <= FillVerificationAttempts; attempt++)
            {
                cell.Shape.Fill.Visible = OfficeCore.MsoTriState.msoTrue;
                cell.Shape.Fill.Solid();
                cell.Shape.Fill.ForeColor.RGB = fillBgr;

                int actualVisible;
                int actualBgr;
                try
                {
                    actualVisible = (int)cell.Shape.Fill.Visible;
                    actualBgr = cell.Shape.Fill.ForeColor.RGB;
                }
                catch (COMException)
                {
                    continue;
                }
                catch (InvalidComObjectException)
                {
                    continue;
                }

                if (actualVisible == (int)OfficeCore.MsoTriState.msoTrue && actualBgr == fillBgr) return;

                if (attempt == FillVerificationAttempts)
                {
                    ExceptionLogger.Log(
                        new InvalidOperationException(
                            "The transferred cell fill did not stick: expected BGR " + fillBgr +
                            ", the cell reports visible=" + actualVisible + ", BGR " + actualBgr + "."),
                        "DataLinkFormatTransfer.ApplyCellFormatting.FillVerification");
                }
            }
        }

        private static Ppt.PpParagraphAlignment ToParagraphAlignment(DataLinkAlignment alignment)
        {
            switch (alignment)
            {
                case DataLinkAlignment.Center: return Ppt.PpParagraphAlignment.ppAlignCenter;
                case DataLinkAlignment.Right: return Ppt.PpParagraphAlignment.ppAlignRight;
                case DataLinkAlignment.Justify: return Ppt.PpParagraphAlignment.ppAlignJustify;
                default: return Ppt.PpParagraphAlignment.ppAlignLeft;
            }
        }

        private static OfficeCore.MsoVerticalAnchor ToVerticalAnchor(DataLinkVerticalAnchor anchor)
        {
            switch (anchor)
            {
                case DataLinkVerticalAnchor.Top: return OfficeCore.MsoVerticalAnchor.msoAnchorTop;
                case DataLinkVerticalAnchor.Bottom: return OfficeCore.MsoVerticalAnchor.msoAnchorBottom;
                default: return OfficeCore.MsoVerticalAnchor.msoAnchorMiddle;
            }
        }

        private static void TryReadOptionalFormat(Action action)
        {
            try
            {
                action();
            }
            catch (COMException)
            {
                // Mixed or unsupported Excel formatting keeps the payload default.
            }
            catch (InvalidComObjectException)
            {
                // A disconnected cell proxy keeps the payload default.
            }
            catch (InvalidCastException)
            {
                // Excel may expose mixed formatting as a non-convertible COM value.
            }
            catch (FormatException)
            {
                // Non-numeric formatting values keep the payload default.
            }
            catch (OverflowException)
            {
                // Formatting values outside the target type keep the payload default.
            }
        }

        private static void TryApplyOptionalFormat(Action action, string context)
        {
            try
            {
                action();
            }
            catch (COMException)
            {
                // Unsupported PowerPoint formatting is optional; the cell text stays valid.
            }
            catch (InvalidComObjectException)
            {
                // The cell proxy can disconnect while a table is being written.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, context);
            }
        }
    }
}
