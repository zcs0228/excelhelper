﻿using System.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;

namespace ExcelHelper.OpenXML
{
    public class CustomStylesheet : Stylesheet
    {
        public enum CustomCellFormats : uint
        {
            // these are referenced by index, must be added in this order
            DefaultText = 0,
            DefaultDate,
            DefaultNumber2DecimalPlace,
            DefaultNumber5DecimalPlace,
            DefaultDateTime,
            HeaderText,
            TotalsNumber,
            TotalsNumber2DecimalPlace,
            TotalsText,
            TitleText,
            SubtitleText,
            Duration,
            TotalsDuration,
            Hyperlink
        }

        public CustomStylesheet()
        {
            NumberingFormat nfDateTime;
            NumberingFormat nf5Decimal;
            NumberingFormat nfDuration;
            NumberingFormat nfTotalDuration;

            Append(CreateNumberingFormats(out nfDateTime, out nf5Decimal, out nfDuration, out nfTotalDuration));
            Append(CreateFonts());
            Append(CreateFills());
            Append(CreateBorders());
            Append(CreateCellStyleFormats());
            Append(CreateCellFormats(nfDateTime, nf5Decimal, nfDuration, nfTotalDuration));
            Append(CreateCellStyles());
            Append(CreateDifferentialFormats());
            Append(CreateTableStyles());
        }

        private static TableStyles CreateTableStyles()
        {
            var tss = new TableStyles();
            tss.Count = 0;
            tss.DefaultTableStyle = StringValue.FromString("TableStyleMedium9");
            tss.DefaultPivotStyle = StringValue.FromString("PivotStyleLight16");
            return tss;
        }

        private static DifferentialFormats CreateDifferentialFormats()
        {
            var dfs = new DifferentialFormats();
            dfs.Count = 0;
            return dfs;
        }

        private static CellStyles CreateCellStyles()
        {
            var css = new CellStyles();

            // cell style 0
            var cs = new CellStyle();
            cs.Name = StringValue.FromString("Normal");
            cs.FormatId = 0;
            cs.BuiltinId = 0;
            css.AppendChild(cs);
            css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);
            return css;
        }

        /// <summary>
        /// Ensure cell formats are added in the order specified by the enumeration
        /// </summary>
        private static CellFormats CreateCellFormats(NumberingFormat nfDateTime, NumberingFormat nf5Decimal,
            NumberingFormat nfDuration, NumberingFormat nfTotalDuration)
        {
            var cfs = new CellFormats();

            // CustomCellFormats.DefaultText
            var cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(false);
            cfs.AppendChild(cf);

            // CustomCellFormats.DefaultDate
            cf = new CellFormat();
            cf.NumberFormatId = 14; // mm-dd-yy
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cfs.AppendChild(cf);

            // CustomCellFormats.DefaultNumber2DecimalPlace
            cf = new CellFormat();
            cf.NumberFormatId = 4; // #,##0.00
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cfs.AppendChild(cf);

            // CustomCellFormats.DefaultNumber5DecimalPlace
            cf = new CellFormat();
            cf.NumberFormatId = nf5Decimal.NumberFormatId;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cfs.AppendChild(cf);

            // CustomCellFormats.DefaultDateTime
            cf = new CellFormat();
            cf.NumberFormatId = nfDateTime.NumberFormatId;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cfs.AppendChild(cf);

            // CustomCellFormats.HeaderText
            cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 1;
            cf.FillId = 2;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(false);
            cfs.AppendChild(cf);

            // CustomCellFormats.TotalsNumber
            cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 3;
            cf.BorderId = 2;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cfs.AppendChild(cf);

            // CustomCellFormats.TotalsNumber2DecimalPlace
            cf = new CellFormat();
            cf.NumberFormatId = 4; // #,##0.00
            cf.FontId = 0;
            cf.FillId = 3;
            cf.BorderId = 2;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cfs.AppendChild(cf);

            // CustomCellFormats.TotalsText
            cf = new CellFormat();
            cf.NumberFormatId = 49; // @
            cf.FontId = 0;
            cf.FillId = 3;
            cf.BorderId = 2;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cfs.AppendChild(cf);

            // CustomCellFormats.TitleText
            cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 2;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(false);
            cf.Alignment = new Alignment
            {
                Vertical = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Bottom)
            };
            cfs.AppendChild(cf);

            // CustomCellFormats.SubtitleText
            cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 3;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(false);
            cf.Alignment = new Alignment
            {
                Vertical = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Top)
            };
            cfs.AppendChild(cf);

            // CustomCellFormats.Duration
            cf = new CellFormat();
            cf.NumberFormatId = nfDuration.NumberFormatId; // [h]:mm
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cf.Alignment = new Alignment
            {
                Horizontal = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Right)
            };
            cfs.AppendChild(cf);

            // CustomCellFormats.TotalsNumber
            cf = new CellFormat();
            cf.NumberFormatId = nfTotalDuration.NumberFormatId; // d:h:mm
            cf.FontId = 0;
            cf.FillId = 3;
            cf.BorderId = 2;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cf.Alignment = new Alignment
            {
                Horizontal = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Right)
            };
            cfs.AppendChild(cf);

            // CustomCellFormats.Hyperlink
            cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 4;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(false);
            cfs.AppendChild(cf);

            cfs.Count = UInt32Value.FromUInt32((uint)cfs.ChildElements.Count);
            return cfs;
        }

        private static NumberingFormats CreateNumberingFormats(out NumberingFormat nfDateTime,
            out NumberingFormat nf5Decimal, out NumberingFormat nfDuration, out NumberingFormat nfTotalDuration)
        {
            // built-in formats go up to 164
            uint iExcelIndex = 164;

            var nfs = new NumberingFormats();
            nfDateTime = new NumberingFormat();
            nfDateTime.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nfDateTime.FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss");
            nfs.AppendChild(nfDateTime);

            nf5Decimal = new NumberingFormat();
            nf5Decimal.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nf5Decimal.FormatCode = StringValue.FromString("#,##0.00000");
            nfs.AppendChild(nf5Decimal);

            nfDuration = new NumberingFormat();
            nfDuration.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nfDuration.FormatCode = StringValue.FromString("[h]:mm");
            nfs.AppendChild(nfDuration);

            nfTotalDuration = new NumberingFormat();
            nfTotalDuration.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nfTotalDuration.FormatCode = StringValue.FromString("d:h:mm");
            nfs.AppendChild(nfTotalDuration);

            nfs.Count = UInt32Value.FromUInt32((uint)nfs.ChildElements.Count);
            return nfs;
        }

        private static CellStyleFormats CreateCellStyleFormats()
        {
            var csfs = new CellStyleFormats();

            // cell style 0
            var cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            csfs.AppendChild(cf);
            csfs.Count = UInt32Value.FromUInt32((uint)csfs.ChildElements.Count);
            return csfs;
        }

        private static Borders CreateBorders()
        {
            var borders = new Borders();

            // boarder index 0
            var border = new Border();
            border.LeftBorder = new LeftBorder();
            border.RightBorder = new RightBorder();
            border.TopBorder = new TopBorder();
            border.BottomBorder = new BottomBorder();
            border.DiagonalBorder = new DiagonalBorder();
            borders.AppendChild(border);

            // boarder Index 1
            border = new Border();
            border.LeftBorder = new LeftBorder();
            border.LeftBorder.Style = BorderStyleValues.Thin;
            border.RightBorder = new RightBorder();
            border.RightBorder.Style = BorderStyleValues.Thin;
            border.TopBorder = new TopBorder();
            border.TopBorder.Style = BorderStyleValues.Thin;
            border.BottomBorder = new BottomBorder();
            border.BottomBorder.Style = BorderStyleValues.Thin;
            border.DiagonalBorder = new DiagonalBorder();
            borders.AppendChild(border);

            // boarder Index 2
            border = new Border();
            border.LeftBorder = new LeftBorder();
            border.RightBorder = new RightBorder();
            border.TopBorder = new TopBorder();
            border.TopBorder.Style = BorderStyleValues.Thin;
            border.BottomBorder = new BottomBorder();
            border.BottomBorder.Style = BorderStyleValues.Thin;
            border.DiagonalBorder = new DiagonalBorder();
            borders.AppendChild(border);

            borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);
            return borders;
        }

        private static Fills CreateFills()
        {
            // fill 0
            var fills = new Fills();
            var fill = new Fill();
            var patternFill = new PatternFill { PatternType = PatternValues.None };
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);

            // fill 1 (in-built fill)
            fill = new Fill();
            patternFill = new PatternFill { PatternType = PatternValues.Gray125 };
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);

            // fill 2
            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Solid;
            var fillColor = Color.LightSkyBlue;
            patternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValueFromColor(fillColor) };
            patternFill.BackgroundColor = new BackgroundColor { Rgb = HexBinaryValueFromColor(fillColor) };
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);

            // fill 3
            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Solid;
            fillColor = Color.Orange;
            patternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValueFromColor(fillColor) };
            patternFill.BackgroundColor = new BackgroundColor { Rgb = HexBinaryValueFromColor(fillColor) };
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);

            fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);
            return fills;
        }

        private static Fonts CreateFonts()
        {
            var fts = new Fonts();

            // font 0
            var ft = new Font();
            var ftn = new FontName { Val = StringValue.FromString("Arial") };
            var ftsz = new FontSize { Val = DoubleValue.FromDouble(11) };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            fts.AppendChild(ft);

            // font 1
            ft = new Font();
            ftn = new FontName { Val = StringValue.FromString("Arial") };
            ftsz = new FontSize { Val = DoubleValue.FromDouble(12) };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            ft.Bold = new Bold();
            fts.AppendChild(ft);

            // font 2
            ft = new Font();
            ftn = new FontName { Val = StringValue.FromString("Arial") };
            ftsz = new FontSize { Val = DoubleValue.FromDouble(18) };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            ft.Bold = new Bold();
            fts.AppendChild(ft);

            // font 3
            ft = new Font();
            ftn = new FontName { Val = StringValue.FromString("Arial") };
            ftsz = new FontSize { Val = DoubleValue.FromDouble(14) };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            fts.AppendChild(ft);

            // font 4
            ft = new Font();
            ftn = new FontName { Val = StringValue.FromString("Arial") };
            ftsz = new FontSize { Val = DoubleValue.FromDouble(11) };
            var fontColor = Color.MediumBlue;
            ft.Color = new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = HexBinaryValueFromColor(fontColor) };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            fts.AppendChild(ft);

            fts.Count = UInt32Value.FromUInt32((uint)fts.ChildElements.Count);
            return fts;
        }

        private static HexBinaryValue HexBinaryValueFromColor(Color fillColor)
        {
            return new HexBinaryValue
            {
                Value =
                    ColorTranslator.ToHtml(
                        Color.FromArgb(
                            fillColor.A,
                            fillColor.R,
                            fillColor.G,
                            fillColor.B)).Replace("#", "")
            };
        }
    }
}
