using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelHelper.OpenXML
{
    /// <summary>
    /// 转换为"yyyy-MM-dd"
    /// </summary>
    public class DateCell : Cell
    {
        public DateCell(string header, DateTime dateTime, int index)
        {
            DataType = CellValues.Date;
            CellReference = header + index;
            StyleIndex = (UInt32)CustomStylesheet.CustomCellFormats.DefaultDate;
            CellValue = new CellValue(dateTime.ToString("yyyy-MM-dd"));
        }
    }
}
