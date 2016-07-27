using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelHelper.OpenXML
{
    public class NumberCell : Cell
    {
        public NumberCell(string header, string text, int index)
        {
            DataType = CellValues.Number;
            CellReference = header + index;
            CellValue = new CellValue(text);
        }
    }
}
