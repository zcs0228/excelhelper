using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelHelper.OpenXML
{
    public class TextCell : Cell
    {
        public TextCell(string header, string text, int index)
        {
            DataType = CellValues.InlineString;
            CellReference = header + index;
            InlineString = new InlineString { Text = new Text { Text = text } };
        }
    }
}
