using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelHelper.OpenXML
{
    public class FormulaCell : Cell
    {
        public FormulaCell(string header, string text, int index)
        {
            CellFormula = new CellFormula { CalculateCell = true, Text = text };
            DataType = CellValues.Number;
            CellReference = header + index;
        }
    }
}
