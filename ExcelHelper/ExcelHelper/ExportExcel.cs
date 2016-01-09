using ExcelHelper.Infrastruction;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelHelper
{
    public class ExportExcel : IDisposable//: BaseExcel
    {
        private Excel.ApplicationClass excelApp;
        private Excel.Workbook workBook;
        private Excel.Worksheet workSheet;
        private Excel.Range range;

        public void DataTableToExcel(DataTable sourceTable, string fileName)
        {
            excelApp = new Excel.ApplicationClass();
            if (excelApp == null)
            {
                throw new Exception("打开Excel程序错误！");
            }

            workBook = excelApp.Workbooks.Add(true);
            workSheet = (Excel.Worksheet)workBook.Worksheets[1];
            int rowIndex = 0;          

            //写入列名
            ++rowIndex;
            for (int i = 0; i < sourceTable.Columns.Count; i++)
            {
                workSheet.Cells[rowIndex, i + 1] = sourceTable.Columns[i].ColumnName;
            }
            range = workSheet.get_Range(workSheet.Cells[rowIndex, 1], workSheet.Cells[rowIndex, sourceTable.Columns.Count]);

            FontStyle headerStyle = new FontStyle
            {
                FontSize = 30,
                BordersValue = 1,
                FontBold = true,
                EntireColumnAutoFit = true
            };
            FontStyleHelper.SetFontStyleForRange(range, headerStyle);

            //写入数据
            ++rowIndex;
            for (int r = 0; r < sourceTable.Rows.Count; r++)
            {
                for (int i = 0; i < sourceTable.Columns.Count; i++)
                {
                    workSheet.Cells[rowIndex, i + 1] = ExportHelper.ConvertToCellData(sourceTable, r, i);
                }
                rowIndex++;
            }
            range = workSheet.get_Range(workSheet.Cells[2, 1], workSheet.Cells[sourceTable.Rows.Count + 1, sourceTable.Columns.Count]);
            FontStyle bodyStyle = new FontStyle
            {
                FontSize = 16,
                BordersValue = 1,
                FontAlign = Infrastruction.FontAlign.Right,
                EntireColumnAutoFit = true
            };
            FontStyleHelper.SetFontStyleForRange(range, bodyStyle);

            //只保存一个sheet页
            //workSheet.SaveAs(fileName, Excel.XlFileFormat.xlTemplate, Type.Missing, Type.Missing, Type.Missing,
            //        Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            //保存整个Excel
            workBook.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(false, Type.Missing, Type.Missing);
            excelApp.Quit();

            Dispose();
        }

        public void Dispose()
        {
            Dispose(true);
            //GC.SuppressFinalize(this); //不需要在调用本对象的Finalize()方法
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                //清理托管的代码
                GC.Collect();
            }        

            BaseExcel.Dispose(excelApp, workSheet, workBook, range);
        }
    }
}