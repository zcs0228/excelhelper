using ExcelHelper.Infrastruction;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelHelper
{
    public class ImportExcel : IDisposable
    {
        private Excel.ApplicationClass excelApp;
        private Excel.Workbook workBook;
        private Excel.Worksheet workSheet;
        private Excel.Range range;

        public DataSet ExcelToDataSet(string fileName)
        {
            if (!File.Exists(fileName))
            {
                return null;
            }
            FileInfo file = new FileInfo(fileName);
            string strConnection = string.Empty;
            string extension = file.Extension;
            string vsSql = string.Empty;
            switch (extension)
            {
                case ".xls":
                    strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                    break;
                case ".xlsx":
                    strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
                    break;
                default:
                    strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
                    break;
            }
            DataSet ds = ImportHelper.GetDataSetFromExcel(strConnection);

            Dispose();
            ds = ImportHelper.ConvertDataSet(ds);
            return ds;
        }

        public DataSet ExcelToDataSetByDcom(string fileName)
        {
            DataSet result = null;
            excelApp = new Excel.ApplicationClass();
            if (excelApp == null)
            {
                throw new Exception("打开Excel程序错误！");
            }

            excelApp.Visible = false; excelApp.UserControl = true;
            // 以只读的形式打开EXCEL文件
            workBook = excelApp.Application.Workbooks.Open(fileName, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            int sheets = workBook.Worksheets.Count;
            if (sheets >= 1)
            {
                result = new DataSet();
            }
            for(int i = 1; i <= sheets; i++)
            {
                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(i);
                string sheetName = workSheet.Name;

                DataTable dt = new DataTable();
                dt.TableName = sheetName;

                //取得总记录行数
                int rows = workSheet.UsedRange.Cells.Rows.Count; //得到行数
                int columns = workSheet.UsedRange.Cells.Columns.Count;//得到列数
                if (rows == 0 || columns == 0) return null;
                //取得数据范围区域
                range = workSheet.get_Range(workSheet.Cells[1, 1], workSheet.Cells[rows, columns]);
                object[,] arryItem = (object[,])range.Value2; //get range's value

                //生成DataTable的列
                for(int col = 1; col <= columns; col++)
                {
                    string dcName = arryItem[1, col].ToString().Trim();
                    DataColumn dc = new DataColumn(dcName, typeof(string));
                    dt.Columns.Add(dc);
                }
                //将数据填充到DataTable
                for(int row = 2; row <= rows; row++)
                {
                    object[] rowvalue = new object[columns];
                    for (int col = 1; col <= columns; col++)
                    {
                        rowvalue[col - 1] = arryItem[row, col];
                    }
                    dt.Rows.Add(rowvalue);
                }
                //将DataTable填充到DataSet
                result.Tables.Add(dt);
            }

            //清理非托管对象
            workBook.Close(false, Type.Missing, Type.Missing);
            excelApp.Quit();
            Dispose();
            return result;
        }

        public void Dispose()
        {
            GC.Collect();
            BaseExcel.Dispose(excelApp, workSheet, workBook, range);
        }
    }
}