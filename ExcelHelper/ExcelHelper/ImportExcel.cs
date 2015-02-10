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
    public class ImportExcel : IDisposable
    {
        private Excel.ApplicationClass excelApp;
        private Excel.Workbook workBook;
        private Excel.Worksheet workSheet;
        public DataSet ExcelToDataSet(string fileName)
        {
            try
            {
                excelApp = new Excel.ApplicationClass();
                workBook = excelApp.Workbooks.Open(fileName, 0,
                    false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, 1, 0);
                workSheet = (Excel.Worksheet)workBook.Worksheets[1];
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Source + ":" + ex.Message);
            }

            int sheetNumber = workBook.Worksheets.Count;
            string[] sheetName = new string[sheetNumber];
            for (int i = 0; i < sheetNumber; i++)
            {
                sheetName[i] = ((Excel.Worksheet)workBook.Worksheets[i + 1]).Name;
            }
            

            DataSet ds = null;
            List<string> connStrs = new List<string>();
            connStrs.Add("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + fileName 
                + ";Extended Properties=\"Excel 8.0;HDR=No;IMEX=1;\"");
            connStrs.Add("Provider = Microsoft.ACE.OLEDB.12.0 ; Data Source = " + fileName 
                + ";Extended Properties=\"Excel 12.0;HDR=No;IMEX=1;\"");
            foreach (string item in connStrs)
            {
                ds = ImportHelper.GetDataSetFromExcel(item, sheetName);
                if (ds != null)
                    break;
            }

            Dispose();
            ds = ImportHelper.ConvertDataSet(ds);
            return ds;
        }

        ~ImportExcel()
        {
            Dispose(false);
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
            }
            //清理非托管的代码
            if (workSheet != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);
                workSheet = null;
            }
            if (workBook != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workBook);
                workBook = null;
            }
            BaseExcel.Dispose(excelApp);
        }
    }
}
