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
        }
    }
}