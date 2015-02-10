using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper.Infrastruction
{
    public class ImportHelper
    {
        /// <summary>
        /// 通过OleDb获得DataSet
        /// </summary>
        /// <param name="connStr"></param>
        /// <param name="sheetNames"></param>
        /// <returns></returns>
        public static DataSet GetDataSetFromExcel(string connStr, string[] sheetNames)
        {
            DataSet ds = null;
            using (OleDbConnection conn = new OleDbConnection(connStr))
            {
                try
                {
                    conn.Open();
                    OleDbDataAdapter da;
                    ds = new DataSet();
                    foreach (string item in sheetNames)
                    {
                        string sql = "select * from [" + item + "$] ";
                        da = new OleDbDataAdapter(sql, conn);
                        da.Fill(ds, item);
                        da.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    conn.Close();
                    throw new Exception(ex.Source + ":" + ex.Message);
                }
            }
            return ds;
        }

        public static DataSet ConvertDataSet(DataSet source)
        {
            DataSet result = new DataSet();
            int dataTableCount = source.Tables.Count;
            DataTable temp = null;
            for (int i = 0; i < dataTableCount; i++)
            {
                temp = ConvertDataTable(source.Tables[i]);
                result.Tables.Add(temp);
                result.Tables[i].TableName = source.Tables[i].TableName;
            }
            return result;
        }

        private static DataTable ConvertDataTable(DataTable source)
        {
            DataTable result = new DataTable();
            int columnsCount = source.Columns.Count;
            int rowsCount = source.Rows.Count;
            for (int i = 0; i < columnsCount; i++)
            {
                DataColumn column = new DataColumn(source.Rows[0][i].ToString().Trim());
                result.Columns.Add(column);
            }
            DataRow dr;
            for (int r = 1; r < rowsCount; r++)
            {
                dr = result.NewRow();
                for (int c = 0; c < columnsCount; c++)
                {
                    dr[c] = source.Rows[r][c].ToString().Trim();
                }
                result.Rows.Add(dr);
            }
            return result;
        }
    }
}
