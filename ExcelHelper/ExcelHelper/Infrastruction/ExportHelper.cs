using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ExcelHelper.Infrastruction
{
    public class ExportHelper
    {
        public static string ConvertToCellData(DataTable sourceTable, int rowIndex, int colIndex)
        {
            DataColumn col = sourceTable.Columns[colIndex];
            object data = sourceTable.Rows[rowIndex][colIndex];
            if (col.DataType == System.Type.GetType("System.DateTime"))
            {
                if (data.ToString().Trim() != "")
                {
                    return Convert.ToDateTime(data).ToString("yyyy-MM-dd HH:mm:ss");
                }
                else
                {
                    return (Convert.ToDateTime(DateTime.Now)).ToString("yyyy-MM-dd HH:mm:ss");
                }
            }
            else if (col.DataType == System.Type.GetType("System.String"))
            {
                return "'" + data.ToString().Trim();
            }
            else
            {
                return data.ToString().Trim();
            }
        }
    }
}
