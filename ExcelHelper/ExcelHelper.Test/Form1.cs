using ExcelHelper.Infrastruction;
using ExcelHelper.OpenXML;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelHelper.Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            DataColumn column1 = new DataColumn("a", typeof(int));
            dt.Columns.Add(column1);
            DataColumn column2 = new DataColumn("b", typeof(string));
            dt.Columns.Add(column2);
            DataRow row1 = dt.NewRow();
            row1[0] = 1;
            row1[1] = "dfa1111";
            dt.Rows.Add(row1);
            DataRow row2 = dt.NewRow();
            row2[0] = 2;
            row2[1] = "dafd1111";
            dt.Rows.Add(row2);
            dt.TableName = "a";

            string fileName = @"C:\Users\Administrator\Desktop\excel\c.xlsx";
            ExcelOperater o = new ExcelOperater();
            o.DataTableToExcel(dt, fileName);
            //o.ExcelToDataSet(fileName);

            //ExportExcel export = new ExportExcel();
            //export.DataTableToExcel(dt, "D:\\test1.xls");

            /*
            string fileName = @"C:\Users\Administrator\Desktop\excel\a.xlsx";
            //Zip.GetPart(fileName);
            ReadExcelByXML excel = new ReadExcelByXML(fileName);
            excel.Sheets();
            */
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ImportExcel import = new ImportExcel();
            DataSet ds = import.ExcelToDataSet("D:\\test1.xls");
            DataTable dt = ds.Tables[0];
            string s = dt.Rows[1][0].ToString();
            string s1 = DBNull.Value.ToString();
            bool flag = Convert.IsDBNull(dt.Rows[1][0]);
        }
    }
}
