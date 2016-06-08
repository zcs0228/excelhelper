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
            DataSet ds = new DataSet();
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
            ds.Tables.Add(dt);

            DataTable dt1 = new DataTable();
            DataColumn column11 = new DataColumn("a", typeof(int));
            dt1.Columns.Add(column11);
            DataColumn column21 = new DataColumn("b", typeof(string));
            dt1.Columns.Add(column21);
            DataRow row11 = dt1.NewRow();
            row11[0] = 1;
            row11[1] = "12345";
            dt1.Rows.Add(row11);
            DataRow row21 = dt1.NewRow();
            row21[0] = 2;
            row21[1] = "1111111122222";
            dt1.Rows.Add(row21);
            dt1.TableName = "b";
            ds.Tables.Add(dt1);
            
            string fileName = @"D:\test1.xlsx";
            ExcelOperater o = new ExcelOperater();
            //o.DataTableToExcel(dt, fileName);
            //o.DataSetToExcel(ds, fileName);
            DataSet dstest = o.ExcelToDataSet(fileName);

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
            DataSet ds = import.ExcelToDataSet("D:\\test1.xlsx");
            DataTable dt = ds.Tables[0];
            string s = dt.Rows[1][0].ToString();
            string s1 = DBNull.Value.ToString();
            bool flag = Convert.IsDBNull(dt.Rows[1][0]);
        }
    }
}
