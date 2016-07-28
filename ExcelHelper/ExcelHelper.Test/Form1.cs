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
            DataColumn column1 = new DataColumn("a", typeof(Decimal));
            dt.Columns.Add(column1);
            DataColumn column2 = new DataColumn("b", typeof(string));
            dt.Columns.Add(column2);
            DataColumn column3 = new DataColumn("c", typeof(DateTime));
            dt.Columns.Add(column3);
            DataColumn column4 = new DataColumn("d", typeof(bool));
            dt.Columns.Add(column4);
            DataColumn column5 = new DataColumn("e", typeof(double));
            dt.Columns.Add(column5);
            DataColumn column6 = new DataColumn("f", typeof(Int16));
            dt.Columns.Add(column6);
            DataColumn column7 = new DataColumn("g", typeof(Int32));
            dt.Columns.Add(column7);
            DataColumn column8 = new DataColumn("h", typeof(Int64));
            dt.Columns.Add(column8);
            DataColumn column9 = new DataColumn("i", typeof(int));
            dt.Columns.Add(column9);

            DataRow row1 = dt.NewRow();
            row1[0] = 111111;
            row1[1] = "dfa111111111111111111111111";
            row1[2] = "2016/7/8 12:00:00";
            row1[3] = 0;
            row1[4] = 0.123;
            row1[5] = 1;
            row1[6] = 2;
            row1[7] = 3;
            row1[8] = 4;
            dt.Rows.Add(row1);
            DataRow row2 = dt.NewRow();
            row2[0] = 2.1;
            row2[1] = "dafd1111";
            row2[2] = "2016/7/8 12:00:00";
            row2[3] = 0;
            row2[4] = 0.123;
            dt.Rows.Add(row2);

            dt.TableName = "a";
            ds.Tables.Add(dt);

            DataTable dt1 = new DataTable();
            DataColumn column11 = new DataColumn("a", typeof(int));
            dt1.Columns.Add(column11);
            DataColumn column21 = new DataColumn("b", typeof(string));
            dt1.Columns.Add(column21);

            //DataRow row11 = dt1.NewRow();
            //row11[0] = 1;
            //row11[1] = "12345";
            //dt1.Rows.Add(row11);
            //DataRow row21 = dt1.NewRow();
            //row21[0] = 2;
            //row21[1] = "1111111122222";
            //dt1.Rows.Add(row21);

            dt1.TableName = "b";
            ds.Tables.Add(dt1);

            string fileName = @"D:\test1.xlsx";
            ExcelOperater o = new ExcelOperater();
            //o.DataTableToExcel(dt, fileName);
            o.DataSet2Excel(ds, fileName);
            //DataSet dstest = o.ExcelToDataSet(fileName);

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
            ExcelOperater o = new ExcelOperater();
            DataSet ds = o.ExcelToDataSet("D:\\test1.xlsx");

            //ImportExcel import = new ImportExcel();
            //DataSet ds = import.ExcelToDataSet("D:\\test1.xlsx");
            DataTable dt = ds.Tables[0];
            string s = dt.Rows[1][0].ToString();
            string s1 = DBNull.Value.ToString();
            bool flag = Convert.IsDBNull(dt.Rows[1][0]);
        }
    }
}
