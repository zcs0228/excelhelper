using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelHelper.Infrastruction
{
    public class BaseExcel
    {
        /// <summary>
        /// 释放Excel资源
        /// </summary>
        /// <param name="excelApp"></param>
        public static void Dispose(Excel.ApplicationClass excelApp, Excel.Worksheet workSheet, Excel.Workbook workBook, Excel.Range range)
        {
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
            if (excelApp != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                excelApp = null;
            }
            if (range != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                range = null;
            }
            KillProcess();
        }
        /// <summary>
        /// 关闭进程
        /// </summary>
        /// <param name="hwnd"></param>
        /// <param name="ID"></param>
        /// <returns></returns>
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        private static void Kill(Excel.Application excel)
        {
            int id = 0;
            IntPtr intptr = new IntPtr(excel.Hwnd);    //得到这个句柄，具体作用是得到这块内存入口
            System.Diagnostics.Process p = null;
            try
            {
                GetWindowThreadProcessId(intptr, out id);  //得到本进程唯一标志
                p = System.Diagnostics.Process.GetProcessById(id);  //得到对进程k的引用
                if (p != null)
                {
                    p.Kill();  //关闭进程k
                    p.Dispose();
                }
            }
            catch
            {
            }
        }
        //强制结束进程
        private static void KillProcess()
        {
            System.Diagnostics.Process[] allProcess = System.Diagnostics.Process.GetProcesses();
            foreach (System.Diagnostics.Process thisprocess in allProcess)
            {
                string processName = thisprocess.ProcessName;
                if (processName.ToLower() == "excel")
                {
                    try
                    {
                        thisprocess.Kill();
                    }
                    catch
                    {
                    }
                }
            }
        }
    }
}