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
        public static void Dispose(Excel.ApplicationClass excelApp)
        {
            if (excelApp != null)
            {
                excelApp.Quit();
                Kill(excelApp);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
            
            int generation = GC.GetGeneration(excelApp);
            System.GC.Collect(generation);
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
            IntPtr t = new IntPtr(excel.Hwnd);   //得到这个句柄，具体作用是得到这块内存入口 

            int k = 0;
            GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
            p.Kill();     //关闭进程k
        }
    }
}
