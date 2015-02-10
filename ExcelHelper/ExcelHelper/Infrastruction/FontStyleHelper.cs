using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelHelper.Infrastruction
{
    public class FontStyleHelper
    {
        /// <summary>
        /// 对选中区域设置格式
        /// </summary>
        /// <param name="range">选中区域</param>
        /// <param name="fontStyle">样式表</param>
        public static void SetFontStyleForRange(Excel.Range range, FontStyle fontStyle)
        {
            if (fontStyle.FontSize != 0)
            {
                range.Font.Size = fontStyle.FontSize;
            }
            if (fontStyle.FontName != null)
            {
                range.Font.Name = fontStyle.FontName;
            }
            if (fontStyle.FontBold != false)
            {
                range.Font.Bold = fontStyle.FontBold;
            }
            if (fontStyle.FontAlign == FontAlign.Center)
            {
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            else if (fontStyle.FontAlign == FontAlign.Left)
            {
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            }
            else if (fontStyle.FontAlign == FontAlign.Right)
            {
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            }
            if (fontStyle.BordersValue != 0)
            {
                range.Borders.Value = fontStyle.BordersValue;
            }
            if (fontStyle.FontColorIndex != 0)
            {
                range.Font.ColorIndex = fontStyle.FontColorIndex;
            }
            if (fontStyle.InteriorColorIndex != 0)
            {
                range.Interior.ColorIndex = fontStyle.InteriorColorIndex;
            }
            if (fontStyle.EntireColumnAutoFit == true)
            {
                range.EntireColumn.AutoFit();
            }
        }
    }
}
