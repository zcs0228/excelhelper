using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper.Infrastruction
{
    public class FontStyle
    {
        /// <summary>
        /// 字体大小
        /// </summary>
        public int FontSize { get; set; }
        /// <summary>
        /// 字体名称
        /// </summary>
        public string FontName { get; set; }
        /// <summary>
        /// 是否为粗体
        /// </summary>
        public bool FontBold { get; set; }
        /// <summary>
        /// 字体对齐方式
        /// </summary>
        public FontAlign FontAlign { get; set; }
        /// <summary>
        /// 边框样式
        /// </summary>
        public int BordersValue { get; set; }
        /// <summary>
        /// 字体颜色索引
        /// </summary>
        public int FontColorIndex { get; set; }
        /// <summary>
        /// 背景颜色索引
        /// </summary>
        public int InteriorColorIndex { get; set; }
        /// <summary>
        /// 列宽自适应
        /// </summary>
        public bool EntireColumnAutoFit { get; set; }
    }

    public enum FontAlign
    {
        Center,
        Right,
        Left
    }
}
