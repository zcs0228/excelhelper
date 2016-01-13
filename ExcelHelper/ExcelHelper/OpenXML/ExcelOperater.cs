using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace ExcelHelper.OpenXML
{
    public class ExcelOperater
    {
        /// <summary>
        /// 将DataTable转换为XML
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="fileName"></param>
        public void DataTableToXML(DataTable dataTable, string fileName)
        {
            //指定程序安装目录
            string filePath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + fileName;
            using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write))
            {
                using (XmlWriter xmlWriter = XmlWriter.Create(fs))
                {
                    dataTable.WriteXml(xmlWriter, XmlWriteMode.IgnoreSchema);
                }
            }
            Process.Start(filePath);
        }

        #region 读取Excel
        /// <summary>
        /// 将Excel数据读取到DataSet
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public DataSet ExcelToDataSet(string filePath)
        {
            DataSet dataSet = new DataSet();
            try
            {
                using (SpreadsheetDocument spreadDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    //指定WorkbookPart对象
                    WorkbookPart workBookPart = spreadDocument.WorkbookPart;
                    //获取Excel中SheetName集合
                    List<string> sheetNames = GetSheetNames(workBookPart);

                    foreach (string sheetName in sheetNames)
                    {
                        DataTable dataTable = WorkSheetToTable(workBookPart, sheetName);
                        if (dataTable != null)
                        {
                            dataSet.Tables.Add(dataTable);//将表添加到数据集
                        }
                    }
                }
            }
            catch (Exception exp)
            {
                //throw new Exception("可能Excel正在打开中,请关闭重新操作！");
            }
            return dataSet;
        }

        /// <summary>
        /// 将Excel数据读取到DataTable
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public DataTable ExcelToDataTable(string sheetName, string filePath)
        {
            DataTable dataTable = new DataTable();
            try
            {
                //根据Excel流转换为spreadDocument对象
                using (SpreadsheetDocument spreadDocument = SpreadsheetDocument.Open(filePath, false))//Excel文档包
                {
                    //Workbook workBook = spreadDocument.WorkbookPart.Workbook;//主文档部件的根元素
                    //Sheets sheeets = workBook.Sheets;//块级结构（如工作表、文件版本等）的容器
                    WorkbookPart workBookPart = spreadDocument.WorkbookPart;
                    //获取Excel中SheetName集合
                    List<string> sheetNames = GetSheetNames(workBookPart);

                    if (sheetNames.Contains(sheetName))
                    {
                        //根据WorkSheet转化为Table
                        dataTable = WorkSheetToTable(workBookPart, sheetName);
                    }
                }
            }
            catch (Exception exp)
            {
                //throw new Exception("可能Excel正在打开中,请关闭重新操作！");
            }
            return dataTable;
        }

        /// <summary>
        /// 获取Excel中的sheet页名称
        /// </summary>
        /// <param name="workBookPart"></param>
        /// <returns></returns>
        private List<string> GetSheetNames(WorkbookPart workBookPart)
        {
            List<string> sheetNames = new List<string>();
            Sheets sheets = workBookPart.Workbook.Sheets;
            foreach (Sheet sheet in sheets)
            {
                string sheetName = sheet.Name;
                if (!string.IsNullOrEmpty(sheetName))
                {
                    sheetNames.Add(sheetName);
                }
            }
            return sheetNames;
        }

        /// <summary>
        /// 获取指定sheet名称的Excel数据行集合
        /// </summary>
        /// <param name="workBookPart"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public IEnumerable<Row> GetWorkBookPartRows(WorkbookPart workBookPart, string sheetName)
        {
            IEnumerable<Row> sheetRows = null;
            //根据表名在WorkbookPart中获取Sheet集合
            IEnumerable<Sheet> sheets = workBookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() == 0)
            {
                return null;//没有数据
            }

            WorksheetPart workSheetPart = workBookPart.GetPartById(sheets.First().Id) as WorksheetPart;
            //获取Excel中得到的行
            sheetRows = workSheetPart.Worksheet.Descendants<Row>();

            return sheetRows;
        }

        /// <summary>
        /// 将指定sheet名称的数据转换成DataTable
        /// </summary>
        /// <param name="workBookPart"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private DataTable WorkSheetToTable(WorkbookPart workBookPart, string sheetName)
        {
            //创建Table
            DataTable dataTable = new DataTable(sheetName);

            //根据WorkbookPart和sheetName获取该Sheet下所有行数据
            IEnumerable<Row> sheetRows = GetWorkBookPartRows(workBookPart, sheetName);
            if (sheetRows == null || sheetRows.Count() <= 0)
            {
                return null;
            }

            //将数据导入DataTable,假定第一行为列名,第二行以后为数据
            foreach (Row row in sheetRows)
            {
                //获取Excel中的列头
                if (row.RowIndex == 1)
                {
                    List<DataColumn> listCols = GetDataColumn(row, workBookPart);
                    dataTable.Columns.AddRange(listCols.ToArray());
                }
                else
                {
                    //Excel第二行同时为DataTable的第一行数据
                    DataRow dataRow = GetDataRow(row, dataTable, workBookPart);
                    if (dataRow != null)
                    {
                        dataTable.Rows.Add(dataRow);
                    }
                }
            }
            return dataTable;
        }

        /// <summary>
        /// 获取数字类型格式集合
        /// </summary>
        /// <param name="workBookPart"></param>
        /// <returns></returns>
        private List<string> GetNumberFormatsStyle(WorkbookPart workBookPart)
        {
            List<string> dicStyle = new List<string>();
            Stylesheet styleSheet = workBookPart.WorkbookStylesPart.Stylesheet;
            var test = styleSheet.NumberingFormats;
            if (test == null) return null;
            OpenXmlElementList list = styleSheet.NumberingFormats.ChildElements;//获取NumberingFormats样式集合

            foreach (var element in list)//格式化节点
            {
                if (element.HasAttributes)
                {
                    using (OpenXmlReader reader = OpenXmlReader.Create(element))
                    {
                        if (reader.Read())
                        {
                            if (reader.Attributes.Count > 0)
                            {
                                string numFmtId = reader.Attributes[0].Value;//格式化ID
                                string formatCode = reader.Attributes[1].Value;//格式化Code
                                dicStyle.Add(formatCode);//将格式化Code写入List集合
                            }
                        }
                    }
                }
            }
            return dicStyle;
        }

        /// <summary>
        /// 获得DataColumn
        /// </summary>
        /// <param name="row"></param>
        /// <param name="workBookPart"></param>
        /// <returns></returns>
        private List<DataColumn> GetDataColumn(Row row, WorkbookPart workBookPart)
        {
            List<DataColumn> listCols = new List<DataColumn>();
            foreach (Cell cell in row)
            {
                string cellValue = GetCellValue(cell, workBookPart);
                DataColumn col = new DataColumn(cellValue);
                listCols.Add(col);
            }
            return listCols;
        }

        /// <summary>
        /// 将sheet页中的一行数据转换成DataRow
        /// </summary>
        /// <param name="row"></param>
        /// <param name="dateTable"></param>
        /// <param name="workBookPart"></param>
        /// <returns></returns>
        private DataRow GetDataRow(Row row, DataTable dateTable, WorkbookPart workBookPart)
        {
            //读取Excel中数据,一一读取单元格,若整行为空则忽视该行
            DataRow dataRow = dateTable.NewRow();
            IEnumerable<Cell> cells = row.Elements<Cell>();

            int cellIndex = 0;//单元格索引
            int nullCellCount = cellIndex;//空行索引
            foreach (Cell cell in row)
            {
                string cellVlue = GetCellValue(cell, workBookPart);
                if (string.IsNullOrEmpty(cellVlue))
                {
                    nullCellCount++;
                }

                dataRow[cellIndex] = cellVlue;
                cellIndex++;
            }
            if (nullCellCount == cellIndex)//剔除空行
            {
                dataRow = null;//一行中单元格索引和空行索引一样
            }
            return dataRow;
        }

        /// <summary>
        /// 获得单元格数据值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="workBookPart"></param>
        /// <returns></returns>
        public string GetCellValue(Cell cell, WorkbookPart workBookPart)
        {
            string cellValue = string.Empty;
            if (cell.ChildElements.Count == 0)//Cell节点下没有子节点
            {
                return cellValue;
            }
            string cellRefId = cell.CellReference.InnerText;//获取引用相对位置
            string cellInnerText = cell.CellValue.InnerText;//获取Cell的InnerText
            cellValue = cellInnerText;//指定默认值(其实用来处理Excel中的数字)

            //获取WorkbookPart中NumberingFormats样式集合
            //List<string> dicStyles = GetNumberFormatsStyle(workBookPart);
            //获取WorkbookPart中共享String数据
            SharedStringTable sharedTable = workBookPart.SharedStringTablePart.SharedStringTable;

            try
            {
                EnumValue<CellValues> cellType = cell.DataType;//获取Cell数据类型
                if (cellType != null)//Excel对象数据
                {
                    switch (cellType.Value)
                    {
                        case CellValues.SharedString://字符串
                            //获取该Cell的所在的索引
                            int cellIndex = int.Parse(cellInnerText);
                            cellValue = sharedTable.ChildElements[cellIndex].InnerText;
                            break;
                        case CellValues.Boolean://布尔
                            cellValue = (cellInnerText == "1") ? "TRUE" : "FALSE";
                            break;
                        case CellValues.Date://日期
                            cellValue = Convert.ToDateTime(cellInnerText).ToString();
                            break;
                        case CellValues.Number://数字
                            cellValue = Convert.ToDecimal(cellInnerText).ToString();
                            break;
                        default: cellValue = cellInnerText; break;
                    }
                }
                else//格式化数据
                {
                    #region 根据Excel单元格格式设置数据类型，该部分代码有误，暂未处理
                    /*
                    if (dicStyles.Count > 0 && cell.StyleIndex != null)//对于数字,cell.StyleIndex==null
                    {
                        int styleIndex = Convert.ToInt32(cell.StyleIndex.Value);
                        string cellStyle = dicStyles[styleIndex - 1];//获取该索引的样式
                        if (cellStyle.Contains("yyyy") || cellStyle.Contains("h")
                            || cellStyle.Contains("dd") || cellStyle.Contains("ss"))
                        {
                            //如果为日期或时间进行格式处理,去掉“;@”
                            cellStyle = cellStyle.Replace(";@", "");
                            while (cellStyle.Contains("[") && cellStyle.Contains("]"))
                            {
                                int otherStart = cellStyle.IndexOf('[');
                                int otherEnd = cellStyle.IndexOf("]");

                                cellStyle = cellStyle.Remove(otherStart, otherEnd - otherStart + 1);
                            }
                            double doubleDateTime = double.Parse(cellInnerText);
                            DateTime dateTime = DateTime.FromOADate(doubleDateTime);//将Double日期数字转为日期格式
                            if (cellStyle.Contains("m")) { cellStyle = cellStyle.Replace("m", "M"); }
                            if (cellStyle.Contains("AM/PM")) { cellStyle = cellStyle.Replace("AM/PM", ""); }
                            cellValue = dateTime.ToString(cellStyle);//不知道为什么Excel 2007中格式日期为yyyy/m/d
                        }
                        else//其他的货币、数值
                        {
                            cellStyle = cellStyle.Substring(cellStyle.LastIndexOf('.') - 1).Replace("\\", "");
                            decimal decimalNum = decimal.Parse(cellInnerText);
                            cellValue = decimal.Parse(decimalNum.ToString(cellStyle)).ToString();
                        }
                    }
                    */
                    #endregion
                }
            }
            catch
            {
                //string expMessage = string.Format("Excel中{0}位置数据有误,请确认填写正确！", cellRefId);
                //throw new Exception(expMessage);
                cellValue = "N/A";
            }
            return cellValue;
        }

        /// <summary>
        /// 获得sheet页集合
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private List<string> GetExcelSheetNames(string filePath)
        {
            string sheetName = string.Empty;
            List<string> sheetNames = new List<string>();//所有Sheet表名
            using (SpreadsheetDocument spreadDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workBook = spreadDocument.WorkbookPart;
                Stream stream = workBook.GetStream(FileMode.Open);
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(stream);

                XmlNamespaceManager xmlNSManager = new XmlNamespaceManager(xmlDocument.NameTable);
                xmlNSManager.AddNamespace("default", xmlDocument.DocumentElement.NamespaceURI);
                XmlNodeList nodeList = xmlDocument.SelectNodes("//default:sheets/default:sheet", xmlNSManager);

                foreach (XmlNode node in nodeList)
                {
                    sheetName = node.Attributes["name"].Value;
                    sheetNames.Add(sheetName);
                }
            }
            return sheetNames;
        }

        #region SaveCell
        private void InsertTextCellValue(Worksheet worksheet, string column, uint row, string value)
        {
            Cell cell = ReturnCell(worksheet, column, row);
            CellValue v = new CellValue();
            v.Text = value;
            cell.AppendChild(v);
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
            worksheet.Save();
        }
        private void InsertNumberCellValue(Worksheet worksheet, string column, uint row, string value)
        {
            Cell cell = ReturnCell(worksheet, column, row);
            CellValue v = new CellValue();
            v.Text = value;
            cell.AppendChild(v);
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            worksheet.Save();
        }
        private static Cell ReturnCell(Worksheet worksheet, string columnName, uint row)
        {
            Row targetRow = ReturnRow(worksheet, row);

            if (targetRow == null)
                return null;

            return targetRow.Elements<Cell>().Where(c =>
               string.Compare(c.CellReference.Value, columnName + row,
               true) == 0).First();
        }
        private static Row ReturnRow(Worksheet worksheet, uint row)
        {
            return worksheet.GetFirstChild<SheetData>().
            Elements<Row>().Where(r => r.RowIndex == row).First();
        }
        #endregion

        #endregion


        #region 写入Excel
        /// <summary>
        /// 在指定路径创建SpreadsheetDocument文档
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private SpreadsheetDocument CreateParts(string filePath)
        {
            SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookPart = document.AddWorkbookPart();

            workbookPart.Workbook = new Workbook();

            return document;
        }

        /// <summary>
        /// 创建WorksheetPart
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private WorksheetPart CreateWorksheet(WorkbookPart workbookPart, string sheetName)
        {
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            if (sheets == null)
                sheets = workbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            uint sheetId = 1;

            if (sheets.Elements<Sheet>().Count() > 0)
            {//确定sheet的唯一编号
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }
            if (string.IsNullOrEmpty(sheetName))
            {
                sheetName = "Sheet" + sheetId;
            }

            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);

            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        /// <summary>
        /// 创建sheet样式
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <returns></returns>
        private Stylesheet CreateStylesheet(WorkbookPart workbookPart)
        {
            Stylesheet stylesheet = null;

            if (workbookPart.WorkbookStylesPart != null)
            {
                stylesheet = workbookPart.WorkbookStylesPart.Stylesheet;
                if (stylesheet != null)
                {
                    return stylesheet;
                }
            }
            workbookPart.AddNewPart<WorkbookStylesPart>("Style");
            workbookPart.WorkbookStylesPart.Stylesheet = new Stylesheet();
            stylesheet = workbookPart.WorkbookStylesPart.Stylesheet;

            stylesheet.Fonts = new Fonts()
            {
                Count = (UInt32Value)3U
            };

            //fontId =0,默认样式
            Font fontDefault = new Font(
                                         new FontSize() { Val = 11D },
                                         new FontName() { Val = "Calibri" },
                                         new FontFamily() { Val = 2 },
                                         new FontScheme() { Val = FontSchemeValues.Minor });

            stylesheet.Fonts.Append(fontDefault);

            //fontId =1，标题样式
            Font fontTitle = new Font(new FontSize() { Val = 15D },
                                         new Bold() { Val = true },
                                         new FontName() { Val = "Calibri" },
                                         new FontFamily() { Val = 2 },
                                         new FontScheme() { Val = FontSchemeValues.Minor });
            stylesheet.Fonts.Append(fontTitle);

            //fontId =2，列头样式
            Font fontHeader = new Font(new FontSize() { Val = 13D },
                              new Bold() { Val = true },
                              new FontName() { Val = "Calibri" },
                              new FontFamily() { Val = 2 },
                              new FontScheme() { Val = FontSchemeValues.Minor });
            stylesheet.Fonts.Append(fontHeader);

            //fillId,0总是None，1总是gray125，自定义的从fillid =2开始
            stylesheet.Fills = new Fills()
            {
                Count = (UInt32Value)3U
            };

            //fillid=0
            Fill fillDefault = new Fill(new PatternFill() { PatternType = PatternValues.None });
            stylesheet.Fills.Append(fillDefault);

            //fillid=1
            Fill fillGray = new Fill();
            PatternFill patternFillGray = new PatternFill()
            {
                PatternType = PatternValues.Gray125
            };
            fillGray.Append(patternFillGray);
            stylesheet.Fills.Append(fillGray);

            //fillid=2
            Fill fillYellow = new Fill();
            PatternFill patternFillYellow = new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } })
            {
                PatternType = PatternValues.Solid
            };
            fillYellow.Append(patternFillYellow);
            stylesheet.Fills.Append(fillYellow);

            stylesheet.Borders = new Borders()
            {
                Count = (UInt32Value)2U
            };

            //borderID=0
            Border borderDefault = new Border(new LeftBorder(), new RightBorder(), new TopBorder() { }, new BottomBorder(), new DiagonalBorder());
            stylesheet.Borders.Append(borderDefault);

            //borderID=1
            Border borderContent = new Border(
                new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new DiagonalBorder()
                );
            stylesheet.Borders.Append(borderContent);

            stylesheet.CellFormats = new CellFormats();
            stylesheet.CellFormats.Count = 3;

            //styleIndex =0U
            CellFormat cfDefault = new CellFormat();
            cfDefault.Alignment = new Alignment();
            cfDefault.NumberFormatId = 0;
            cfDefault.FontId = 0;
            cfDefault.BorderId = 0;
            cfDefault.FillId = 0;
            cfDefault.ApplyAlignment = true;
            cfDefault.ApplyBorder = true;
            stylesheet.CellFormats.Append(cfDefault);

            //styleIndex =1U
            CellFormat cfTitle = new CellFormat();
            cfTitle.Alignment = new Alignment();
            cfTitle.NumberFormatId = 0;
            cfTitle.FontId = 1;
            cfTitle.BorderId = 1;
            cfTitle.FillId = 0;
            cfTitle.ApplyBorder = true;
            cfTitle.ApplyAlignment = true;
            cfTitle.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            stylesheet.CellFormats.Append(cfTitle);

            //styleIndex =2U
            CellFormat cfHeader = new CellFormat();
            cfHeader.Alignment = new Alignment();
            cfHeader.NumberFormatId = 0;
            cfHeader.FontId = 2;
            cfHeader.BorderId = 1;
            cfHeader.FillId = 2;
            cfHeader.ApplyAlignment = true;
            cfHeader.ApplyBorder = true;
            cfHeader.ApplyFill = true;
            cfHeader.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            stylesheet.CellFormats.Append(cfHeader);

            //styleIndex =3U
            CellFormat cfContent = new CellFormat();
            cfContent.Alignment = new Alignment();
            cfContent.NumberFormatId = 0;
            cfContent.FontId = 0;
            cfContent.BorderId = 1;
            cfContent.FillId = 0;
            cfContent.ApplyAlignment = true;
            cfContent.ApplyBorder = true;
            stylesheet.CellFormats.Append(cfContent);

            workbookPart.WorkbookStylesPart.Stylesheet.Save();
            return stylesheet;
        }

        /// <summary>
        /// 创建文本单元格,Cell的内容均视为文本
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="cellValue"></param>
        /// <param name="styleIndex"></param>
        /// <returns></returns>
        private Cell CreateTextCell(int columnIndex, int rowIndex, object cellValue, Nullable<uint> styleIndex)
        {
            Cell cell = new Cell();

            cell.DataType = CellValues.InlineString;

            cell.CellReference = GetCellReference(columnIndex) + rowIndex;

            if (styleIndex.HasValue)
                cell.StyleIndex = styleIndex.Value;

            InlineString inlineString = new InlineString();
            Text t = new Text();

            t.Text = cellValue.ToString();
            inlineString.AppendChild(t);
            cell.AppendChild(inlineString);

            return cell;
        }

        /// <summary>
        /// 创建值单元格，Cell会根据单元格值的类型
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <param name="rowIndex"></param>
        /// <param name="cellValue"></param>
        /// <param name="styleIndex"></param>
        /// <returns></returns>
        private Cell CreateValueCell(int columnIndex, int rowIndex, object cellValue, Nullable<uint> styleIndex)
        {
            Cell cell = new Cell();
            cell.CellReference = GetCellReference(columnIndex) + rowIndex;
            CellValue value = new CellValue();
            value.Text = cellValue.ToString();

            //apply the cell style if supplied
            if (styleIndex.HasValue)
                cell.StyleIndex = styleIndex.Value;

            cell.AppendChild(value);

            return cell;
        }

        /// <summary>
        /// 获取行引用，如A1
        /// </summary>
        /// <param name="colIndex"></param>
        /// <returns></returns>
        private string GetCellReference(int colIndex)
        {
            int dividend = colIndex;
            string columnName = String.Empty;
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName =
                    Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }
            return columnName;
        }

        /// <summary>
        /// 创建行数据,不同类型使用不同的styleIndex
        /// </summary>
        /// <param name="dataRow"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        private Row CreateDataRow(DataRow dataRow, int rowIndex)
        {
            Row row = new Row
            {
                RowIndex = (UInt32)rowIndex
            };

            //Nullable<uint> styleIndex = null;
            double doubleValue;
            int intValue;
            DateTime dateValue;
            decimal decValue;

            for (int i = 0; i < dataRow.Table.Columns.Count; i++)
            {
                Cell dataCell;
                if (DateTime.TryParse(dataRow[i].ToString(), out dateValue) && dataRow[i].GetType() == typeof(DateTime))
                {
                    dataCell = CreateTextCell(i + 1, rowIndex, dataRow[i], 3u);
                    //dataCell.DataType = CellValues.Date;
                }
                else if (decimal.TryParse(dataRow[i].ToString(), out decValue) && dataRow[i].GetType() == typeof(decimal))
                {
                    dataCell = CreateValueCell(i + 1, rowIndex, decValue, 3u);
                }
                else if (int.TryParse(dataRow[i].ToString(), out intValue) && dataRow[i].GetType() == typeof(int))
                {
                    dataCell = CreateValueCell(i + 1, rowIndex, intValue, 3u);
                }
                else if (Double.TryParse(dataRow[i].ToString(), out doubleValue) && dataRow[i].GetType() == typeof(double))
                {
                    dataCell = CreateValueCell(i + 1, rowIndex, doubleValue, 3u);
                }
                else
                {
                    dataCell = CreateTextCell(i + 1, rowIndex, dataRow[i], 3u);
                }

                row.AppendChild(dataCell);
                //styleIndex = null;
            }
            return row;
        }

        /// <summary>
        /// 将DataTable的列名称导入Excel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="sheetData"></param>
        private void CreateTableHeader(DataTable dt, SheetData sheetData)
        {
            Row header = new Row
            {
                RowIndex = (UInt32)1
            };
            int colCount = dt.Columns.Count;
            for(int i = 0; i < colCount; i++)
            {
                string colName = dt.Columns[i].ColumnName;
                Cell dataCell = CreateTextCell( i + 1, 1, colName, 3u);
                header.AppendChild(dataCell);
            }
            //Row contentRow = CreateDataRow(header, 1);
            sheetData.AppendChild(header);
        }

        /// <summary>
        /// 将DataTable的数据导入Excel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="sheetData"></param>
        private void InsertDataIntoSheet(DataTable dt, SheetData sheetData)
        {
            //SheetData sheetData = newWorksheetPart.Worksheet.GetFirstChild<SheetData>();

            //CreateTableHeader(dt, sheetData);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                Row contentRow = CreateDataRow(dt.Rows[i], i + 2);
                sheetData.AppendChild(contentRow);
            }
            return;
        }

        /// <summary>
        /// 创建一个SharedStringTablePart(相当于各Sheet共用的存放字符串的容器)
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <returns></returns>
        private SharedStringTablePart CreateSharedStringTablePart(WorkbookPart workbookPart)
        {
            SharedStringTablePart shareStringPart = null;
            if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            }
            return shareStringPart;
        }

        /// <summary>
        /// 导出Excel，执行函数
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="filePath"></param>
        public void DataTableToExcel(DataTable dt, string filePath)
        {
            try
            {
                using (SpreadsheetDocument document = CreateParts(filePath))
                {
                    WorksheetPart worksheetPart = CreateWorksheet(document.WorkbookPart, dt.TableName);

                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                    Stylesheet styleSheet = CreateStylesheet(document.WorkbookPart);

                    //InsertTableTitle(parameter.SheetName, sheetData, styleSheet);

                    // MergeTableTitleCells(dt.Columns.Count, worksheetPart.Worksheet);

                    CreateTableHeader(dt, sheetData);

                    InsertDataIntoSheet(dt, sheetData);

                    SharedStringTablePart sharestringTablePart = CreateSharedStringTablePart(document.WorkbookPart);
                    sharestringTablePart.SharedStringTable = new SharedStringTable();

                    sharestringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text("ExcelReader")));
                    sharestringTablePart.SharedStringTable.Save();
                }
                //result = 0;
            }
            catch (Exception ex)
            {
                //iSession.AddError(ex);
                //result = error_result_prefix - 99;
            }
            //return result;
        }

        #endregion
    }
}