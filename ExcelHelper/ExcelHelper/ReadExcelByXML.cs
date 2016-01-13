using ExcelHelper.Infrastruction;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace ExcelHelper
{
    public class ReadExcelByXML
    {
        private string _fileName = String.Empty;

        public ReadExcelByXML(string fileName)
        {
            _fileName = fileName;
        }

        public List<Sheet> Sheets()
        {
            List<Sheet> result = new List<Sheet>();
            using (Stream zs = Zip.GetPartStream(_fileName, "/xl/workbook.xml"))
            {
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(zs);
                XmlNodeList elms = xmlDocument.DocumentElement["sheets"].ChildNodes;
                for (int i = 0; i < elms.Count; i++)
                {
                    XmlAttributeCollection attrs = elms[i].Attributes;
                    result.Add(new Sheet(attrs["name"].Value, attrs["sheetId"].Value, attrs["r:id"].Value));
                }
            }
            return result;
        }
    }
}
