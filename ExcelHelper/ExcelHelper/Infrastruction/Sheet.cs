using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelHelper.Infrastruction
{
    public struct Sheet
    {
        string _name;
        string _sheetId;
        string _rid;

        public Sheet(string name, string sheetId, string rid)
        {
            _name = name;
            _sheetId = sheetId;
            _rid = rid;
        }
    }
}
