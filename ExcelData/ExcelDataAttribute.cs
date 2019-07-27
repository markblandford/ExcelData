using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelData
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelDataSheetNameAttribute : Attribute
    {
        private string _shtName;

        public string SheetName { get { return _shtName; } }

        public ExcelDataSheetNameAttribute(string sheetName)
        {
            _shtName = sheetName;  
        }

        public override string ToString()
        {
 	         return SheetName;
        }
    }

    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelDataColumnAttribute : Attribute
    {
        private string _column;
        public string Column { get { return _column; } }

        public ExcelDataColumnAttribute(string column)
        {
            _column = column;
        }

        public override string ToString()
        {
 	         return Column;
        }
    }

}
