using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel
{
    public class ExcelSheetAttribute : Attribute
    {
        private string name;
        public ExcelSheetAttribute()
        {
        }
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
    }
}
