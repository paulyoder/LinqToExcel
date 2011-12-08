using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LinqToExcel.Domain;

namespace LinqToExcel.Query
{
    internal class ExcelQueryConstructorArgs
    {
        internal string FileName { get; set; }
        internal DatabaseEngine DatabaseEngine { get; set; }
        internal Dictionary<string, string> ColumnMappings { get; set; }
        internal Dictionary<string, Func<string, object>> Transformations { get; set; }
        internal bool StrictMapping { get; set; }
    }
}
