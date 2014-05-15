using System;
using System.Collections.Generic;
using System.Data.OleDb;
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
        internal StrictMappingType? StrictMapping { get; set; }
		internal bool UsePersistentConnection { get; set; }
        internal TrimSpacesType TrimSpaces { get; set; }
        internal bool ReadOnly { get; set; }
    }
}
