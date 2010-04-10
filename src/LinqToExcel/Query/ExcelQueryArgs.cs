using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel.Query
{
    public class ExcelQueryArgs
    {
        public string FileName { get; private set; }
        public string WorksheetName { get; set; }
        public int? WorksheetIndex { get; set; }
        public Dictionary<string, string> ColumnMappings { get; set; }
        public string StartRange { get; set; }
        public string EndRange { get; set; }

        public ExcelQueryArgs(string fileName, Dictionary<string, string> columnMappings)
        {
            FileName = fileName;
            ColumnMappings = columnMappings ?? new Dictionary<string, string>();
        }

        public override string ToString()
        {
            var columnMappingsString = new StringBuilder();
            foreach (var kvp in ColumnMappings)
                columnMappingsString.AppendFormat("[{0} = '{1}'] ", kvp.Key, kvp.Value);
            
            return string.Format("FileName: '{0}'; WorksheetName: '{1}'; WorksheetIndex: {2}; StartRange: {3}; EndRange {4}; ColumnMappings: {5}",
                FileName, WorksheetName, WorksheetIndex, StartRange, EndRange, columnMappingsString);
        }
    }
}
