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
        public Dictionary<string, Func<string, object>> Transformations { get; private set; }
        public string StartRange { get; set; }
        public string EndRange { get; set; }
        public bool NoHeader { get; set; }

        public ExcelQueryArgs(string fileName, Dictionary<string, string> columnMappings, Dictionary<string, Func<string, object>> transformations)
        {
            FileName = fileName;
            ColumnMappings = columnMappings ?? new Dictionary<string, string>();
            Transformations = transformations ?? new Dictionary<string, Func<string, object>>();
            NoHeader = false;
        }

        public override string ToString()
        {
            var columnMappingsString = new StringBuilder();
            foreach (var kvp in ColumnMappings)
                columnMappingsString.AppendFormat("[{0} = '{1}'] ", kvp.Key, kvp.Value);
            var transformationsString = string.Join(", ", Transformations.Keys.ToArray());
            
            return string.Format("FileName: '{0}'; WorksheetName: '{1}'; WorksheetIndex: {2}; StartRange: {3}; EndRange: {4}; NoHeader: {5}; ColumnMappings: {6}; Transformations: {7}",
                FileName, WorksheetName, WorksheetIndex, StartRange, EndRange, NoHeader, columnMappingsString, transformationsString);
        }
    }
}
