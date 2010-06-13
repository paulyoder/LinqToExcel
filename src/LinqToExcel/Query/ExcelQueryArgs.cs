using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel.Query
{
    internal class ExcelQueryArgs
    {
        internal string FileName { get; private set; }
        internal string WorksheetName { get; set; }
        internal int? WorksheetIndex { get; set; }
        internal Dictionary<string, string> ColumnMappings { get; set; }
        internal Dictionary<string, Func<string, object>> Transformations { get; private set; }
        internal string StartRange { get; set; }
        internal string EndRange { get; set; }
        internal bool NoHeader { get; set; }
        internal bool StrictMapping { get; set; }

        internal ExcelQueryArgs(ExcelQueryConstructorArgs Args)
        {
            FileName = Args.FileName;
            ColumnMappings = Args.ColumnMappings ?? new Dictionary<string, string>();
            Transformations = Args.Transformations ?? new Dictionary<string, Func<string, object>>();
            StrictMapping = Args.StrictMapping;
        }

        public override string ToString()
        {
            var columnMappingsString = new StringBuilder();
            foreach (var kvp in ColumnMappings)
                columnMappingsString.AppendFormat("[{0} = '{1}'] ", kvp.Key, kvp.Value);
            var transformationsString = string.Join(", ", Transformations.Keys.ToArray());
            
            return string.Format("FileName: '{0}'; WorksheetName: '{1}'; WorksheetIndex: {2}; StartRange: {3}; EndRange: {4}; NoHeader: {5}; ColumnMappings: {6}; Transformations: {7}, StrictMapping: {8}",
                FileName, WorksheetName, WorksheetIndex, StartRange, EndRange, NoHeader, columnMappingsString, transformationsString, StrictMapping);
        }
    }
}
