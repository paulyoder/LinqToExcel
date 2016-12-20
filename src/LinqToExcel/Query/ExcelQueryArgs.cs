using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using LinqToExcel.Domain;

namespace LinqToExcel.Query
{
    internal class ExcelQueryArgs
    {
        internal string FileName { get; set; }
        internal DatabaseEngine DatabaseEngine { get; set; }
        internal string WorksheetName { get; set; }
        internal int? WorksheetIndex { get; set; }
        internal Dictionary<string, string> ColumnMappings { get; set; }
        internal Dictionary<string, Func<string, object>> Transformations { get; private set; }
        internal string NamedRangeName { get; set; }
        internal string StartRange { get; set; }
        internal string EndRange { get; set; }
        internal bool NoHeader { get; set; }
        internal StrictMappingType? StrictMapping { get; set; }
        internal bool ReadOnly { get; set; }
        internal bool UsePersistentConnection { get; set; }
		internal OleDbConnection PersistentConnection { get; set; }
        internal TrimSpacesType TrimSpaces { get; set; }

        internal ExcelQueryArgs()
            : this(new ExcelQueryConstructorArgs() { DatabaseEngine = ExcelUtilities.DefaultDatabaseEngine() })
        { }

        internal ExcelQueryArgs(ExcelQueryConstructorArgs args)
        {
            FileName = args.FileName;
            DatabaseEngine = args.DatabaseEngine;
            ColumnMappings = args.ColumnMappings ?? new Dictionary<string, string>();
            Transformations = args.Transformations ?? new Dictionary<string, Func<string, object>>();
            StrictMapping = args.StrictMapping ?? StrictMappingType.None;
            UsePersistentConnection = args.UsePersistentConnection;
            TrimSpaces = args.TrimSpaces;
            ReadOnly = args.ReadOnly;
        }

        public override string ToString()
        {
            var columnMappingsString = new StringBuilder();
            foreach (var kvp in ColumnMappings)
                columnMappingsString.AppendFormat("[{0} = '{1}'] ", kvp.Key, kvp.Value);
            var transformationsString = string.Join(", ", Transformations.Keys.ToArray());

            return string.Format("FileName: '{0}'; WorksheetName: '{1}'; WorksheetIndex: {2}; StartRange: {3}; EndRange: {4}; Named Range: {11}; NoHeader: {5}; ColumnMappings: {6}; Transformations: {7}, StrictMapping: {8}, UsePersistentConnection: {9}, TrimSpaces: {10}",
                FileName, WorksheetName, WorksheetIndex, StartRange, EndRange, NoHeader, columnMappingsString, transformationsString, StrictMapping, UsePersistentConnection, TrimSpaces, NamedRangeName);
        }
    }
}
