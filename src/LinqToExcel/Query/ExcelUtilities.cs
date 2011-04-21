using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.IO;
using LinqToExcel.Extensions;

namespace LinqToExcel.Query
{
    internal static class ExcelUtilities
    {
        internal static string GetConnectionString(string fileName)
        {
            return GetConnectionString(fileName, false);
        }

        internal static string GetConnectionString(string fileName, bool noHeader)
        {
            var connString = "";
            var fileNameLower = fileName.ToLower();

            if (fileNameLower.EndsWith("xlsx") ||
                fileNameLower.EndsWith("xlsm"))
                connString = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""",
                    fileName);
            else if (fileNameLower.EndsWith("xlsb"))
            {
                connString = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                    fileName);
            }
            else if (fileNameLower.EndsWith("csv"))
            {
                connString = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""text;HDR=YES;FMT=Delimited;IMEX=1""",
                    Path.GetDirectoryName(fileName));
            }
            else
                connString = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1""",
                    fileName);

            if (noHeader)
                connString = connString.Replace("HDR=YES", "HDR=NO");

            return connString;
        }

        internal static IEnumerable<string> GetWorksheetNames(string fileName)
        {
            var worksheetNames = new List<string>();
            using (var conn = new OleDbConnection(GetConnectionString(fileName)))
            {
                conn.Open();
                var excelTables = conn.GetOleDbSchemaTable(
                    OleDbSchemaGuid.Tables,
                    new Object[] { null, null, null, "TABLE" });

                worksheetNames.AddRange(
                    from DataRow row in excelTables.Rows
                    let tableName = row["TABLE_NAME"].ToString()
                        .Replace("$", "")
                        .RegexReplace("(^'|'$)", "")
                        .Replace("''", "'")
                    where IsNotBuiltinTable(tableName)
                    select tableName);

                excelTables.Dispose();
            }
            return worksheetNames;
        }

        internal static bool IsNotBuiltinTable(string tableName)
        {
            return !tableName.Contains("FilterDatabase") && !tableName.Contains("Print_Area");
        }

        internal static IEnumerable<string> GetColumnNames(string worksheetName, string fileName)
        {
            var columns = new List<string>();
            using (var conn = new OleDbConnection(GetConnectionString(fileName)))
            using (var command = conn.CreateCommand())
            {
                conn.Open();
                command.CommandText = string.Format("SELECT TOP 1 * FROM [{0}$]", worksheetName);
                var data = command.ExecuteReader();
                columns.AddRange(GetColumnNames(data));
            }
            return columns;
        }

        internal static IEnumerable<string> GetColumnNames(IDataReader data)
        {
            var columns = new List<string>();
            var sheetSchema = data.GetSchemaTable();
            foreach (DataRow row in sheetSchema.Rows)
                columns.Add(row["ColumnName"].ToString());

            return columns;
        }
    }
}
