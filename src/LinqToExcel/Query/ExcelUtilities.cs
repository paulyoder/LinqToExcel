using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.IO;
using LinqToExcel.Domain;
using LinqToExcel.Extensions;

namespace LinqToExcel.Query
{
    internal static class ExcelUtilities
    {
        internal static string GetConnectionString(ExcelQueryArgs args)
        {
            var connString = "";
            var fileNameLower = args.FileName.ToLower();

            if (fileNameLower.EndsWith("xlsx") ||
                fileNameLower.EndsWith("xlsm"))
            {
                connString = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""",
                    args.FileName);
            }
            else if (fileNameLower.EndsWith("xlsb"))
            {
                connString = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                    args.FileName);
            }
            else if (fileNameLower.EndsWith("csv"))
            {
                if (args.DatabaseEngine == DatabaseEngine.Jet)
                {
                    connString = string.Format(
                        @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""text;HDR=YES;FMT=Delimited;IMEX=1""",
                        Path.GetDirectoryName(args.FileName));
                }
                else
                {
                    connString = string.Format(
                        @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""text;Excel 12.0;HDR=YES;IMEX=1""",
                        Path.GetDirectoryName(args.FileName));
                }
            }
            else
            {
                if (args.DatabaseEngine == DatabaseEngine.Jet)
                {
                    connString = string.Format(
                        @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1""",
                        args.FileName);
                }
                else
                {
                    connString = string.Format(
                        @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                        args.FileName);
                }
            }

            if (args.NoHeader)
                connString = connString.Replace("HDR=YES", "HDR=NO");

            return connString;
        }

        internal static IEnumerable<string> GetWorksheetNames(string fileName)
        {
            var args = new ExcelQueryArgs();
            args.FileName = fileName;
            return GetWorksheetNames(args);
        }

        internal static IEnumerable<string> GetWorksheetNames(ExcelQueryArgs args)
        {
            var worksheetNames = new List<string>();
            using (var conn = new OleDbConnection(GetConnectionString(args)))
            {
                conn.Open();
                var excelTables = conn.GetOleDbSchemaTable(
                    OleDbSchemaGuid.Tables,
                    new Object[] { null, null, null, "TABLE" });

                worksheetNames.AddRange(
                    from DataRow row in excelTables.Rows
                    where IsTable(row)
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

        internal static bool IsTable(DataRow row)
        {
            return row["TABLE_NAME"].ToString().Contains("$");
        }

        internal static bool IsNotBuiltinTable(string tableName)
        {
            return !tableName.Contains("FilterDatabase") && !tableName.Contains("Print_Area");
        }

        internal static IEnumerable<string> GetColumnNames(string worksheetName, string fileName)
        {
            var args = new ExcelQueryArgs();
            args.WorksheetName = worksheetName;
            args.FileName = fileName;
            return GetColumnNames(args);
        }

        internal static IEnumerable<string> GetColumnNames(ExcelQueryArgs args)
        {
            var columns = new List<string>();
            using (var conn = new OleDbConnection(GetConnectionString(args)))
            using (var command = conn.CreateCommand())
            {
                conn.Open();
                command.CommandText = string.Format("SELECT TOP 1 * FROM [{0}$]", args.WorksheetName);
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
