using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.IO;

namespace LinqToExcel.Query
{
    internal static class ExcelUtilities
    {
        internal static string GetConnectionString(string FileName)
        {
            return GetConnectionString(FileName, false);
        }

        internal static string GetConnectionString(string FileName, bool NoHeader)
        {
            var connString = "";
            var fileNameLower = FileName.ToLower();

            if (fileNameLower.EndsWith("xlsx") ||
                fileNameLower.EndsWith("xlsm"))
                connString = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""",
                    FileName);
            else if (fileNameLower.EndsWith("xlsb"))
            {
                connString = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                    FileName);
            }
            else if (fileNameLower.EndsWith("csv"))
            {
                connString = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""text;HDR=YES;FMT=Delimited;IMEX=1""",
                    Path.GetDirectoryName(FileName));
            }
            else
                connString = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1""",
                    FileName);

            if (NoHeader)
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

                foreach (DataRow row in excelTables.Rows)
                    worksheetNames.Add(row["TABLE_NAME"].ToString()
                                                        .Replace("$", "")
                                                        .Replace("'", ""));

                excelTables.Dispose();
            }
            return worksheetNames;
        }

        internal static IEnumerable<string> GetColumnNames(string WorksheetName, string FileName)
        {
            var columns = new List<string>();
            using (var conn = new OleDbConnection(GetConnectionString(FileName)))
            using (var command = conn.CreateCommand())
            {
                conn.Open();
                command.CommandText = string.Format("SELECT TOP 1 * FROM [{0}$]", WorksheetName);
                var data = command.ExecuteReader();
                columns.AddRange(GetColumnNames(WorksheetName, data));
            }
            return columns;
        }

        internal static IEnumerable<string> GetColumnNames(string WorksheetName, IDataReader data)
        {
            var columns = new List<string>();
            var sheetSchema = data.GetSchemaTable();
            foreach (DataRow row in sheetSchema.Rows)
                columns.Add(row["ColumnName"].ToString());

            return columns;
        }
    }
}
