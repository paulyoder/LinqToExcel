﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.IO;
using LinqToExcel.Domain;
using LinqToExcel.Extensions;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

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

            if (args.ReadOnly)
                connString = connString.Replace("IMEX=1", "IMEX=1;READONLY=TRUE");

            return connString;
        }

        internal static IEnumerable<string> GetWorksheetNames(string fileName)
        {
	        return GetWorksheetNames(fileName, new ExcelQueryArgs());
        }

		internal static IEnumerable<string> GetWorksheetNames(string fileName, ExcelQueryArgs args)
		{
			args.FileName = fileName;
            args.ReadOnly = true;
			return GetWorksheetNames(args);
		} 

		internal static OleDbConnection GetConnection(ExcelQueryArgs args)
		{
			if (args.UsePersistentConnection)
			{
                if (args.PersistentConnection == null)
                    args.PersistentConnection = new OleDbConnection(GetConnectionString(args));

				return args.PersistentConnection;
			}

            return new OleDbConnection(GetConnectionString(args));
		}

        internal static IEnumerable<string> GetWorksheetNames(ExcelQueryArgs args)
        {
            var worksheetNames = new List<string>();

	        var conn = GetConnection(args);
            try
            {
                if (conn.State == ConnectionState.Closed)
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
            finally
            {
                if (!args.UsePersistentConnection)
                    conn.Dispose();
            }
			
            return worksheetNames;
        }

        internal static IEnumerable<string> GetWorksheetNamesOrdered(string fileName)
        {
            if (fileName.ToLower().EndsWith("csv"))
                return new List<string>();

            //open the excel file
            using (FileStream data = new FileStream(fileName, FileMode.Open))
            {
                //unzip
                ZipArchive archive = new ZipArchive(data);

                //select the correct file from the archive
                ZipArchiveEntry appxmlFile = archive.Entries.SingleOrDefault(e => e.FullName == "docProps/app.xml");

                //read the xml
                XDocument xdoc = XDocument.Load(appxmlFile.Open());

                //find the titles element
                XElement titlesElement = xdoc.Descendants().Where(e => e.Name.LocalName == "TitlesOfParts").Single();

                //extract the worksheet names
                return titlesElement
                    .Elements().Where(e => e.Name.LocalName == "vector").Single()
                    .Elements().Where(e => e.Name.LocalName == "lpstr")
                    .Select(e => e.Value);
            }
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
            var conn = GetConnection(args);
            try
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();

                using (var command = conn.CreateCommand())
                {
                    command.CommandText = string.Format("SELECT TOP 1 * FROM [{0}$]", args.WorksheetName);
                    var data = command.ExecuteReader();
                    columns.AddRange(GetColumnNames(data));
                }
            }
            finally
            {
                if (!args.UsePersistentConnection)
                    conn.Dispose();
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

        internal static DatabaseEngine DefaultDatabaseEngine()
        {
            return Is64BitProcess() ? DatabaseEngine.Ace : DatabaseEngine.Jet;
        }

        internal static bool Is64BitProcess()
        {
            return (IntPtr.Size == 8);
        }
    }
}
