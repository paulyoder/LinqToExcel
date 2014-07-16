using System;
using System.Data.OleDb;
using System.Linq.Expressions;
using LinqToExcel.Query;
using System.Collections.Generic;
using LinqToExcel.Domain;

namespace LinqToExcel
{
    public interface IExcelQueryFactory : IDisposable
    {
        /// <summary>
        /// Full path to the Excel spreadsheet
        /// </summary>
        string FileName { get; set; }

        /// <summary>
        /// Confirms all the worksheet columns are mapped to a property, and if not, throws a StrictMappingException
        /// </summary>
        StrictMappingType? StrictMapping { get; set; }

        /// <summary>
        /// Indicates how to treat leading and trailing spaces in string values.
        /// </summary>
        TrimSpacesType TrimSpaces { get; set; }

        /// <summary>
        /// Sets the database engine to use (spreadsheets ending in xlsx, xlsm, xlsb will always use the Ace engine)
        /// </summary>
        DatabaseEngine DatabaseEngine { get; set; }

        /// <summary>
        /// If true, uses a single, persistent connection for the lifetime of the factory.
        /// If false, a new connection is created for each query
        /// Default is false
        /// </summary>
        bool UsePersistentConnection { get; set; }

        /// <summary>
        /// Add a column to property mapping
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="property">Class property to map to</param>
        /// <param name="column">Worksheet column name to map from</param>
        void AddMapping<TSheetData>(Expression<Func<TSheetData, object>> property, string column);

        /// <summary>
        /// Add a column to property mapping
        /// </summary>
        /// <param name="propertyName">Class property to map to</param>
        /// <param name="column">Worksheet column name to map from</param>
        void AddMapping(string propertyName, string column);

        /// <summary>
        /// Add a column to property mapping with a transformation operation
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="property">Class property to map to</param>
        /// <param name="column">Worksheet column name to map from</param>
        /// <param name="transformation">Lambda expression that transforms a cell value in the spreadsheet to the desired property value</param>
        void AddMapping<TSheetData>(Expression<Func<TSheetData, object>> property, string column, Func<string, object> transformation);

        /// <summary>
        /// Transforms a cell value in the spreadsheet to the desired property value
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="property">Class property value to transform</param>
        /// <param name="transformation">Lambda expression that transforms a cell value in the spreadsheet to the desired property value</param>
        /// <example>
        /// AddTransformation{Person}(p => p.IsActive, x => x == "Y");
        /// AddTransformation{Person}(p => p.IsYoung, x => DateTime.Parse(x) > new DateTime(2000, 1, 1));
        /// </example>
        void AddTransformation<TSheetData>(Expression<Func<TSheetData, object>> property, Func<string, object> transformation);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        ExcelQueryable<TSheetData> Worksheet<TSheetData>();

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetName">Name of the worksheet</param>
        ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        ExcelQueryable<TSheetData> Worksheet<TSheetData>(int worksheetIndex);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        ExcelQueryable<Row> Worksheet();

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        ExcelQueryable<Row> Worksheet(string worksheetName);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        ExcelQueryable<Row> Worksheet(int worksheetIndex);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string startRange, string endRange);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string startRange, string endRange, string worksheetName);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string startRange, string endRange, int worksheetIndex);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        ExcelQueryable<Row> WorksheetRange(string startRange, string endRange);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        ExcelQueryable<Row> WorksheetRange(string startRange, string endRange, string worksheetName);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        ExcelQueryable<Row> WorksheetRange(string startRange, string endRange, int worksheetIndex);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        ExcelQueryable<RowNoHeader> WorksheetNoHeader();

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        ExcelQueryable<RowNoHeader> WorksheetNoHeader(string worksheetName);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        ExcelQueryable<RowNoHeader> WorksheetNoHeader(int worksheetIndex);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string startRange, string endRange);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string startRange, string endRange, string worksheetName);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string startRange, string endRange, int worksheetIndex);

        /// <summary>
        /// Enables Linq queries against a workbook-scope named range
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="namedRangeName">Name of the workbook-scope named range</param>
        ExcelQueryable<TSheetData> NamedRange<TSheetData>(string namedRangeName);

        /// <summary>
        /// Enables Linq queries against a named range
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="namedRangeName">Name of the named range</param>
        ExcelQueryable<TSheetData> NamedRange<TSheetData>(string worksheetName, string namedRangeName);

        /// <summary>
        /// Enables Linq queries against a named range
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="namedRangeName">Name of the named range</param>
        ExcelQueryable<TSheetData> NamedRange<TSheetData>(int worksheetIndex, string namedRangeName);

        /// <summary>
        /// Enables Linq queries against a workbook-scope named range
        /// </summary>
        /// <param name="namedRangeName">Name of the workbook-scope named range</param>
        ExcelQueryable<Row> NamedRange(string namedRangeName);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="namedRangeName">Name of the named range</param>
        ExcelQueryable<Row> NamedRange(string worksheetName, string namedRangeName);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="namedRangeName">Name of the named range</param>
        ExcelQueryable<Row> NamedRange(int worksheetIndex, string namedRangeName);

        /// <summary>
        /// Enables Linq queries against a workbook-scope named range that does not have a header row
        /// </summary>
        /// <param name="namedRangeName">Name of the workbook-scope named range</param>
        ExcelQueryable<RowNoHeader> NamedRangeNoHeader(string namedRangeName);

        /// <summary>
        /// Enables Linq queries against a named range that does not have a header row
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="namedRangeName">Name of the named range</param>
        ExcelQueryable<RowNoHeader> NamedRangeNoHeader(string worksheetName, string namedRangeName);

        /// <summary>
        /// Enables Linq queries against a named range that does not have a header row
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="namedRangeName">Name of the named range</param>
        ExcelQueryable<RowNoHeader> NamedRangeNoHeader(int worksheetIndex, string namedRangeName);

        /// <summary>
        /// Returns a list of worksheet names that the spreadsheet contains
        /// </summary>
        IEnumerable<string> GetWorksheetNames();

        /// <summary>
        /// Returns a list of workbook-scope named ranges that the spreadsheet contains
        /// </summary>
        IEnumerable<string> GetNamedRanges();

        /// <summary>
        /// Returns a list of worksheet-scope named ranges that the worksheet contains
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        IEnumerable<string> GetNamedRanges(string worksheetName);

        /// <summary>
        /// Returns a list of columns names that a named range contains
        /// </summary>
        /// <param name="worksheetName">Worksheet name to get the list of column names from</param>
        /// <param name="namedRangeName">Named Range name to get the list of column names from</param>
        IEnumerable<string> GetColumnNames(string worksheetName, string namedRangeName);

        /// <summary>
        /// Returns a list of columns names that a worksheet contains
        /// </summary>
        /// <param name="worksheetName">Worksheet name to get the list of column names from</param>
        IEnumerable<string> GetColumnNames(string worksheetName);
    }
}
