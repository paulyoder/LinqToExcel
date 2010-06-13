using System;
using System.Linq.Expressions;
using LinqToExcel.Query;
using System.Collections.Generic;

namespace LinqToExcel
{
    public interface IExcelQueryFactory
    {
        /// <summary>
        /// Full path to the Excel spreadsheet
        /// </summary>
        string FileName { get; set; }

        /// <summary>
        /// Add a column to property mapping
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="Property">Class property to map to</param>
        /// <param name="Column">Worksheet column name to map from</param>
        void AddMapping<TSheetData>(Expression<Func<TSheetData, object>> Property, string Column);

        /// <summary>
        /// Add a column to property mapping with a transformation operation
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="Property">Class property to map to</param>
        /// <param name="Column">Worksheet column name to map from</param>
        /// <param name="Transformation">Lambda expression that transforms a cell value in the spreadsheet to the desired property value</param>
        void AddMapping<TSheetData>(Expression<Func<TSheetData, object>> Property, string Column, Func<string, object> Transformation);

        /// <summary>
        /// Transforms a cell value in the spreadsheet to the desired property value
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="Property">Class property value to transform</param>
        /// <param name="Transformation">Lambda expression that transforms a cell value in the spreadsheet to the desired property value</param>
        /// <example>
        /// AddTransformation{Person}(p => p.IsActive, x => x == "Y");
        /// AddTransformation{Person}(p => p.IsYoung, x => DateTime.Parse(x) > new DateTime(2000, 1, 1));
        /// </example>
        void AddTransformation<TSheetData>(Expression<Func<TSheetData, object>> Property, Func<string, object> Transformation);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        ExcelQueryable<TSheetData> Worksheet<TSheetData>();

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="WorksheetName">Name of the worksheet</param>
        ExcelQueryable<TSheetData> Worksheet<TSheetData>(string WorksheetName);
        
        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="WorksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        ExcelQueryable<TSheetData> Worksheet<TSheetData>(int WorksheetIndex);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        ExcelQueryable<Row> Worksheet();

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="WorksheetName">Name of the worksheet</param>
        ExcelQueryable<Row> Worksheet(string WorksheetName);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="WorksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        ExcelQueryable<Row> Worksheet(int WorksheetIndex);
        
        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string StartRange, string EndRange);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="WorksheetName">Name of the worksheet</param>
        ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string StartRange, string EndRange, string WorksheetName);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="WorksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string StartRange, string EndRange, int WorksheetIndex);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        ExcelQueryable<Row> WorksheetRange(string StartRange, string EndRange);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        ExcelQueryable<Row> WorksheetRange(string StartRange, string EndRange, string WorksheetName);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        ExcelQueryable<Row> WorksheetRange(string StartRange, string EndRange, int WorksheetIndex);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        ExcelQueryable<RowNoHeader> WorksheetNoHeader();

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="WorksheetName">Name of the worksheet</param>
        ExcelQueryable<RowNoHeader> WorksheetNoHeader(string WorksheetName);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="WorksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        ExcelQueryable<RowNoHeader> WorksheetNoHeader(int WorksheetIndex);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string StartRange, string EndRange);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="WorksheetName">Name of the worksheet</param>
        ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string StartRange, string EndRange, string WorksheetName);

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="WorksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string StartRange, string EndRange, int WorksheetIndex);

        /// <summary>
        /// Returns a list of worksheet names that the spreadsheet contains
        /// </summary>
        IEnumerable<string> GetWorksheetNames();

        /// <summary>
        /// Returns a list of columns names that a worksheet contains
        /// </summary>
        /// <param name="WorksheetName">Worksheet name to get the list of column names from</param>
        IEnumerable<string> GetColumnNames(string WorksheetName);
    }
}
