using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using LinqToExcel.Query;

namespace LinqToExcel
{
    public class ExcelQueryFactory : IExcelQueryFactory
    {
        private readonly Dictionary<string, string> _columnMappings = new Dictionary<string, string>();
        private readonly Dictionary<string, Func<string, object>> _transformations = new Dictionary<string, Func<string, object>>();
        
        /// <summary>
        /// Full path to the Excel spreadsheet
        /// </summary>
        public string FileName { get; set; }

        public ExcelQueryFactory()
            : this(null)
        { }

        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public ExcelQueryFactory(string fileName)
        {
            FileName = fileName;
        }

        /// <summary>
        /// Add a column to property mapping
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="property">Class property to map to</param>
        /// <param name="column">Worksheet column name to map from</param>
        public void AddMapping<TSheetData>(Expression<Func<TSheetData, object>> property, string column)
        {
            _columnMappings[GetPropertyName(property)] = column;
        }

        /// <summary>
        /// Add a column to property mapping with a transformation operation
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="property">Class property to map to</param>
        /// <param name="column">Worksheet column name to map from</param>
        /// <param name="transformation">Lambda expression that transforms the original string value to the desired property value</param>
        public void AddMapping<TSheetData>(Expression<Func<TSheetData, object>> property, string column, Func<string, object> transformation)
        {
            AddMapping(property, column);
            AddTransformation(property, transformation);
        }

        private string GetPropertyName<TSheetData>(Expression<Func<TSheetData, object>> property)
        {
            var exp = (LambdaExpression)property;

            //exp.Body has 2 possible types
            //If the property type is native, then exp.Body == typeof(MemberExpression)
            //If the property type is not native, then exp.Body == typeof(UnaryExpression) in which 
            //case we can get the MemberExpression from its Operand property
            var mExp = (exp.Body.NodeType == ExpressionType.MemberAccess) ?
                (MemberExpression)exp.Body :
                (MemberExpression)((UnaryExpression)exp.Body).Operand;
            return mExp.Member.Name;
        }

        /// <summary>
        /// Transforms a cell value in the spreadsheet to the desired property value
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="property">Class property value to transform</param>
        /// <param name="transformation">Lambda expression that transforms the original string value to the desired property value</param>
        /// <example>
        /// AddTransformation{Person}(p => p.IsActive, x => x == "Y");
        /// AddTransformation{Person}(p => p.IsYoung, x => DateTime.Parse(x) > new DateTime(2000, 1, 1));
        /// </example>
        public void AddTransformation<TSheetData>(Expression<Func<TSheetData, object>> property, Func<string, object> transformation)
        {
            _transformations.Add(GetPropertyName(property), transformation);
        }

        #region Worksheet Query Methods

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        public ExcelQueryable<TSheetData> Worksheet<TSheetData>()
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        public ExcelQueryable<Row> Worksheet()
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    WorksheetName = worksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<TSheetData> Worksheet<TSheetData>(int worksheetIndex)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    WorksheetIndex = worksheetIndex
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<Row> Worksheet(string worksheetName)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    WorksheetName = worksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<Row> Worksheet(int worksheetIndex)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    WorksheetIndex = worksheetIndex
                });
        }


        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="startRange"></param>
        /// <param name="endRange"></param>
        public ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string startRange, string endRange)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    StartRange = startRange,
                    EndRange = endRange
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        public ExcelQueryable<Row> WorksheetRange(string startRange, string endRange)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    StartRange = startRange,
                    EndRange = endRange
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<Row> WorksheetRange(string startRange, string endRange, string worksheetName)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    WorksheetName = worksheetName,
                    StartRange = startRange,
                    EndRange = endRange
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<Row> WorksheetRange(string startRange, string endRange, int worksheetIndex)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    WorksheetIndex = worksheetIndex,
                    StartRange = startRange,
                    EndRange = endRange
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string startRange, string endRange, string worksheetName)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    WorksheetName = worksheetName,
                    StartRange = startRange,
                    EndRange = endRange
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string startRange, string endRange, int worksheetIndex)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    WorksheetIndex = worksheetIndex,
                    StartRange = startRange,
                    EndRange = endRange
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        public ExcelQueryable<RowNoHeader> WorksheetNoHeader()
        {
            return new ExcelQueryable<RowNoHeader>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    NoHeader = true
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<RowNoHeader> WorksheetNoHeader(string worksheetName)
        {
            return new ExcelQueryable<RowNoHeader>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    NoHeader = true,
                    WorksheetName = worksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<RowNoHeader> WorksheetNoHeader(int worksheetIndex)
        {
            return new ExcelQueryable<RowNoHeader>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    NoHeader = true,
                    WorksheetIndex = worksheetIndex
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        public ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string startRange, string endRange)
        {
            return new ExcelQueryable<RowNoHeader>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    NoHeader = true,
                    StartRange = startRange,
                    EndRange = endRange
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string startRange, string endRange, string worksheetName)
        {
            return new ExcelQueryable<RowNoHeader>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    NoHeader = true,
                    StartRange = startRange,
                    EndRange = endRange,
                    WorksheetName = worksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string startRange, string endRange, int worksheetIndex)
        {
            return new ExcelQueryable<RowNoHeader>(
                new ExcelQueryArgs(FileName, _columnMappings, _transformations)
                {
                    NoHeader = true,
                    StartRange = startRange,
                    EndRange = endRange,
                    WorksheetIndex = worksheetIndex
                });
        }

        #endregion

        #region Static Methods

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName, string fileName)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(fileName, null, null)
                {
                    WorksheetName = worksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(int worksheetIndex, string fileName)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(fileName, null, null)
                {
                    WorksheetIndex = worksheetIndex
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<Row> Worksheet(string worksheetName, string fileName)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(fileName, null, null)
                {
                    WorksheetName = worksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<Row> Worksheet(int worksheetIndex, string fileName)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(fileName, null, null)
                {
                    WorksheetIndex = worksheetIndex
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        /// <param name="columnMappings">Column to property mappings</param>
        public static ExcelQueryable<Row> Worksheet(string worksheetName, string fileName, Dictionary<string, string> columnMappings)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(fileName, columnMappings, null)
                {
                    WorksheetName = worksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        /// <param name="columnMappings">Column to property mappings</param>
        public static ExcelQueryable<Row> Worksheet(int worksheetIndex, string fileName, Dictionary<string, string> columnMappings)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(fileName, columnMappings, null)
                {
                    WorksheetIndex = worksheetIndex
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        /// <param name="columnMappings">Column to property mappings</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName, string fileName, Dictionary<string, string> columnMappings)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(fileName, columnMappings, null)
                {
                    WorksheetName = worksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        /// <param name="columnMappings">Column to property mappings</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(int worksheetIndex, string fileName, Dictionary<string, string> columnMappings)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(fileName, columnMappings, null)
                {
                    WorksheetIndex = worksheetIndex
                });
        }

        #endregion
    }
}
