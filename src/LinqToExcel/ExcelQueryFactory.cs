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

        /// <summary>
        /// Confirms all the worksheet columns are mapped to a property, and if not, throws a StrictMappingException
        /// </summary>
        public bool StrictMapping { get; set; }

        public ExcelQueryFactory()
            : this(null)
        { }

        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public ExcelQueryFactory(string fileName)
        {
            this.FileName = fileName;
        }

        #region Other Methods

        /// <summary>
        /// Add a column to property mapping
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="property">Class property to map to</param>
        /// <param name="column">Worksheet column name to map from</param>
        public void AddMapping<TSheetData>(Expression<Func<TSheetData, object>> property, string column)
        {
            AddMapping(GetPropertyName(property), column);
        }

        /// <summary>
        /// Add a column to property mapping
        /// </summary>
        /// <param name="propertyName">Class property to map to</param>
        /// <param name="column">Worksheet column name to map from</param>
        public void AddMapping(string propertyName, string column)
        {
            _columnMappings[propertyName] = column;
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

        /// <summary>
        /// Returns a list of worksheet names that the spreadsheet contains
        /// </summary>
        public IEnumerable<string> GetWorksheetNames()
        {
            if (String.IsNullOrEmpty(FileName))
                throw new NullReferenceException("FileName property is not set");

            return ExcelUtilities.GetWorksheetNames(FileName);
        }

        /// <summary>
        /// Returns a list of columns names that a worksheet contains
        /// </summary>
        /// <param name="worksheetName">Worksheet name to get the list of column names from</param>
        public IEnumerable<string> GetColumnNames(string worksheetName)
        {
            if (String.IsNullOrEmpty(FileName))
                throw new NullReferenceException("FileName property is not set");

            return ExcelUtilities.GetColumnNames(worksheetName, FileName);
        }

        internal ExcelQueryConstructorArgs GetConstructorArgs()
        {
            return new ExcelQueryConstructorArgs()
            {
                FileName = FileName,
                StrictMapping = StrictMapping,
                ColumnMappings = _columnMappings,
                Transformations = _transformations
            };
        }

        #endregion

        #region Worksheet Query Methods

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        public ExcelQueryable<TSheetData> Worksheet<TSheetData>()
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(GetConstructorArgs()));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        public ExcelQueryable<Row> Worksheet()
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(GetConstructorArgs()));
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(GetConstructorArgs())
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
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName })
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
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName })
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
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName })
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
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName })
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
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName, ColumnMappings = columnMappings })
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
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName, ColumnMappings = columnMappings })
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
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName, ColumnMappings = columnMappings })
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
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName, ColumnMappings = columnMappings })
                {
                    WorksheetIndex = worksheetIndex
                });
        }

        #endregion
    }
}
