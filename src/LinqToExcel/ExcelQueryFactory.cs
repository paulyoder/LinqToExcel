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

        /// <param name="FileName">Full path to the Excel spreadsheet</param>
        public ExcelQueryFactory(string FileName)
        {
            this.FileName = FileName;
        }

        #region Other Methods

        /// <summary>
        /// Add a column to property mapping
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="Property">Class property to map to</param>
        /// <param name="Column">Worksheet column name to map from</param>
        public void AddMapping<TSheetData>(Expression<Func<TSheetData, object>> Property, string Column)
        {
            _columnMappings[GetPropertyName(Property)] = Column;
        }

        /// <summary>
        /// Add a column to property mapping with a transformation operation
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="Property">Class property to map to</param>
        /// <param name="Column">Worksheet column name to map from</param>
        /// <param name="Transformation">Lambda expression that transforms the original string value to the desired property value</param>
        public void AddMapping<TSheetData>(Expression<Func<TSheetData, object>> Property, string Column, Func<string, object> Transformation)
        {
            AddMapping(Property, Column);
            AddTransformation(Property, Transformation);
        }

        private string GetPropertyName<TSheetData>(Expression<Func<TSheetData, object>> Property)
        {
            var exp = (LambdaExpression)Property;

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
        /// <param name="Property">Class property value to transform</param>
        /// <param name="Transformation">Lambda expression that transforms the original string value to the desired property value</param>
        /// <example>
        /// AddTransformation{Person}(p => p.IsActive, x => x == "Y");
        /// AddTransformation{Person}(p => p.IsYoung, x => DateTime.Parse(x) > new DateTime(2000, 1, 1));
        /// </example>
        public void AddTransformation<TSheetData>(Expression<Func<TSheetData, object>> Property, Func<string, object> Transformation)
        {
            _transformations.Add(GetPropertyName(Property), Transformation);
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
        /// <param name="WorksheetName">Worksheet name to get the list of column names from</param>
        public IEnumerable<string> GetColumnNames(string WorksheetName)
        {
            if (String.IsNullOrEmpty(FileName))
                throw new NullReferenceException("FileName property is not set");

            return ExcelUtilities.GetColumnNames(WorksheetName, FileName);
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
        /// <param name="WorksheetName">Name of the worksheet</param>
        public ExcelQueryable<TSheetData> Worksheet<TSheetData>(string WorksheetName)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetName = WorksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="WorksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<TSheetData> Worksheet<TSheetData>(int WorksheetIndex)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetIndex = WorksheetIndex
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="WorksheetName">Name of the worksheet</param>
        public ExcelQueryable<Row> Worksheet(string WorksheetName)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetName = WorksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="WorksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<Row> Worksheet(int WorksheetIndex)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetIndex = WorksheetIndex
                });
        }


        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="StartRange"></param>
        /// <param name="EndRange"></param>
        public ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string StartRange, string EndRange)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    StartRange = StartRange,
                    EndRange = EndRange
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        public ExcelQueryable<Row> WorksheetRange(string StartRange, string EndRange)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    StartRange = StartRange,
                    EndRange = EndRange
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="WorksheetName">Name of the worksheet</param>
        public ExcelQueryable<Row> WorksheetRange(string StartRange, string EndRange, string WorksheetName)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetName = WorksheetName,
                    StartRange = StartRange,
                    EndRange = EndRange
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="WorksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<Row> WorksheetRange(string StartRange, string EndRange, int WorksheetIndex)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetIndex = WorksheetIndex,
                    StartRange = StartRange,
                    EndRange = EndRange
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="WorksheetName">Name of the worksheet</param>
        public ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string StartRange, string EndRange, string WorksheetName)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetName = WorksheetName,
                    StartRange = StartRange,
                    EndRange = EndRange
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="WorksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string StartRange, string EndRange, int WorksheetIndex)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetIndex = WorksheetIndex,
                    StartRange = StartRange,
                    EndRange = EndRange
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
        /// <param name="WorksheetName">Name of the worksheet</param>
        public ExcelQueryable<RowNoHeader> WorksheetNoHeader(string WorksheetName)
        {
            return new ExcelQueryable<RowNoHeader>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NoHeader = true,
                    WorksheetName = WorksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="WorksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<RowNoHeader> WorksheetNoHeader(int WorksheetIndex)
        {
            return new ExcelQueryable<RowNoHeader>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NoHeader = true,
                    WorksheetIndex = WorksheetIndex
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="StartRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="EndRange">Bottom right cell name of the range (eg 'D4')</param>
        public ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string StartRange, string EndRange)
        {
            return new ExcelQueryable<RowNoHeader>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NoHeader = true,
                    StartRange = StartRange,
                    EndRange = EndRange
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string StartRange, string EndRange, string WorksheetName)
        {
            return new ExcelQueryable<RowNoHeader>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NoHeader = true,
                    StartRange = StartRange,
                    EndRange = EndRange,
                    WorksheetName = WorksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string StartRange, string EndRange, int WorksheetIndex)
        {
            return new ExcelQueryable<RowNoHeader>(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NoHeader = true,
                    StartRange = StartRange,
                    EndRange = EndRange,
                    WorksheetIndex = WorksheetIndex
                });
        }

        #endregion

        #region Static Methods

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="WorksheetName">Name of the worksheet</param>
        /// <param name="FileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string WorksheetName, string FileName)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = FileName })
                {
                    WorksheetName = WorksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="WorksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="FileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(int WorksheetIndex, string FileName)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = FileName })
                {
                    WorksheetIndex = WorksheetIndex
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="WorksheetName">Name of the worksheet</param>
        /// <param name="FileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<Row> Worksheet(string WorksheetName, string FileName)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = FileName })
                {
                    WorksheetName = WorksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="WorksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="FileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<Row> Worksheet(int WorksheetIndex, string FileName)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = FileName })
                {
                    WorksheetIndex = WorksheetIndex
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="WorksheetName">Name of the worksheet</param>
        /// <param name="FileName">Full path to the Excel spreadsheet</param>
        /// <param name="ColumnMappings">Column to property mappings</param>
        public static ExcelQueryable<Row> Worksheet(string WorksheetName, string FileName, Dictionary<string, string> ColumnMappings)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = FileName, ColumnMappings = ColumnMappings })
                {
                    WorksheetName = WorksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="WorksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="FileName">Full path to the Excel spreadsheet</param>
        /// <param name="ColumnMappings">Column to property mappings</param>
        public static ExcelQueryable<Row> Worksheet(int WorksheetIndex, string FileName, Dictionary<string, string> ColumnMappings)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = FileName, ColumnMappings = ColumnMappings })
                {
                    WorksheetIndex = WorksheetIndex
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="WorksheetName">Name of the worksheet</param>
        /// <param name="FileName">Full path to the Excel spreadsheet</param>
        /// <param name="ColumnMappings">Column to property mappings</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string WorksheetName, string FileName, Dictionary<string, string> ColumnMappings)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = FileName, ColumnMappings = ColumnMappings })
                {
                    WorksheetName = WorksheetName
                });
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="WorksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="FileName">Full path to the Excel spreadsheet</param>
        /// <param name="ColumnMappings">Column to property mappings</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(int WorksheetIndex, string FileName, Dictionary<string, string> ColumnMappings)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = FileName, ColumnMappings = ColumnMappings })
                {
                    WorksheetIndex = WorksheetIndex
                });
        }

        #endregion
    }
}
