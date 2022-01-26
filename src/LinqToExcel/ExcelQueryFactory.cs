﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq.Expressions;
using System.Reflection;
using LinqToExcel.Domain;
using LinqToExcel.Logging;
using LinqToExcel.Query;

namespace LinqToExcel
{
    public class ExcelQueryFactory : IExcelQueryFactory
    {
        private readonly Dictionary<string, string> _columnMappings = new Dictionary<string, string>();
        private readonly Dictionary<string, Func<string, object>> _transformations = new Dictionary<string, Func<string, object>>();
        private readonly ILogManagerFactory _logManagerFactory;
        private readonly ILogProvider _log;
        private ExcelQueryArgs _queryArgs;
        private bool _disposed;

	    /// <summary>
        /// Full path to the Excel spreadsheet
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// Confirms all the worksheet columns are mapped to a property, and if not, throws a StrictMappingException
        /// </summary>
        public StrictMappingType? StrictMapping { get; set; }

        /// <summary>
        /// Indicates how to treat leading and trailing spaces in string values.
        /// </summary>
        public TrimSpacesType TrimSpaces { get; set; }

        /// <summary>
        /// Opens the excel file in read-only mode
        /// </summary>
        public bool ReadOnly { get; set; }

        /// <summary>
        /// If true, uses a single, persistent connection for the lifetime of the factory.
        /// If false, a new connection is created for each query
        /// Default is false
        /// </summary>
        public bool UsePersistentConnection { get; set; }

        /// <summary>
        /// If true, the query engine iterates a row at a time.
        /// If false, the entire query is read into a List.
        /// Default is false
        /// </summary>
        public bool Lazy { get; set; }
        
        /// Gets or sets the value of the OLE DB Services flag sent in the connection strings,
        /// which among other features, can disable auto-enlistment in TransactionScopes.
        /// </summary>
        public OleDbServices OleDbServices { get; set; }

        /// <summary>
        /// Gets or sets the value of the Code Page Identifier sent in the connection strings,
        /// this allows for files with different encodings. For the full
        /// list see https://docs.microsoft.com/en-us/windows/desktop/intl/code-page-identifiers
        /// </summary>
        public int? CodePageIdentifier { get; set; }

        /// <summary>
        /// If true, skips all empty rows
        /// Default is false
        /// </summary>
        public bool SkipEmptyRows { get; set; }

        public ExcelQueryFactory()
          : this(null, null) { }

        /// <param name="logManagerFactory">
        /// Factory that facilitates the creation of an external log manager (i.e. log4net) to
        /// allow internal methods of LinqToExcel to perform diagnostic logging.
        /// </param>
        public ExcelQueryFactory(ILogManagerFactory logManagerFactory)
            : this(null, logManagerFactory) { }

        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public ExcelQueryFactory(string fileName)
            : this(fileName, null) { }

        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        /// <param name="logManagerFactory">
        /// Factory that facilitates the creation of an external log manager (i.e. log4net) to
        /// allow internal methods of LinqToExcel to perform diagnostic logging.
        /// </param>
        public ExcelQueryFactory(string fileName, ILogManagerFactory logManagerFactory)
        {
            FileName = fileName;
            DatabaseEngine = ExcelUtilities.DefaultDatabaseEngine();
            OleDbServices = OleDbServices.AllServices;

            if (logManagerFactory != null) {
               _logManagerFactory = logManagerFactory;
               _log = _logManagerFactory.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
            }
        }

        /// <summary>
        /// Sets the database engine to use 
        /// (Spreadsheets ending in xlsx, xlsm, and xlsb must use the Ace database engine)
        /// (If running 64 bit this defaults to ACE (JET doesn't work anyway), if running 32 bit this detaults to JET)
        /// </summary>
        public DatabaseEngine DatabaseEngine { get; set; }

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
            _transformations.Add(string.Format("{0}.{1}", typeof(TSheetData).Name, GetPropertyName(property)), transformation);
        }

        /// <summary>
        /// Returns a list of worksheet names that the spreadsheet contains
        /// </summary>
        public IEnumerable<string> GetWorksheetNames()
        {
            if (String.IsNullOrEmpty(FileName))
                throw new NullReferenceException("FileName property is not set");

            return ExcelUtilities.GetWorksheetNames(FileName, GetQueryArgs());
        }

        /// <summary>
        /// Returns a list of workbook-scope named ranges that the spreadsheet contains
        /// </summary>
        public IEnumerable<string> GetNamedRanges()
        {
            if (String.IsNullOrEmpty(FileName))
                throw new NullReferenceException("FileName property is not set");

            return ExcelUtilities.GetNamedRanges(FileName, GetQueryArgs());
        }

        /// <summary>
        /// Returns a list of worksheet-scope named ranges that the worksheet contains
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        public IEnumerable<string> GetNamedRanges(string worksheetName)
        {
            if (String.IsNullOrEmpty(FileName))
                throw new NullReferenceException("FileName property is not set");

            return ExcelUtilities.GetNamedRanges(FileName, worksheetName, GetQueryArgs());
        }

        /// <summary>
        /// Returns a list of columns names that a worksheet contains
        /// </summary>
        /// <param name="worksheetName">Worksheet name to get the list of column names from</param>
        public IEnumerable<string> GetColumnNames(string worksheetName)
        {
            if (String.IsNullOrEmpty(FileName))
                throw new NullReferenceException("FileName property is not set");

            return ExcelUtilities.GetColumnNames(worksheetName, FileName, GetQueryArgs());
        }

        /// <summary>
        /// Returns a list of columns names that a worksheet contains
        /// </summary>
        /// <param name="worksheetName">Worksheet name to get the list of column names from</param>
        /// <param name="namedRangeName">Named Range name to get the list of column names from</param>
        public IEnumerable<string> GetColumnNames(string worksheetName, string namedRange)
        {
            if (String.IsNullOrEmpty(FileName))
                throw new NullReferenceException("FileName property is not set");

            return ExcelUtilities.GetColumnNames(worksheetName, namedRange, FileName, GetQueryArgs());
        }

        internal ExcelQueryConstructorArgs GetConstructorArgs()
        {
            return new ExcelQueryConstructorArgs()
            {
                FileName = FileName,
              DatabaseEngine = DatabaseEngine,
              StrictMapping = StrictMapping,
                ColumnMappings = _columnMappings,
                Transformations = _transformations,
                UsePersistentConnection = UsePersistentConnection,
                TrimSpaces = TrimSpaces,
                ReadOnly = ReadOnly,
                Lazy = Lazy,
                OleDbServices = OleDbServices,
                CodePageIdentifier = CodePageIdentifier ?? 0,
                SkipEmptyRows = SkipEmptyRows
            };
        }

        internal ExcelQueryArgs GetQueryArgs()
        {
            return new ExcelQueryArgs(GetConstructorArgs());
        }

        #endregion

        #region Worksheet Query Methods

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        public ExcelQueryable<TSheetData> Worksheet<TSheetData>()
        {
            return new ExcelQueryable<TSheetData>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        public ExcelQueryable<Row> Worksheet()
        {
            return new ExcelQueryable<Row>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName)
        {
            return new ExcelQueryable<TSheetData>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetName = worksheetName
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<TSheetData> Worksheet<TSheetData>(int worksheetIndex)
        {
            return new ExcelQueryable<TSheetData>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetIndex = worksheetIndex
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<Row> Worksheet(string worksheetName)
        {
            return new ExcelQueryable<Row>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetName = worksheetName
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<Row> Worksheet(int worksheetIndex)
        {
            return new ExcelQueryable<Row>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetIndex = worksheetIndex
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="startRange"></param>
        /// <param name="endRange"></param>
        public ExcelQueryable<TSheetData> WorksheetRange<TSheetData>(string startRange, string endRange)
        {
            return new ExcelQueryable<TSheetData>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    StartRange = startRange,
                    EndRange = endRange
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        public ExcelQueryable<Row> WorksheetRange(string startRange, string endRange)
        {
            return new ExcelQueryable<Row>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    StartRange = startRange,
                    EndRange = endRange
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<Row> WorksheetRange(string startRange, string endRange, string worksheetName)
        {
            return new ExcelQueryable<Row>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetName = worksheetName,
                    StartRange = startRange,
                    EndRange = endRange
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<Row> WorksheetRange(string startRange, string endRange, int worksheetIndex)
        {
            return new ExcelQueryable<Row>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetIndex = worksheetIndex,
                    StartRange = startRange,
                    EndRange = endRange
                }), _logManagerFactory);
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
            return new ExcelQueryable<TSheetData>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetName = worksheetName,
                    StartRange = startRange,
                    EndRange = endRange
                }), _logManagerFactory);
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
            return new ExcelQueryable<TSheetData>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetIndex = worksheetIndex,
                    StartRange = startRange,
                    EndRange = endRange
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        public ExcelQueryable<RowNoHeader> WorksheetNoHeader()
        {
            return new ExcelQueryable<RowNoHeader>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NoHeader = true
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<RowNoHeader> WorksheetNoHeader(string worksheetName)
        {
            return new ExcelQueryable<RowNoHeader>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NoHeader = true,
                    WorksheetName = worksheetName
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<RowNoHeader> WorksheetNoHeader(int worksheetIndex)
        {
            return new ExcelQueryable<RowNoHeader>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NoHeader = true,
                    WorksheetIndex = worksheetIndex
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        public ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string startRange, string endRange)
        {
            return new ExcelQueryable<RowNoHeader>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NoHeader = true,
                    StartRange = startRange,
                    EndRange = endRange
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetName">Name of the worksheet</param>
        public ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string startRange, string endRange, string worksheetName)
        {
            return new ExcelQueryable<RowNoHeader>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NoHeader = true,
                    StartRange = startRange,
                    EndRange = endRange,
                    WorksheetName = worksheetName
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet that does not have a header row
        /// </summary>
        /// <param name="startRange">Upper left cell name of the range (eg 'B2')</param>
        /// <param name="endRange">Bottom right cell name of the range (eg 'D4')</param>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<RowNoHeader> WorksheetRangeNoHeader(string startRange, string endRange, int worksheetIndex)
        {
            return new ExcelQueryable<RowNoHeader>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NoHeader = true,
                    StartRange = startRange,
                    EndRange = endRange,
                    WorksheetIndex = worksheetIndex
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against a workbook-scope named range
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="namedRangeName">Name of the workbook-scope named range</param>
        public ExcelQueryable<TSheetData> NamedRange<TSheetData>(string namedRangeName)
        {
            return new ExcelQueryable<TSheetData>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NamedRangeName = namedRangeName
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against a worksheet-scope named range
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="namedRangeName">Name of the worksheet-scope named range</param>
        public ExcelQueryable<TSheetData> NamedRange<TSheetData>(string worksheetName, string namedRangeName)
        {
            return new ExcelQueryable<TSheetData>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetName = worksheetName,
                    NamedRangeName = namedRangeName
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against a worksheet-scope named range
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="namedRangeName">Name of the worksheet-scope named range</param>
        public ExcelQueryable<TSheetData> NamedRange<TSheetData>(int worksheetIndex, string namedRangeName)
        {
            return new ExcelQueryable<TSheetData>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetIndex = worksheetIndex,
                    NamedRangeName = namedRangeName
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against a workbook-scope named range
        /// </summary>
        /// <param name="namedRangeName">Name of the workbook-scope named range</param>
        public ExcelQueryable<Row> NamedRange(string namedRangeName)
        {
            return new ExcelQueryable<Row>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NamedRangeName = namedRangeName
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against a worksheet-scope named range
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="namedRangeName">Name of the worksheet-scope named range</param>
        public ExcelQueryable<Row> NamedRange(string worksheetName, string namedRangeName)
        {
            return new ExcelQueryable<Row>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetName = worksheetName,
                    NamedRangeName = namedRangeName
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against a worksheet-scope named range
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="namedRangeName">Name of the worksheet-scope named range</param>
        public ExcelQueryable<Row> NamedRange(int worksheetIndex, string namedRangeName)
        {
            return new ExcelQueryable<Row>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    WorksheetIndex = worksheetIndex,
                    NamedRangeName = namedRangeName
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against a workbook-scope named range that does not have a header row
        /// </summary>
        /// <param name="namedRangeName">Name of the workbook-scope named range</param>
        public ExcelQueryable<RowNoHeader> NamedRangeNoHeader(string namedRangeName)
        {
            return new ExcelQueryable<RowNoHeader>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NoHeader = true,
                    NamedRangeName = namedRangeName
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against a worksheet-scope named range that does not have a header row
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="namedRangeName">Name of the worksheet-scope named range</param>
        public ExcelQueryable<RowNoHeader> NamedRangeNoHeader(int worksheetIndex, string namedRangeName)
        {
            return new ExcelQueryable<RowNoHeader>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NoHeader = true,
                    WorksheetIndex = worksheetIndex,
                    NamedRangeName = namedRangeName
                }), _logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against a worksheet-scope named range that does not have a header row
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="namedRangeName">Name of the worksheet-scope named range</param>
        public ExcelQueryable<RowNoHeader> NamedRangeNoHeader(string worksheetName, string namedRangeName)
        {
            return new ExcelQueryable<RowNoHeader>(PersistQueryArgs(
                new ExcelQueryArgs(GetConstructorArgs())
                {
                    NoHeader = true,
                    WorksheetName = worksheetName,
                    NamedRangeName = namedRangeName
                }), _logManagerFactory);
        }

        private ExcelQueryArgs PersistQueryArgs(ExcelQueryArgs args)
        {
            // We want to keep the persistent connection if there is one
            if (_queryArgs != null)
                args.PersistentConnection = _queryArgs.PersistentConnection;

            return _queryArgs = args;
        }

        #endregion

		#region IDisposable Methods

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~ExcelQueryFactory()
        {
            Dispose(false);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                if (_queryArgs?.PersistentConnection != null)
                {
                    try
                    {
                        _queryArgs.PersistentConnection.Dispose();
                        _queryArgs.PersistentConnection = null;
                    }
                    catch (Exception ex) {
                       _log?.Error("Error disposing OleDbConnection", ex);
                    }
                }
            }

            _queryArgs = null;
            _disposed = true;
        }

		#endregion

		#region Static Methods

		/// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName, string fileName,
                                                                       ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName })
                {
                    WorksheetName = worksheetName
                }, logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(int worksheetIndex, string fileName,
                                                                       ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName })
                {
                    WorksheetIndex = worksheetIndex
                }, logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<Row> Worksheet(string worksheetName, string fileName,
                                                    ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName })
                {
                    WorksheetName = worksheetName
                }, logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        public static ExcelQueryable<Row> Worksheet(int worksheetIndex, string fileName,
                                                    ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName })
                {
                    WorksheetIndex = worksheetIndex
                }, logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        /// <param name="columnMappings">Column to property mappings</param>
        public static ExcelQueryable<Row> Worksheet(string worksheetName, string fileName,
                                                    Dictionary<string, string> columnMappings,
                                                    ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName, ColumnMappings = columnMappings })
                {
                    WorksheetName = worksheetName
                }, logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        /// <param name="columnMappings">Column to property mappings</param>
        public static ExcelQueryable<Row> Worksheet(int worksheetIndex, string fileName,
                                                    Dictionary<string, string> columnMappings,
                                                    ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<Row>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName, ColumnMappings = columnMappings })
                {
                    WorksheetIndex = worksheetIndex
                }, logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        /// <param name="columnMappings">Column to property mappings</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName, string fileName,
                                                                       Dictionary<string, string> columnMappings,
                                                                       ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName, ColumnMappings = columnMappings })
                {
                    WorksheetName = worksheetName
                }, logManagerFactory);
        }

        /// <summary>
        /// Enables Linq queries against an Excel worksheet
        /// </summary>
        /// <typeparam name="TSheetData">Class type to return row data as</typeparam>
        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        /// <param name="fileName">Full path to the Excel spreadsheet</param>
        /// <param name="columnMappings">Column to property mappings</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(int worksheetIndex, string fileName,
                                                                       Dictionary<string, string> columnMappings,
                                                                       ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryable<TSheetData>(
                new ExcelQueryArgs(
                    new ExcelQueryConstructorArgs() { FileName = fileName, ColumnMappings = columnMappings })
                {
                    WorksheetIndex = worksheetIndex
                }, logManagerFactory);
        }

        #endregion
	}
}
