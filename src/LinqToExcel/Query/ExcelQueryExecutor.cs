using System;
using System.Collections.Generic;
using System.Linq;
using Remotion.Data.Linq;
using System.IO;
using System.Data.OleDb;
using System.Data;
using System.Reflection;
using Remotion.Data.Linq.Clauses.ResultOperators;
using System.Collections;
using LinqToExcel.Extensions;
using log4net;
using System.Text.RegularExpressions;
using System.Text;

namespace LinqToExcel.Query
{
    public class ExcelQueryExecutor : IQueryExecutor
    {
        private readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private readonly ExcelQueryArgs _args;
        private readonly string _connectionString;

        public ExcelQueryExecutor(ExcelQueryArgs args)
        {
            ValidateArgs(args);
            _args = args;
            _connectionString = GetConnectionString();
            GetWorksheetName();
        }

        private void ValidateArgs(ExcelQueryArgs args)
        {
            if (_log.IsDebugEnabled)
                _log.DebugFormat("ExcelQueryArgs = {0}", args.ToString());

            if (args.FileName == null)
                throw new ArgumentNullException("FileName", "FileName property cannot be null.");

            if (!String.IsNullOrEmpty(args.StartRange) &&
                !Regex.Match(args.StartRange, "^[a-zA-Z]{1,3}[0-9]{1,7}$").Success)
                throw new ArgumentException(string.Format(
                    "StartRange argument '{0}' is invalid format for cell name", args.StartRange));

            if (!String.IsNullOrEmpty(args.EndRange) &&
                !Regex.Match(args.EndRange, "^[a-zA-Z]{1,3}[0-9]{1,7}$").Success)
                throw new ArgumentException(string.Format(
                    "EndRange argument '{0}' is invalid format for cell name", args.EndRange));

            if (args.NoHeader &&
                !String.IsNullOrEmpty(args.StartRange) &&
                args.FileName.ToLower().Contains(".csv"))
                throw new ArgumentException("Cannot use WorksheetRangeNoHeader on csv files");
        }

        /// <summary>
        /// Executes a query with a scalar result, i.e. a query that ends with a result operator such as Count, Sum, or Average.
        /// </summary>
        public T ExecuteScalar<T>(QueryModel queryModel)
        {
            return ExecuteSingle<T>(queryModel, false);
        }

        /// <summary>
        /// Executes a query with a single result object, i.e. a query that ends with a result operator such as First, Last, Single, Min, or Max.
        /// </summary>
        public T ExecuteSingle<T>(QueryModel queryModel, bool returnDefaultWhenEmpty)
        {
            var results = ExecuteCollection<T>(queryModel);

            foreach (var resultOperator in queryModel.ResultOperators)
            {
                if (resultOperator is LastResultOperator)
                    return results.Last();
            }

            return (returnDefaultWhenEmpty) ?
                results.FirstOrDefault() :
                results.First();
        }

        /// <summary>
        /// Executes a query with a collection result.
        /// </summary>
        public IEnumerable<T> ExecuteCollection<T>(QueryModel queryModel)
        {
            var sql = GetSqlStatement(queryModel);
            LogSqlStatement(sql);

            var objectResults = GetDataResults(sql, queryModel);
            var projector = GetSelectProjector<T>(objectResults.FirstOrDefault(), queryModel);
            var returnResults = objectResults.Cast<T>(projector);

            foreach (var resultOperator in queryModel.ResultOperators)
            {
                if (resultOperator is ReverseResultOperator)
                    returnResults = returnResults.Reverse();
                if (resultOperator is SkipResultOperator)
                    returnResults = returnResults.Skip(resultOperator.Cast<SkipResultOperator>().GetConstantCount());
            }

            return returnResults;
        }

        protected Func<object, T> GetSelectProjector<T>(object firstResult, QueryModel queryModel)
        {
            Func<object, T> projector = (result) => result.Cast<T>();
            if ((firstResult.GetType() != typeof(T)) &&
                (typeof(T) != typeof(int)) &&
                (typeof(T) != typeof(long)))
            {
                var proj = ProjectorBuildingExpressionTreeVisitor.BuildProjector<T>(queryModel.SelectClause.Selector);
                projector = (result) => proj(new ResultObjectMapping(queryModel.MainFromClause, result));
            }
            return projector;
        }

        protected SqlParts GetSqlStatement(QueryModel queryModel)
        {
            var sqlVisitor = new SqlGeneratorQueryModelVisitor(_args);
            sqlVisitor.VisitQueryModel(queryModel);
            return sqlVisitor.SqlStatement;
        }

        private void GetWorksheetName()
        {
            if (_args.FileName.ToLower().EndsWith("csv"))
                _args.WorksheetName = Path.GetFileName(_args.FileName);
            else if (_args.WorksheetIndex.HasValue)
            {
                var worksheetNames = GetWorksheetNames();
                if (_args.WorksheetIndex.Value < worksheetNames.Count())
                    _args.WorksheetName = worksheetNames.ElementAt(_args.WorksheetIndex.Value);
                else
                    throw new DataException("Worksheet Index Out of Range");
            }
            else if (String.IsNullOrEmpty(_args.WorksheetName))
                _args.WorksheetName = "Sheet1";
        }

        private IEnumerable<string> GetWorksheetNames()
        {
            var worksheetNames = new List<string>();
            using (var conn = new OleDbConnection(_connectionString))
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

        /// <summary>
        /// Executes the sql query and returns the data results
        /// </summary>
        /// <typeparam name="T">Data type in the main from clause (queryModel.MainFromClause.ItemType)</typeparam>
        /// <param name="queryModel">Linq query model</param>
        protected IEnumerable<object> GetDataResults(SqlParts sql, QueryModel queryModel)
        {
            IEnumerable<object> results;
            OleDbDataReader data = null;
            using (var conn = new OleDbConnection(_connectionString))
            using (var command = conn.CreateCommand())
            {
                conn.Open();
                command.CommandText = sql.ToString();
                command.Parameters.AddRange(sql.Parameters.ToArray());
                try { data = command.ExecuteReader(); }
                catch (OleDbException e)
                {
                    if (e.Message.Contains(_args.WorksheetName))
                        throw new DataException(
                            string.Format("'{0}' is not a valid worksheet name. Valid worksheet names are: '{1}'",
                                          _args.WorksheetName, string.Join("', '", GetWorksheetNames().ToArray())));
                    if (!CheckIfInvalidColumnNameUsed(sql))
                        throw e;
                }

                var columns = GetColumnNames(data);
                if (columns.Count() == 1 && columns.First() == "Expr1000")
                    results = GetScalarResults(data);
                else if (queryModel.MainFromClause.ItemType == typeof(Row))
                    results = GetRowResults(data, columns);
                else if (queryModel.MainFromClause.ItemType == typeof(RowNoHeader))
                    results = GetRowNoHeaderResults(data);
                else
                    results = GetTypeResults(data, columns, queryModel);
            }
            return results;
        }

        private bool CheckIfInvalidColumnNameUsed(SqlParts sql)
        {
            var usedColumns = sql.ColumnNamesUsed;
            var tableColumns = GetColumnNames();
            foreach (var column in usedColumns)
            {
                if (!tableColumns.Contains(column))
                {
                    throw new DataException(string.Format(
                        "'{0}' is not a valid column name. " +
                        "Valid column names are: '{1}'",
                        column,
                        string.Join("', '", tableColumns.ToArray())));
                }
            }
            return false;
        }

        private string GetConnectionString()
        {
            var connString = "";

            if (_args.FileName.ToLower().EndsWith("xlsx") ||
                _args.FileName.ToLower().EndsWith("xlsm"))
                connString = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""",
                    _args.FileName);
            else if (_args.FileName.ToLower().EndsWith("xlsb"))
            {
                connString = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                    _args.FileName);
            }
            else if (_args.FileName.ToLower().EndsWith("csv"))
            {
                connString = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""text;HDR=YES;FMT=Delimited;IMEX=1""",
                    Path.GetDirectoryName(_args.FileName));
            }
            else
                connString = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1""",
                    _args.FileName);

            if (_args.NoHeader)
                connString = connString.Replace("HDR=YES", "HDR=NO");

            if (_log.IsDebugEnabled)
                _log.DebugFormat("Connection String: {0}", connString);

            return connString;
        }

        private IEnumerable<object> GetRowResults(IDataReader data, IEnumerable<string> columns)
        {
            var results = new List<object>();
            var columnIndexMapping = new Dictionary<string, int>();
            for (var i = 0; i < columns.Count(); i++)
                columnIndexMapping[columns.ElementAt(i)] = i;

            while (data.Read())
            {
                IList<Cell> cells = new List<Cell>();
                for (var i = 0; i < columns.Count(); i++)
                    cells.Add(new Cell(data[i]));
                results.CallMethod("Add", new Row(cells, columnIndexMapping));
            }
            return results.AsEnumerable();
        }

        private IEnumerable<object> GetRowNoHeaderResults(OleDbDataReader data)
        {
            var results = new List<object>();
            while (data.Read())
            {
                IList<Cell> cells = new List<Cell>();
                for (var i = 0; i < data.FieldCount; i++)
                    cells.Add(new Cell(data[i]));
                results.CallMethod("Add", new RowNoHeader(cells));
            }
            return results.AsEnumerable();
        }

        private IEnumerable<object> GetTypeResults(IDataReader data, IEnumerable<string> columns, QueryModel queryModel)
        {
            var results = new List<object>();
            var fromType = queryModel.MainFromClause.ItemType;
            var props = fromType.GetProperties();
            while (data.Read())
            {
                var result = Activator.CreateInstance(fromType);
                foreach (var prop in props)
                {
                    var columnName = (_args.ColumnMappings.ContainsKey(prop.Name)) ?
                        _args.ColumnMappings[prop.Name] :
                        prop.Name;
                    if (columns.Contains(columnName))
                        result.SetProperty(prop.Name, GetColumnValue(data, columnName, prop.Name).Cast(prop.PropertyType));
                }
                results.Add(result);
            } 
            return results.AsEnumerable();
        }

        private object GetColumnValue(IDataRecord data, string columnName, string propertyName)
        {
            //Perform the property transformation if there is one
            return (_args.Transformations.ContainsKey(propertyName)) ?
                _args.Transformations[propertyName](data[columnName].ToString()) :
                data[columnName];
        }

        private IEnumerable<object> GetScalarResults(IDataReader data)
        {
            data.Read();
            return new List<object> { data[0] };
        }

        private void LogSqlStatement(SqlParts sqlParts)
        {
            if (_log.IsDebugEnabled)
            {
                var logMessage = new StringBuilder();
                logMessage.AppendFormat("{0};", sqlParts.ToString());
                for (var i = 0; i < sqlParts.Parameters.Count(); i++)
                {
                    var paramValue = sqlParts.Parameters.ElementAt(i).Value.ToString();
                    var paramMessage = string.Format(" p{0} = '{1}';",
                        i, sqlParts.Parameters.ElementAt(i).Value.ToString());

                    if (paramValue.IsNumber())
                        paramMessage = paramMessage.Replace("'", "");
                    logMessage.Append(paramMessage);
                }
                
                var sqlLog = LogManager.GetLogger("LinqToExcel.SQL");
                sqlLog.Debug(logMessage.ToString());
            }
        }

        private IEnumerable<string> GetColumnNames()
        {
            var columns = new List<string>();
            using (var conn = new OleDbConnection(_connectionString))
            using (var command = conn.CreateCommand())
            {
                conn.Open();
                command.CommandText = string.Format("SELECT TOP 1 * FROM [{0}$]", _args.WorksheetName);
                var data = command.ExecuteReader();
                columns.AddRange(GetColumnNames(data));
            }
            return columns;
        }

        private IEnumerable<string> GetColumnNames(IDataReader data)
        {
            var columns = new List<string>();
            var sheetSchema = data.GetSchemaTable();
            foreach (DataRow row in sheetSchema.Rows)
                columns.Add(row["ColumnName"].ToString());

            //Log a warning for any property to column mappings that do not exist in the excel worksheet
            foreach (var kvp in _args.ColumnMappings)
            {
                if (!columns.Contains(kvp.Value))
                    _log.WarnFormat("'{0}' column that is mapped to the '{1}' property does not exist in the '{2}' worksheet",
                        kvp.Value, kvp.Key, _args.WorksheetName);
            }

            return columns;
        }
    }
}
