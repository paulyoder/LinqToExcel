using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using Remotion.Data.Linq;
using LinqToExcel.Query;
using System.IO;
using System.Data.OleDb;
using System.Data;
using Remotion.Logging;
using System.Reflection;
using Remotion.Data.Linq.Clauses.ResultOperators;
using Remotion.Data.Linq.Clauses;
using System.Collections;
using Remotion.Data.Linq.Clauses.StreamedData;
using LinqToExcel.Extensions;
using log4net;

namespace LinqToExcel.Query
{
    public class ExcelQueryExecutor : IQueryExecutor
    {
        private readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private readonly string _fileName;
        private readonly Dictionary<string, string> _columnMappings;
        private string _worksheetName;

        public ExcelQueryExecutor(string worksheetName, string fileName,  Dictionary<string, string> columnMappings)
        {
            _fileName = fileName;
            _columnMappings = columnMappings;
            _worksheetName = worksheetName;
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
            var connString = GetConnectionString();
            var sql = GetSqlStatement(queryModel);
            LogSqlStatement(connString, sql);

            var objectResults = GetDataResults(connString, sql, queryModel);
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
            var sqlVisitor = new SqlGeneratorQueryModelVisitor(_worksheetName, _columnMappings);
            sqlVisitor.VisitQueryModel(queryModel);
            return sqlVisitor.SqlStatement;
        }

        /// <summary>
        /// Executes the sql query and returns the data results
        /// </summary>
        /// <typeparam name="T">Data type in the main from clause (queryModel.MainFromClause.ItemType)</typeparam>
        /// <param name="queryModel">Linq query model</param>
        protected IEnumerable<object> GetDataResults(string connectionString, SqlParts sql, QueryModel queryModel)
        {
            IEnumerable<object> results;
            using (var conn = new OleDbConnection(connectionString))
            using (var command = conn.CreateCommand())
            {
                conn.Open();
                command.CommandText = sql.ToString();
                command.Parameters.AddRange(sql.Parameters.ToArray());
                var data = command.ExecuteReader();

                var columns = GetColumnNames(data);
                if (columns.Count() == 1 && columns.First() == "Expr1000")
                    results = GetScalarResults(data);
                else if (queryModel.MainFromClause.ItemType == typeof(Row))
                    results = GetRowResults(data, columns);
                else
                    results = GetTypeResults(data, columns, queryModel);
            }
            return results;
        }

        private string GetConnectionString()
        {
            var connString = "";

            if (_fileName.ToLower().EndsWith("xlsx") ||
                _fileName.ToLower().EndsWith("xlsm"))
                connString = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""",
                    _fileName);
            else if (_fileName.ToLower().EndsWith("xlsb"))
            {
                connString = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                    _fileName);
            }
            else if (_fileName.ToLower().EndsWith("csv"))
            {
                _worksheetName = Path.GetFileName(_fileName);
                connString = string.Format(
                        @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""text;HDR=Yes;FMT=Delimited;IMEX=1""",
                        Path.GetDirectoryName(_fileName));
            }
            else
                connString = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1""",
                    _fileName);

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
                    var columnName = (_columnMappings.ContainsKey(prop.Name)) ?
                        _columnMappings[prop.Name] :
                        prop.Name;
                    if (columns.Contains(columnName))
                        result.SetProperty(prop.Name, data[columnName].Cast(prop.PropertyType));
                }
                results.Add(result);
            } 
            return results.AsEnumerable();
        }

        private IEnumerable<object> GetScalarResults(IDataReader data)
        {
            data.Read();
            return new List<object> { data[0] };
        }

        private void LogSqlStatement(string connectionString, SqlParts sqlParts)
        {
            if (_log.IsDebugEnabled)
            {
                _log.DebugFormat("Connection String: {0}", connectionString);
                _log.DebugFormat("SQL: {0}", sqlParts.ToString());
                for (var i = 0; i < sqlParts.Parameters.Count(); i++)
                    _log.DebugFormat("Param[{0}]: {1}", i, sqlParts.Parameters.ElementAt(i).Value);
            }
        }

        private IEnumerable<string> GetColumnNames(IDataReader data)
        {
            var columns = new List<string>();
            var sheetSchema = data.GetSchemaTable();
            foreach (DataRow row in sheetSchema.Rows)
                columns.Add(row["ColumnName"].ToString());

            //Log a warning for any property to column mappings that do not exist in the excel worksheet
            foreach (var kvp in _columnMappings)
            {
                if (!columns.Contains(kvp.Value))
                    _log.WarnFormat("'{0}' column that is mapped to the '{1}' property does not exist in the '{2}' worksheet",
                        kvp.Value, kvp.Key, _worksheetName);
            }

            return columns;
        }
    }
}
