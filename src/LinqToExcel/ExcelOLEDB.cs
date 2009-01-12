using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using Microsoft.VisualStudio.DebuggerVisualizers;
using System.Data.OleDb;
using log4net;
using System.Reflection;
using LinqToExcel.Extensions.Reflection;
using System.Data;

namespace LinqToExcel
{
    /// <summary>
    /// Queries the Excel worksheet using an OLEDB connection
    /// </summary>
    public class ExcelOLEDB
    {
        private static ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        /// <summary>
        /// Executes the query based upon the Linq statement against the Excel worksheet
        /// </summary>
        /// <param name="expression">Expression created from the Linq statement</param>
        /// <param name="fileName">File path to the Excel workbook</param>
        /// <param name="columnMapping">
        /// Property to column mapping. 
        /// Properties are the dictionary keys and the dictionary values are the corresponding column names.
        /// </param>
        /// <param name="worksheetName">Name of the Excel worksheet</param>
        /// <returns>Returns the results from the query</returns>
        public object ExecuteQuery(Expression expression, string fileName, Dictionary<string, string> columnMapping, string worksheetName)
        {
            Type dataType = expression.Type.GetGenericArguments()[0];
            PropertyInfo[] props = dataType.GetProperties();

            //Build the SQL string
            SQLExpressionVisitor sql = new SQLExpressionVisitor();
            sql.BuildSQLStatement(expression, columnMapping, worksheetName, dataType);

            string connString = string.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties= ""Excel 8.0;HDR=YES;""", fileName);
            if (_log.IsDebugEnabled) _log.Debug("Connection String: " + connString);

            object results = Activator.CreateInstance(typeof(List<>).MakeGenericType(dataType));
            using (OleDbConnection conn = new OleDbConnection(connString))
            using (OleDbCommand command = conn.CreateCommand())
            {                
                conn.Open();
                command.CommandText = sql.SQLStatement;
                command.Parameters.Clear();
                command.Parameters.AddRange(sql.Parameters.ToArray());
                OleDbDataReader data = command.ExecuteReader();
                
                //Get the excel column names
                List<string> columns = new List<string>();
                DataTable sheetSchema = data.GetSchemaTable();
                foreach (DataRow row in sheetSchema.Rows)
                    columns.Add(row["ColumnName"].ToString());

                //Log a warning for any property to column mappings that do not exist in the excel worksheet
                foreach (KeyValuePair<string, string> kvp in columnMapping)
                {
                    if (!columns.Contains(kvp.Value))
                        _log.Warn(string.Format("'{0}' column that is mapped to the '{1}' property does not exist in the '{2}' worksheet",
                                                kvp.Value, kvp.Key, "Sheet1"));
                }

                if (dataType == typeof(Row))
                {
                    Dictionary<string, int> columnIndexMapping = new Dictionary<string, int>();
                    for (int i = 0; i < columns.Count; i++)
                        columnIndexMapping[columns[i]] = i;
                    
                    while (data.Read())
                    {
                        IList<Cell> cells = new List<Cell>();
                        for (int i = 0; i < columns.Count; i++)
                            cells.Add(new Cell(data[i]));
                        results.CallMethod("Add", new Row(cells, columnIndexMapping));
                    }
                }
                else
                {
                    while (data.Read())
                    {
                        object result = Activator.CreateInstance(dataType);
                        foreach (PropertyInfo prop in props)
                        {
                            //Set the column name to the property mapping if there is one, else use the property name for the column name
                            string columnName = (columnMapping.ContainsKey(prop.Name)) ? columnMapping[prop.Name] : prop.Name;
                            if (columns.Contains(columnName))
                                result.SetProperty(prop.Name, Convert.ChangeType(data[columnName], prop.PropertyType));
                        }
                        results.CallMethod("Add", result);
                    }
                }
            }
            return results;
        }
    }
}
