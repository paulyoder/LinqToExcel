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
    public class ExcelSQL
    {
        private static ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public object ExecuteQuery(Expression expression, string fileName)
        {
            Type dataType = expression.Type.GetGenericArguments()[0];
            PropertyInfo[] props = dataType.GetProperties();

            //Build the SQL string
            ExpressionToSQL sql = new ExpressionToSQL();
            sql.BuildSQLStatement(expression);
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

                while (data.Read())
                {
                    object result = Activator.CreateInstance(dataType);
                    foreach (PropertyInfo prop in props)
                    {
                        //Make sure there is a corresponding column for the class property
                        if (columns.Contains(prop.Name))
                            result.SetProperty(prop.Name, Convert.ChangeType(data[prop.Name], prop.PropertyType));
                    }
                    results.CallMethod("Add", result);
                }
            }
            return results;
        }
    }
}
