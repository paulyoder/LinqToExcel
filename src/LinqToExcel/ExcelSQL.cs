using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using Microsoft.VisualStudio.DebuggerVisualizers;

namespace LinqToExcel
{
    public class ExcelSQL
    {
        public object ExecuteQuery(Expression expression)
        {

            ExpressionToSQL trans = new ExpressionToSQL();
            string query = trans.BuildSQLStatement(expression);
            Type queryType = expression.Type.GetGenericArguments()[0];
            MethodCallExpression call = (MethodCallExpression)expression;
            UnaryExpression unary = (UnaryExpression)call.Arguments[1];


            return Activator.CreateInstance(typeof(List<>).MakeGenericType(expression.Type.GetGenericArguments()[0]));
        }
    }
}
