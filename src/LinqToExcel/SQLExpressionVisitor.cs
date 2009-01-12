using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using System.Reflection;
using log4net;
using System.Data.OleDb;
using System.Collections;
using System.Collections.ObjectModel;
using LinqToExcel.Extensions.Object;

namespace LinqToExcel
{
    internal class SQLExpressionVisitor : ExpressionVisitor
    {
        private static ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private StringBuilder _sql;
        public string SQLStatement { get; private set; }
        public IEnumerable<OleDbParameter> Parameters { get; private set; }
        private List<OleDbParameter> _params;
        private Dictionary<string, string> _map;
        private Type _sheetDataType;

        /// <summary>
        /// Builds the SQL Statement based upon the expression
        /// </summary>
        /// <param name="expression">Expression tree being used</param>
        /// <param name="columnMapping">
        /// Property to column mapping. 
        /// Properties are the dictionary keys and the dictionary values are the corresponding column names.
        /// </param>
        /// <param name="worksheetName">Name of the Excel worksheet</param>
        /// <returns>Returns an SQL statement based upon the expression</returns>
        internal void BuildSQLStatement(Expression expression, Dictionary<string, string> columnMapping, string worksheetName, Type sheetDataType)
        {
            _params = new List<OleDbParameter>();
            _map = columnMapping;
            _sheetDataType = sheetDataType;
            _sql = new StringBuilder();

            string tableName = (String.IsNullOrEmpty(worksheetName)) ? "Sheet1" : worksheetName;
            _sql.Append(string.Format("SELECT * FROM [{0}$]", tableName));
            this.Visit(expression);

            if (_log.IsDebugEnabled)
            {
                _log.Debug("SQL: " + _sql.ToString());
                for (int i = 0; i < _params.Count; i++)
                    _log.Debug(string.Format("Param[{0}]: {1}", i, _params[i].Value));
            }

            this.SQLStatement = _sql.ToString();
            this.Parameters = _params;
        }

        protected override Expression VisitMethodCall(MethodCallExpression m)
        {
            if (m.Method.Name == "Where")
            {
                _sql.Append(" WHERE ");
                this.Visit(m.Arguments[1]);
            }
            else if (IsRowMethodCall(m))
            {
                if (m.Object.As<MethodCallExpression>().Arguments[0].As<ConstantExpression>().Type == typeof(int))
                    throw new ArgumentException("Cannot use column indexes in where clause");

                string columnName = m.Object.As<MethodCallExpression>().Arguments[0].As<ConstantExpression>().Value.ToString();
                _sql.Append(string.Format("[{0}]", columnName));
            }
            else if (m.Method.Name != "Select")
            {
                object methodObject = ((ConstantExpression)m.Object).Value;
                object returnValue = m.Method.Invoke(methodObject, GetMethodArguments(m.Arguments));
                _params.Add(new OleDbParameter("?", returnValue));
                _sql.Append("?");
            }
            return m;
        }

        protected override Expression VisitConstant(ConstantExpression c)
        {
            _params.Add(new OleDbParameter("?", c.Value));
            _sql.Append("?");
            return c;
        }

        protected override Expression VisitBinary(BinaryExpression b)
        {
            _sql.Append("(");
            
            //We always want the MemberAccess (ColumnName) to be on the left side of the statement
            Expression left = b.Left;
            Expression right = b.Right;
            if ((b.Right.NodeType == ExpressionType.MemberAccess) &&
                (((MemberExpression)b.Right).Member.DeclaringType == _sheetDataType))
            {
                left = b.Right;
                right = b.Left;
            }

            this.Visit(left);
            switch (b.NodeType)
            {
                case ExpressionType.AndAlso:
                    _sql.Append(" AND ");
                    break;
                case ExpressionType.Equal:
                    _sql.Append(" = ");
                    break;
                case ExpressionType.GreaterThan:
                    _sql.Append(" > ");
                    break;
                case ExpressionType.GreaterThanOrEqual:
                    _sql.Append(" >= ");
                    break;
                case ExpressionType.LessThan:
                    _sql.Append(" < ");
                    break;
                case ExpressionType.LessThanOrEqual:
                    _sql.Append(" <= ");
                    break;
                case ExpressionType.NotEqual:
                    _sql.Append(" <> ");
                    break;
                case ExpressionType.OrElse:
                    _sql.Append(" OR ");
                    break;
                default:
                    throw new NotSupportedException(string.Format("{0} statement is not supported", b.NodeType.ToString()));
                    break;
            }
            this.Visit(right);
            _sql.Append(")");
            return b;
        }

        protected override Expression VisitMemberAccess(MemberExpression m)
        {
            if ((m.Member.MemberType == MemberTypes.Property) &&
                (m.Member.DeclaringType == _sheetDataType))
            {
                //Set the column name to the property mapping if there is one, else use the property name for the column name
                string columnName = (_map.ContainsKey(m.Member.Name)) ? _map[m.Member.Name] : m.Member.Name;
                _sql.Append(string.Format("[{0}]", columnName));
            }
            else if ((m.Member.MemberType == MemberTypes.Property) ||
                (m.Member.MemberType == MemberTypes.Field))
            {
                //A field or property on another type has been used as a value in the linq statement
                object value = Expression.Lambda(m).Compile().DynamicInvoke();
                _params.Add(new OleDbParameter("?", value));
                _sql.Append("?");
            }
            else
                throw new NotSupportedException(string.Format("{0} member type is not supported. Only fields and properties are supported", m.Member.MemberType.ToString()));
            return m;
        }

        protected override NewExpression VisitNew(NewExpression nex)
        {
            object[] args = GetMethodArguments(nex.Arguments);
            object newObject = Activator.CreateInstance(nex.Type, args);
            _params.Add(new OleDbParameter("?", newObject));
            _sql.Append("?");
            return nex;
        }

        private object[] GetMethodArguments(ReadOnlyCollection<Expression> methodArguments)
        {
            List<object> args = new List<object>();
            foreach (Expression exp in methodArguments)
            {
                if (exp.NodeType == ExpressionType.Constant)
                {
                    args.Add(((ConstantExpression)exp).Value);
                }
                else if (exp.NodeType == ExpressionType.MemberAccess)
                {
                    object value = Expression.Lambda(exp).Compile().DynamicInvoke();
                    args.Add(value);
                }
                else
                {
                    throw new NotSupportedException(string.Format("{0} is not supported as a method argument", exp.NodeType));
                }
            }
            return args.ToArray();
        }

        /// <summary>
        /// Determines if the method call is on a LinqToExcel.Row object
        /// </summary>
        private bool IsRowMethodCall(MethodCallExpression m)
        {
            return ((m.Object is MethodCallExpression) &&
                (m.Object.As<MethodCallExpression>().Object is ParameterExpression) &&
                (m.Object.As<MethodCallExpression>().Object.As<ParameterExpression>().Type == typeof(Row)));
        }
    }
}
