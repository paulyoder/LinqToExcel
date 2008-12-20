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

namespace LinqToExcel
{
    internal class ExpressionToSQL : ExpressionVisitor
    {
        private static ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private StringBuilder sb;
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
            sb = new StringBuilder();

            string tableName = (String.IsNullOrEmpty(worksheetName)) ? "Sheet1" : worksheetName;
            sb.Append(string.Format("SELECT * FROM [{0}$]", tableName));
            this.Visit(expression);

            if (_log.IsDebugEnabled)
            {
                _log.Debug("SQL: " + sb.ToString());
                for (int i = 0; i < _params.Count; i++)
                    _log.Debug(string.Format("Param[{0}]: {1}", i, _params[i].Value));
            }

            this.SQLStatement = sb.ToString();
            this.Parameters = _params;
        }

        protected override Expression VisitMethodCall(MethodCallExpression m)
        {
            if (m.Method.Name == "Where")
            {
                sb.Append(" WHERE ");
                this.Visit(m.Arguments[1]);
            }
            else if (m.Method.Name != "Select")
                throw new NotSupportedException(string.Format("{0} method is not supported. Only the 'Where' method call is supported", m.Method.Name));
            return m;
        }

        protected override Expression VisitConstant(ConstantExpression c)
        {
            _params.Add(new OleDbParameter("?", c.Value));
            sb.Append("?");
            return c;
        }

        protected override Expression VisitBinary(BinaryExpression b)
        {
            sb.Append("(");
            
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
                    sb.Append(" AND ");
                    break;
                case ExpressionType.Equal:
                    sb.Append(" = ");
                    break;
                case ExpressionType.GreaterThan:
                    sb.Append(" > ");
                    break;
                case ExpressionType.GreaterThanOrEqual:
                    sb.Append(" >= ");
                    break;
                case ExpressionType.LessThan:
                    sb.Append(" < ");
                    break;
                case ExpressionType.LessThanOrEqual:
                    sb.Append(" <= ");
                    break;
                case ExpressionType.NotEqual:
                    sb.Append(" <> ");
                    break;
                case ExpressionType.OrElse:
                    sb.Append(" OR ");
                    break;
                default:
                    throw new NotSupportedException(string.Format("{0} statement is not supported", b.NodeType.ToString()));
                    break;
            }
            this.Visit(right);
            sb.Append(")");
            return b;
        }

        protected override Expression VisitMemberAccess(MemberExpression m)
        {
            if ((m.Member.MemberType == MemberTypes.Property) &&
                (m.Member.DeclaringType == _sheetDataType))
            {
                //Set the column name to the property mapping if there is one, else use the property name for the column name
                string columnName = (_map.ContainsKey(m.Member.Name)) ? _map[m.Member.Name] : m.Member.Name;
                sb.Append(string.Format("[{0}]", columnName));
            }
            else if ((m.Member.MemberType == MemberTypes.Property) ||
                (m.Member.MemberType == MemberTypes.Field))
            {
                //A field or property on another type has been used as a value in the linq statement
                object value = Expression.Lambda(m).Compile().DynamicInvoke();
                _params.Add(new OleDbParameter("?", value));
                sb.Append("?");
            }
            else
                throw new NotSupportedException(string.Format("{0} member type is not supported. Only fields and properties are supported", m.Member.MemberType.ToString()));
            return m;
        }

        protected override NewExpression VisitNew(NewExpression nex)
        {
            object[] args = GetConstructorArguments(nex.Arguments);
            object newObject = Activator.CreateInstance(nex.Type, args);
            _params.Add(new OleDbParameter("?", newObject));
            sb.Append("?");
            return nex;
        }

        private object[] GetConstructorArguments(ReadOnlyCollection<Expression> constructorArguments)
        {
            List<object> args = new List<object>();
            foreach (Expression exp in constructorArguments)
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
                    throw new NotSupportedException(string.Format("{0} is not supported as a constructor argument", exp.NodeType));
                }
            }
            return args.ToArray();
        }
    }
}
