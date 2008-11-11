using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using YoderSolutions.Libs.ExpressionTree;
using System.Reflection;
using log4net;

namespace LinqToExcel
{
    public class ExpressionToSQL : ExpressionVisitor
    {
        private static ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private StringBuilder sb;

        /// <summary>
        /// Builds the SQL Statement based upon the expression
        /// </summary>
        /// <param name="expression">Expression tree being used</param>
        /// <returns>Returns an SQL statement based upon the expression</returns>
        public string BuildSQLStatement(Expression expression)
        {
            sb = new StringBuilder();
            sb.Append("SELECT * FROM [Sheet1$]");

            this.Visit(expression);

            if (_log.IsDebugEnabled)
                _log.Debug("SQL: " + sb.ToString());
            
            return sb.ToString();
        }

        protected override Expression VisitMethodCall(MethodCallExpression m)
        {
            if (m.Method.Name == "Where")
            {
                sb.Append(" Where ");
                this.Visit(m.Arguments[1]);
            }
            else if (m.Method.Name != "Select")
                throw new NotSupportedException(string.Format("{0} method is not supported. Only the 'Where' method call is supported", m.Method.Name));
            return m;
        }

        protected override Expression VisitConstant(ConstantExpression c)
        {
            if (c.Value.GetType() == typeof(int))
                sb.Append(c.Value);
            else
                sb.Append(string.Format("'{0}'",c.Value));
            return c;
        }

        protected override Expression VisitBinary(BinaryExpression b)
        {
            sb.Append("(");
            this.Visit(b.Left);
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
            this.Visit(b.Right);
            sb.Append(")");
            return b;
        }

        protected override Expression VisitMemberAccess(MemberExpression m)
        {
            if (m.Member.MemberType == MemberTypes.Property)
                sb.Append(string.Format("{0}", m.Member.Name));
            else if (m.Member.MemberType == MemberTypes.Field)
                this.VisitConstant((ConstantExpression)m.Expression);
            else
                throw new NotSupportedException(string.Format("{0} member type is not supported. Only fields and properties are supported", m.Member.MemberType.ToString()));
            return m;
        }
    }
}
