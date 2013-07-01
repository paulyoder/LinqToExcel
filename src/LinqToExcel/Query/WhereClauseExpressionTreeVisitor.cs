﻿#region

using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using LinqToExcel.Extensions;
using Remotion.Linq.Parsing;

#endregion

namespace LinqToExcel.Query
{
    public class WhereClauseExpressionTreeVisitor : ThrowingExpressionTreeVisitor
    {
        private readonly Dictionary<string, string> _columnMapping;
        private readonly List<string> _columnNamesUsed = new List<string>();
        private readonly List<OleDbParameter> _params = new List<OleDbParameter>();
        private readonly Type _sheetType;
        private readonly List<string> _validStringMethods;
        private readonly StringBuilder _whereClause = new StringBuilder();

        public WhereClauseExpressionTreeVisitor(Type sheetType, Dictionary<string, string> columnMapping)
        {
            _sheetType = sheetType;
            _columnMapping = columnMapping;
            _validStringMethods = new List<string>
                {
                    "Equals",
                    "Contains",
                    "StartsWith",
                    "IsNullOrEmpty",
                    "EndsWith"
                };
        }

        public string WhereClause
        {
            get { return _whereClause.ToString(); }
        }

        public IEnumerable<OleDbParameter> Params
        {
            get { return _params; }
        }

        public IEnumerable<string> ColumnNamesUsed
        {
            get { return _columnNamesUsed.Select(x => x.Replace("[", "").Replace("]", "")); }
        }

        public void Visit(Expression expression)
        {
            base.VisitExpression(expression);
        }

        protected override Exception CreateUnhandledItemException<T>(T unhandledItem, string visitMethod)
        {
            throw new NotImplementedException(visitMethod + " method is not implemented");
        }

        protected override Expression VisitBinaryExpression(BinaryExpression bExp)
        {
            _whereClause.Append("(");

            // Patch for vb.net expression that are always considered a MethodCallExpression even if they are not.
            // see http://www.re-motion.org/blogs/mix/archive/2009/10/16/vb.net-specific-text-comparison-in-linq-queries.aspx
            bExp = ConvertVbStringCompare(bExp);

            //We always want the MemberAccess (ColumnName) to be on the left side of the statement
            var bLeft = bExp.Left;
            var bRight = bExp.Right;
            if ((bExp.Right.NodeType == ExpressionType.MemberAccess) &&
                (((MemberExpression) bExp.Right).Member.DeclaringType == _sheetType))
            {
                bLeft = bExp.Right;
                bRight = bExp.Left;
            }

            VisitExpression(bLeft);
            switch (bExp.NodeType)
            {
                case ExpressionType.AndAlso:
                    _whereClause.Append(" AND ");
                    break;
                case ExpressionType.Equal:
                    if (bRight.IsNullValue())
                        _whereClause.Append(" IS NULL");
                    else
                        _whereClause.Append(" = ");
                    break;
                case ExpressionType.GreaterThan:
                    _whereClause.Append(" > ");
                    break;
                case ExpressionType.GreaterThanOrEqual:
                    _whereClause.Append(" >= ");
                    break;
                case ExpressionType.LessThan:
                    _whereClause.Append(" < ");
                    break;
                case ExpressionType.LessThanOrEqual:
                    _whereClause.Append(" <= ");
                    break;
                case ExpressionType.NotEqual:
                    if (bRight.IsNullValue())
                        _whereClause.Append(" IS NOT NULL");
                    else
                        _whereClause.Append(" <> ");
                    break;
                case ExpressionType.OrElse:
                    _whereClause.Append(" OR ");
                    break;
                default:
                    throw new NotSupportedException(string.Format("{0} statement is not supported", bExp.NodeType.ToString()));
            }
            if (!bRight.IsNullValue())
                VisitExpression(bRight);
            _whereClause.Append(")");
            return bExp;
        }

        protected BinaryExpression ConvertVbStringCompare(BinaryExpression exp)
        {
            if (exp.Left.NodeType == ExpressionType.Call)
            {
                var compareStringCall = (MethodCallExpression) exp.Left;
                if (compareStringCall.Method.DeclaringType.FullName == "Microsoft.VisualBasic.CompilerServices.Operators" && compareStringCall.Method.Name == "CompareString")
                {
                    var arg1 = compareStringCall.Arguments[0];
                    var arg2 = compareStringCall.Arguments[1];

                    switch (exp.NodeType)
                    {
                        case ExpressionType.LessThan:
                            return Expression.LessThan(arg1, arg2);
                        case ExpressionType.LessThanOrEqual:
                            return Expression.LessThanOrEqual(arg1, arg2);
                        case ExpressionType.GreaterThan:
                            return Expression.GreaterThan(arg1, arg2);
                        case ExpressionType.GreaterThanOrEqual:
                            return Expression.GreaterThanOrEqual(arg1, arg2);
                        case ExpressionType.NotEqual:
                            return Expression.NotEqual(arg1, arg2);
                        default:
                            return Expression.Equal(arg1, arg2);
                    }
                }
            }
            return exp;
        }

        protected override Expression VisitMemberExpression(MemberExpression mExp)
        {
            //Set the column name to the property mapping if there is one, 
            //else use the property name for the column name
            var columnName = (_columnMapping.ContainsKey(mExp.Member.Name)) ?
                                 _columnMapping[mExp.Member.Name] :
                                 mExp.Member.Name;
            _whereClause.AppendFormat("[{0}]", columnName);
            _columnNamesUsed.Add(columnName);
            return mExp;
        }

        protected override Expression VisitConstantExpression(ConstantExpression cExp)
        {
            _params.Add(new OleDbParameter("?", cExp.Value));
            _whereClause.Append("?");
            return cExp;
        }

        /// <summary>
        ///     This method is visited when the LinqToExcel.Row type is used in the query
        /// </summary>
        protected override Expression VisitUnaryExpression(UnaryExpression uExp)
        {
            if (IsNotStringIsNullOrEmpty(uExp))
                AddStringIsNullOrEmptyToWhereClause((MethodCallExpression) uExp.Operand, true);
            else
                _whereClause.Append(GetColumnName(uExp.Operand));
            return uExp;
        }

        private bool IsNotStringIsNullOrEmpty(UnaryExpression uExp)
        {
            return uExp.NodeType == ExpressionType.Not && ((MethodCallExpression) uExp.Operand).Method.Name == "IsNullOrEmpty";
        }

        /// <summary>
        ///     Only As<>() method calls on the LinqToExcel.Row type are support
        /// </summary>
        protected override Expression VisitMethodCallExpression(MethodCallExpression mExp)
        {
            if (_validStringMethods.Contains(mExp.Method.Name))
                ProcessStringMethod(mExp);
            else
            {
                var columnName = GetColumnName(mExp);
                _whereClause.Append(columnName);
                _columnNamesUsed.Add(columnName);
            }
            return mExp;
        }

        private void ProcessStringMethod(MethodCallExpression mExp)
        {
            switch (mExp.Method.Name)
            {
                case "Contains":
                    AddStringMethodToWhereClause(mExp, "LIKE", "%{0}%");
                    break;
                case "StartsWith":
                    AddStringMethodToWhereClause(mExp, "LIKE", "{0}%");
                    break;
                case "EndsWith":
                    AddStringMethodToWhereClause(mExp, "LIKE", "%{0}");
                    break;
                case "Equals":
                    AddStringMethodToWhereClause(mExp, "=", "{0}");
                    break;
                case "IsNullOrEmpty":
                    AddStringIsNullOrEmptyToWhereClause(mExp);
                    break;
            }
        }

        private void AddStringMethodToWhereClause(MethodCallExpression mExp, string operatorString, string argumentFormat)
        {
            _whereClause.Append("(");
            VisitExpression(mExp.Object);
            _whereClause.AppendFormat(" {0} ?)", operatorString);

            var value = mExp.Arguments.First().ToString().Replace("\"", "");
            var parameter = string.Format(argumentFormat, value);
            _params.Add(new OleDbParameter("?", parameter));
        }

        private void AddStringIsNullOrEmptyToWhereClause(MethodCallExpression mExp)
        {
            AddStringIsNullOrEmptyToWhereClause(mExp, false);
        }

        private void AddStringIsNullOrEmptyToWhereClause(MethodCallExpression mExp, bool notEqual)
        {
            var columnName = GetColumnName((MemberExpression) mExp.Arguments[0]);
            if (notEqual)
                _whereClause.AppendFormat("(({0} <> '') OR ({0} IS NOT NULL))", columnName);
            else
                _whereClause.AppendFormat("(({0} = '') OR ({0} IS NULL))", columnName);
        }

        /// <summary>
        ///     Retrieves the column name from a member expression or method call expression
        /// </summary>
        /// <param name="exp">Expression</param>
        private string GetColumnName(Expression exp)
        {
            if (exp is MemberExpression)
                return GetColumnName((MemberExpression) exp);
            else
                return GetColumnName((MethodCallExpression) exp);
        }

        /// <summary>
        ///     Retrieves the column name from a member expression
        /// </summary>
        /// <param name="mExp">Member Expression</param>
        private string GetColumnName(MemberExpression mExp)
        {
            return "[" + mExp.Member.Name + "]";
        }

        /// <summary>
        ///     Retrieves the column name from a method call expression
        /// </summary>
        /// <param name="exp">Method Call Expression</param>
        private string GetColumnName(MethodCallExpression mExp)
        {
            MethodCallExpression method = mExp;
            if (mExp.Object is MethodCallExpression)
                method = (MethodCallExpression) mExp.Object;

            var arg = method.Arguments.First();
            if (arg.Type == typeof (int))
            {
                if (_sheetType == typeof (RowNoHeader))
                    return string.Format("F{0}", Int32.Parse(arg.ToString()) + 1);
                else
                    throw new ArgumentException("Can only use column indexes in WHERE clause when using WorksheetNoHeader");
            }

            var columnName = arg.ToString().ToCharArray();
            columnName[0] = "[".ToCharArray().First();
            columnName[columnName.Length - 1] = "]".ToCharArray().First();
            return new string(columnName);
        }
    }
}