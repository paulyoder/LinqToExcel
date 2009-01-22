using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using LinqToExcel.Extensions.Object;
using System.Collections.ObjectModel;

namespace LinqToExcel.Extensions.Expressions
{
    public static class ExpressionExtensions
    {
        /// <summary>
        /// Returns the argument values
        /// </summary>
        public static object[] GetArgValues(this ReadOnlyCollection<Expression> arguments)
        {
            List<object> argValues = new List<object>();
            foreach (Expression exp in arguments)
            {
                if (exp.NodeType == ExpressionType.Constant)
                    argValues.Add(exp.As<ConstantExpression>().Value);
                else if (exp.NodeType == ExpressionType.MemberAccess)
                    argValues.Add(Expression.Lambda(exp).Compile().DynamicInvoke());
                else if (exp.NodeType == ExpressionType.Call)
                    argValues.Add(exp.As<MethodCallExpression>().Invoke());
                else
                    throw new NotSupportedException(string.Format("{0} is not supported as a method argument", exp.NodeType));
            }
            return argValues.ToArray();
        }

        /// <summary>
        /// Invokes the method
        /// </summary>
        /// <returns>
        /// Returns the method's return value
        /// </returns>
        public static object Invoke(this MethodCallExpression method)
        {
            if (method.Object is ConstantExpression)
                return method.Method.Invoke(method.Object.As<ConstantExpression>().Value, method.Arguments.GetArgValues());
            else if (method.Object is MethodCallExpression)
                return method.Method.Invoke(method.Object.As<MethodCallExpression>().Invoke(), method.Arguments.GetArgValues());
            else
                throw new Exception();
        }
    }
}
