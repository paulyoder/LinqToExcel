using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;
using System.Linq.Expressions;
using System.Collections.ObjectModel;
using LinqToExcel.Extensions.Expressions;
using System.Reflection;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Unit")]
    [TestsOn(typeof(ExpressionExtensions))]
    [TestFixture]
    public class ExpressionExtensionsTests
    {
        private ReadOnlyCollection<Expression> CreateArgs(params Expression[] args)
        {
            return new ReadOnlyCollection<Expression>(args.ToList<Expression>());
        }

        [Test]
        public void GetArgValues_returns_value_of_argument_of_type_ConstantExpression()
        {
            var value = 25;
            Expression exp = Expression.Constant(value);
            var args = CreateArgs(exp);
            Assert.AreEqual(value, args.GetArgValues().First());
        }

        [Test]
        public void GetArgValues_returns_value_of_argument_of_type_MemberAccess()
        {
            var animal = new { Name = "Horse" };
            MemberExpression exp = Expression.Property(Expression.Constant(animal), "Name");
            var args = CreateArgs(exp);
            Assert.AreEqual("Horse", args.GetArgValues().First());
        }

        [Test]
        public void GetArgValues_returns_value_of_argument_of_type_MethodCallExpression()
        {
            string city = "Omaha";
            MethodInfo toUpperMethodInfo = city.GetType().GetMethod("ToUpper", new Type[] { });
            MethodCallExpression exp = Expression.Call(Expression.Constant(city), toUpperMethodInfo);
            var args = CreateArgs(exp);
            Assert.AreEqual("OMAHA", args.GetArgValues().First());
        }

        [Test]
        public void Invoke_gets_return_value_for_method_object_of_type_ConstantExpression()
        {
            //testing city.ToUpper()
            string city = "Omaha";
            MethodInfo toUpperMethodInfo = city.GetType().GetMethod("ToUpper", new Type[] { });
            MethodCallExpression exp = Expression.Call(Expression.Constant(city), toUpperMethodInfo);
            Assert.AreEqual("OMAHA", exp.Invoke());
        }

        [Test]
        public void Invoke_gets_return_value_for_method_object_of_type_MethodCallExpression()
        {
            //testing city.ToUpper().IndexOf("H")
            string city = "Omaha";
            MethodInfo toUpperMethodInfo = typeof(string).GetMethod("ToUpper", new Type[] { });
            MethodCallExpression toUpperExp = Expression.Call(Expression.Constant(city), toUpperMethodInfo);
            MethodInfo indexOfMethodInfo = typeof(string).GetMethod("IndexOf", new Type[] { typeof(string) });
            MethodCallExpression indexOfExp = Expression.Call(toUpperExp, indexOfMethodInfo, Expression.Constant("H"));
            Assert.AreEqual(3, indexOfExp.Invoke());
        }
    }
}
