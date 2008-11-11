using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;

namespace LinqToExcel
{
    public class ExcelQueryProvider : IQueryProvider
    {
        public IQueryable<TElement> CreateQuery<TElement>(Expression expression)
        {
            return new ExcelRepository<TElement>(this, expression);
        }

        public IQueryable CreateQuery(Expression expression)
        {
            return (IQueryable)Activator.CreateInstance(typeof(ExcelRepository<>).MakeGenericType(expression.Type), new object[] { this, expression });
        }

        public object Execute(Expression expression)
        {
            ExcelSQL repo = new ExcelSQL();
            return repo.ExecuteQuery(expression);
        }

        public TResult Execute<TResult>(Expression expression)
        {
            ExcelSQL repo = new ExcelSQL();
            return (TResult)repo.ExecuteQuery(expression);
        }
    }
}
