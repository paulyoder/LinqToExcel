using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;

namespace LinqToExcel
{
    public class ExcelQueryProvider : IQueryProvider
    {
        /// <summary>
        /// Excel File Name
        /// </summary>
        private string _fileName;

        /// <param name="fileName">Excel File Name</param>
        public ExcelQueryProvider(string fileName)
        {
            _fileName = fileName;
        }

        public IQueryable<TElement> CreateQuery<TElement>(Expression expression)
        {
            return new QueryableExcelSheet<TElement>(this, expression);
        }

        public IQueryable CreateQuery(Expression expression)
        {
            return (IQueryable)Activator.CreateInstance(typeof(QueryableExcelSheet<>).MakeGenericType(expression.Type), new object[] { this, expression });
        }

        public object Execute(Expression expression)
        {
            throw new NotImplementedException();
        }

        public TResult Execute<TResult>(Expression expression)
        {
            ExcelSQL repo = new ExcelSQL();
            return (TResult)repo.ExecuteQuery(expression, _fileName);
        }
    }
}
