using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using System.Collections;

namespace LinqToExcel
{
    public class ExcelRepository<TData> : IQueryable<TData>
    {
        public Expression Expression { get; private set; }
        public IQueryProvider Provider { get; private set; }
        public Type ElementType { get { return typeof(TData); } }

        /// <summary>
        /// This constructor is called by the client to create the data source.
        /// </summary>
        public ExcelRepository()
        {
            this.Provider = new ExcelQueryProvider();
            this.Expression = Expression.Constant(this);
        }

        /// <summary>
        /// This constructor is called by Provider.CreateQuery().
        /// </summary>
        public ExcelRepository(IQueryProvider provider, Expression expression)
        {
            this.Provider = provider;
            this.Expression = expression;
        }

        public IEnumerator<TData> GetEnumerator()
        {
            return (Provider.Execute<IEnumerable<TData>>(this.Expression)).GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return (Provider.Execute<System.Collections.IEnumerable>(this.Expression)).GetEnumerator();
        }
    }
}
