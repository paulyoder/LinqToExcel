using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using System.Collections;

namespace LinqToExcel
{
    public class QueryableExcelSheet<SheetDataType> : IQueryable<SheetDataType>
    {
        public Expression Expression { get; private set; }
        public IQueryProvider Provider { get; private set; }
        public Type ElementType { get { return typeof(SheetDataType); } }

        /// <summary>
        /// This constructor is called by the client to create the data source.
        /// </summary>
        public QueryableExcelSheet(string fileName)
        {
            this.Provider = new ExcelQueryProvider(fileName);
            this.Expression = Expression.Constant(this);
        }

        /// <summary>
        /// This constructor is called by Provider.CreateQuery().
        /// </summary>
        public QueryableExcelSheet(IQueryProvider provider, Expression expression)
        {
            this.Provider = provider;
            this.Expression = expression;
        }

        public IEnumerator<SheetDataType> GetEnumerator()
        {
            return (Provider.Execute<IEnumerable<SheetDataType>>(this.Expression)).GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return (Provider.Execute<System.Collections.IEnumerable>(this.Expression)).GetEnumerator();
        }
    }
}
