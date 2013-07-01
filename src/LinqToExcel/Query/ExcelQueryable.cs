#region

using System.Linq;
using System.Linq.Expressions;
using Remotion.Linq;
using Remotion.Linq.Parsing.Structure;

#endregion

namespace LinqToExcel.Query
{
    public class ExcelQueryable<T> : QueryableBase<T>
    {
        // This constructor is called by users, create a new IQueryExecutor.
        internal ExcelQueryable(ExcelQueryArgs args)
            : base(QueryParser.CreateDefault(), CreateExecutor(args))
        {
        }

        // This constructor is called indirectly by LINQ's query methods, just pass to base.
        public ExcelQueryable(IQueryProvider provider, Expression expression)
            : base(provider, expression)
        {
        }

        private static IQueryExecutor CreateExecutor(ExcelQueryArgs args)
        {
            return new ExcelQueryExecutor(args);
        }
    }
}