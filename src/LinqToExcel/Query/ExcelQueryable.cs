using System.Collections.Generic;
using System.Linq;
using Remotion.Data.Linq;
using System.Linq.Expressions;

namespace LinqToExcel.Query
{
    public class ExcelQueryable<T> : QueryableBase<T>
    {
        private static IQueryExecutor CreateExecutor(string worksheetName, int? worksheetIndex, string fileName, Dictionary<string, string> columnMappings)
        {
            return new ExcelQueryExecutor(worksheetName, worksheetIndex, fileName, columnMappings);
        }
    
        // This constructor is called by our users, create a new IQueryExecutor.
        public ExcelQueryable(string worksheetName, int? worksheetIndex, string fileName, Dictionary<string, string> columnMappings)
            : base(CreateExecutor(worksheetName, worksheetIndex, fileName, columnMappings))
        { }

        // This constructor is called indirectly by LINQ's query methods, just pass to base.
        public ExcelQueryable(IQueryProvider provider, Expression expression)
            : base(provider, expression)
        { }
    }
}
