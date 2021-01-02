using System.Collections.Generic;
using System.Linq;
using Remotion.Linq;
using System.Linq.Expressions;
using LinqToExcel.Attributes;
using System;

using LinqToExcel.Logging;

using Remotion.Linq.Parsing.Structure;

namespace LinqToExcel.Query
{
    public class ExcelQueryable<T> : QueryableBase<T>
    {
        private static IQueryExecutor CreateExecutor(ExcelQueryArgs args, ILogManagerFactory logManagerFactory)
        {
            return new ExcelQueryExecutor(args, logManagerFactory);
        }

        // This constructor is called by users, create a new IQueryExecutor.
        internal ExcelQueryable(ExcelQueryArgs args, ILogManagerFactory logManagerFactory)
            : base( QueryParser.CreateDefault(), CreateExecutor(args, logManagerFactory) )
        {
            foreach (var property in typeof(T).GetProperties())
            {
                ExcelColumnAttribute att = (ExcelColumnAttribute)Attribute.GetCustomAttribute(property, typeof(ExcelColumnAttribute));
                if (att != null && !args.ColumnMappings.ContainsKey(property.Name))
                {
                    var columnNames = ExcelUtilities.GetColumnNames(args);

                    // Two possible thinks happen here
                    // If the ColumnName exists in worksheet, add mapping for it, otherwise
                    // If have column with similar name in columns list, return it, otherwise force value to ColumnName
                    // So its possible to read imprevisible files with similar column names
                    // Note that the order of columns are important, because return first occurrence of the similar column name
                    var columnName = columnNames.Contains(att.ColumnName) ?
                        att.ColumnName :
                        columnNames.ToList().Find(n => att.HaveSimilarWith(n))??att.ColumnName;

                    // Avoid empty column names
                    if (!string.IsNullOrEmpty(columnName))
                    {
                        args.ColumnMappings.Add(property.Name, columnName);
                    }
                }
            }
        }

        // This constructor is called indirectly by LINQ's query methods, just pass to base.
        public ExcelQueryable(IQueryProvider provider, Expression expression)
            : base(provider, expression)
        { }
    }
}
