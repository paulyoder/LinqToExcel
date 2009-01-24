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
        private ExcelVersion _fileType;
        private Dictionary<string, string> _columnMappings;
        private string _worksheetName;
        private Type _dataType;

        /// <param name="fileName">Excel File Name</param>
        /// <param name="columnMappings">
        /// Property to column mapping. 
        /// Properties are the dictionary keys and the dictionary values are the corresponding column names.
        /// </param>
        /// <param name="worksheetName">Name of the Excel worksheet</param>
        public ExcelQueryProvider(string fileName, ExcelVersion fileType, Type dataType, Dictionary<string, string> columnMappings, string worksheetName)
        {
            _fileName = fileName;
            _fileType = fileType;
            _dataType = dataType;
            _columnMappings = columnMappings;
            _worksheetName = worksheetName;
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
            ExcelOLEDB repo = new ExcelOLEDB();
            return (TResult)repo.ExecuteQuery(expression, _dataType, _fileName, _fileType, _columnMappings, _worksheetName);
        }
    }
}
