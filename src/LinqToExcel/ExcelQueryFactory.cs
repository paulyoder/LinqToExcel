using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using System.Threading;
using LinqToExcel.Query;

namespace LinqToExcel
{
    public class ExcelQueryFactory : IExcelQueryFactory
    {
        public string FileName { get; set; }
        private readonly Dictionary<string, string> _mapping = new Dictionary<string, string>();

        public ExcelQueryFactory()
            : this(null)
        { }

        public ExcelQueryFactory(string fileName)
        {
            FileName = fileName;
        }

        public ExcelQueryable<TSheetData> Worksheet<TSheetData>()
        {
            return Worksheet<TSheetData>("Sheet1", FileName, _mapping);
        }

        public ExcelQueryable<Row> Worksheet()
        {
            return Worksheet<Row>("Sheet1", FileName, _mapping);
        }

        public ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName)
        {
            return Worksheet<TSheetData>(worksheetName, FileName, _mapping);
        }

        public ExcelQueryable<Row> Worksheet(string worksheetName)
        {
            return Worksheet<Row>(worksheetName, FileName, _mapping);
        }

        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName, string fileName)
        {
            return Worksheet<TSheetData>(worksheetName, fileName, new Dictionary<string, string>());
        }

        public static ExcelQueryable<Row> Worksheet(string worksheetName, string fileName)
        {
            return Worksheet<Row>(worksheetName, fileName, new Dictionary<string, string>());
        }

        public static ExcelQueryable<Row> Worksheet(string worksheetName, string fileName, Dictionary<string, string> mapping)
        {
            return Worksheet<Row>(worksheetName, fileName, mapping);
        }

        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName, string fileName, Dictionary<string, string> mapping)
        {
            if (fileName == null)
                throw new ArgumentNullException("fileName cannot be null");
            if (string.IsNullOrEmpty(worksheetName))
                worksheetName = "Sheet1";
            mapping = mapping ?? new Dictionary<string, string>();

            return new ExcelQueryable<TSheetData>(worksheetName, fileName, mapping);
        }

        public void AddMapping<TSheetData>(Expression<Func<TSheetData, object>> property, string column)
        {
            //Get the property name
            var exp = (LambdaExpression)property;
            //exp.Body has 2 possible types
            //If the property type is native, then exp.Body == typeof(MemberExpression)
            //If the property type is not native, then exp.Body == typeof(UnaryExpression) in which 
            //case we can get the MemberExpression from its Operand property
            var mExp = (exp.Body.NodeType == ExpressionType.MemberAccess) ?
                (MemberExpression)exp.Body :
                (MemberExpression)((UnaryExpression)exp.Body).Operand;
            var propertyName = mExp.Member.Name;

            _mapping[propertyName] = column;
        }
    }
}
