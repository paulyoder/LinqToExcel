using System;
using System.Collections.Generic;
using System.Linq.Expressions;
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
            return Worksheet<TSheetData>("Sheet1", null, FileName, _mapping);
        }

        public ExcelQueryable<Row> Worksheet()
        {
            return Worksheet<Row>("Sheet1", null, FileName, _mapping);
        }

        public ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName)
        {
            return Worksheet<TSheetData>(worksheetName, null, FileName, _mapping);
        }

        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<TSheetData> Worksheet<TSheetData>(int worksheetIndex)
        {
            return Worksheet<TSheetData>(null, worksheetIndex, FileName, _mapping);
        }

        public ExcelQueryable<Row> Worksheet(string worksheetName)
        {
            return Worksheet<Row>(worksheetName, null, FileName, _mapping);
        }

        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public ExcelQueryable<Row> Worksheet(int worksheetIndex)
        {
            return Worksheet<Row>(null, worksheetIndex, FileName, _mapping);
        }

        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName, string fileName)
        {
            return Worksheet<TSheetData>(worksheetName, null, fileName, new Dictionary<string, string>());
        }

        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(int worksheetIndex, string fileName)
        {
            return Worksheet<TSheetData>(null, worksheetIndex, fileName, new Dictionary<string, string>());
        }

        public static ExcelQueryable<Row> Worksheet(string worksheetName, string fileName)
        {
            return Worksheet<Row>(worksheetName, null, fileName, new Dictionary<string, string>());
        }

        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public static ExcelQueryable<Row> Worksheet(int worksheetIndex, string fileName)
        {
            return Worksheet<Row>(null, worksheetIndex, fileName, new Dictionary<string, string>());
        }

        public static ExcelQueryable<Row> Worksheet(string worksheetName, string fileName, Dictionary<string, string> mapping)
        {
            return Worksheet<Row>(worksheetName, null, fileName, mapping);
        }

        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public static ExcelQueryable<Row> Worksheet(int worksheetIndex, string fileName, Dictionary<string, string> mapping)
        {
            return Worksheet<Row>(null, worksheetIndex, fileName, mapping);
        }

        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName, string fileName, Dictionary<string, string> mapping)
        {
            return Worksheet<TSheetData>(worksheetName, null, fileName, mapping);
        }

        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        public static ExcelQueryable<TSheetData> Worksheet<TSheetData>(int worksheetIndex, string fileName, Dictionary<string, string> mapping)
        {
            return Worksheet<TSheetData>(null, worksheetIndex, fileName, mapping);
        }

        /// <param name="worksheetIndex">Worksheet index ordered by name, not position in the workbook</param>
        private static ExcelQueryable<TSheetData> Worksheet<TSheetData>(string worksheetName, int? worksheetIndex, string fileName, Dictionary<string, string> mapping)
        {
            if (fileName == null)
                throw new ArgumentNullException("FileName", "FileName property cannot be null.");
            
            mapping = mapping ?? new Dictionary<string, string>();
            return new ExcelQueryable<TSheetData>(worksheetName, worksheetIndex, fileName, mapping);
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
