using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using System.Reflection;

namespace LinqToExcel
{
    public class ExcelRepository<SheetDataType> : IExcelRepository<SheetDataType>
    {
        public string FileName { get; set; }
        private Dictionary<string, string> _mapping = new Dictionary<string, string>();

        public ExcelRepository()
            : this("")
        { }

        /// <param name="fileName">Full path to Excel file</param>
        public ExcelRepository(string fileName)
        {
            FileName = fileName;
        }

        public void AddMapping(Expression<Func<SheetDataType, object>> property, string column)
        {
            //Get the property name
            LambdaExpression exp = (LambdaExpression)property;
            //exp.Body has 2 possible types
            //If the property type is native, then exp.Body == typeof(MemberExpression)
            //If the property type is not native, then exp.Body == typeof(UnaryExpression) in which 
            //case we can get the MemberExpression from its Operand property
            MemberExpression mExp = (exp.Body.NodeType == ExpressionType.MemberAccess) ?
                (MemberExpression)exp.Body :
                (MemberExpression)((UnaryExpression)exp.Body).Operand;
            string propertyName = mExp.Member.Name;
            
            _mapping[propertyName] = column;
        }

        public IQueryable<SheetDataType> Worksheet()
        {
            return Worksheet("Sheet1");
        }

        public IQueryable<SheetDataType> Worksheet(string worksheetName)
        {
            return new QueryableExcelSheet<SheetDataType>(FileName, _mapping, worksheetName);
        }
    }
}
