using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using System.Reflection;

namespace LinqToExcel
{
    public class ExcelRepository
    {
        /// <summary>
        /// Excel File Name
        /// </summary>
        public string FileName { get; private set; }

        /// <param name="fileName">Excel File Name</param>
        public ExcelRepository(string fileName)
        {
            this.FileName = fileName;
        }

        /// <summary>
        /// Creates a Linq queryable interface to an Excel sheet
        /// </summary>
        /// <typeparam name="SheetDataType">A Class representing the data in the Excel sheet</typeparam>
        /// <returns>Returns a Linq queryable interface to an Excel sheet</returns>
        public IQueryable<SheetDataType> GetSheet<SheetDataType>()
        {
            return new QueryableExcelSheet<SheetDataType>(this.FileName, new Dictionary<string,string>(), "");
        }

        public static Dictionary<Expression<Func<SheetDataType, object>>, string> CreateColumnMapping<SheetDataType>()
        {
            return new Dictionary<Expression<Func<SheetDataType, object>>, string>();
        }

        /// <summary>
        /// Creates a Linq queryable interface to an Excel sheet
        /// </summary>
        /// <typeparam name="SheetDataType">A Class representing the data in the Excel sheet</typeparam>
        /// <param name="fileName">File path to the Excel workbook</param>
        /// <returns>Returns a Linq queryable interface to an Excel sheet</returns>
        public static IQueryable<SheetDataType> GetSheet<SheetDataType>(string fileName)
        {
            return GetSheet<SheetDataType>(fileName, new Dictionary<Expression<Func<SheetDataType, object>>, string>(), "");
        }

        /// <summary>
        /// Creates a Linq queryable interface to an Excel sheet
        /// </summary>
        /// <typeparam name="SheetDataType">A Class representing the data in the Excel sheet</typeparam>
        /// <param name="fileName">File path to the Excel workbook</param>
        /// <param name="columnMapping">
        /// Property to column mapping. 
        /// Properties are the dictionary keys and the dictionary values are the corresponding column names.
        /// </param>
        /// <param name="worksheetName">Name of the Excel worksheet</param>
        /// <returns>Returns a Linq queryable interface to an Excel sheet</returns>
        public static IQueryable<SheetDataType> GetSheet<SheetDataType>(string fileName, Dictionary<Expression<Func<SheetDataType, object>>, string> columnMapping)
        {
            return GetSheet<SheetDataType>(fileName, columnMapping, "");
        }

        /// <summary>
        /// Creates a Linq queryable interface to an Excel sheet
        /// </summary>
        /// <typeparam name="SheetDataType">A Class representing the data in the Excel sheet</typeparam>
        /// <param name="fileName">File path to the Excel workbook</param>
        /// <param name="worksheetName">Name of the Excel worksheet</param>
        /// <returns>Returns a Linq queryable interface to an Excel sheet</returns>
        public static IQueryable<SheetDataType> GetSheet<SheetDataType>(string fileName, string worksheetName)
        {
            return GetSheet<SheetDataType>(fileName, new Dictionary<Expression<Func<SheetDataType, object>>, string>(), worksheetName);
        }

        /// <summary>
        /// Creates a Linq queryable interface to an Excel sheet
        /// </summary>
        /// <typeparam name="SheetDataType">A Class representing the data in the Excel sheet</typeparam>
        /// <param name="fileName">File path to the Excel workbook</param>
        /// <param name="columnMapping">
        /// Property to column mapping. 
        /// Properties are the dictionary keys and the dictionary values are the corresponding column names.
        /// </param>
        /// <param name="worksheetName">Name of the Excel worksheet</param>
        /// <returns>Returns a Linq queryable interface to an Excel sheet</returns>
        public static IQueryable<SheetDataType> GetSheet<SheetDataType>(string fileName, Dictionary<Expression<Func<SheetDataType, object>>, string> columnMapping, string worksheetName)
        {
            //Getting the name of the properties in the mapping
            Dictionary<string, string> stringColumnMapping = new Dictionary<string, string>();
            foreach (KeyValuePair<Expression<Func<SheetDataType, object>>, string> kvp in columnMapping)
            {
                LambdaExpression exp = (LambdaExpression)kvp.Key;
                //exp.Body is of 2 possible types
                //If the property type is native, then exp.Body == typeof(MemberExpression)
                //If the property type is not native, then exp.Body == typeof(UnaryExpression) in which 
                //case we can get the MemberExpression from the Operand property
                MemberExpression mExp = (exp.Body.NodeType == ExpressionType.MemberAccess) ?
                    (MemberExpression)exp.Body :
                    (MemberExpression)((UnaryExpression)exp.Body).Operand;
                string propertyName = mExp.Member.Name;

                stringColumnMapping[propertyName] = kvp.Value;
            }

            return new QueryableExcelSheet<SheetDataType>(fileName, stringColumnMapping, worksheetName);
        }
    }
}
