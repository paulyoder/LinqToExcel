using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

        /// <summary>
        /// Creates a Linq queryable interface to an Excel sheet
        /// </summary>
        /// <typeparam name="SheetDataType">A Class representing the data in the Excel sheet</typeparam>
        /// <param name="fileName">File path to the Excel workbook</param>
        /// <returns>Returns a Linq queryable interface to an Excel sheet</returns>
        public static IQueryable<SheetDataType> GetSheet<SheetDataType>(string fileName)
        {
            return GetSheet<SheetDataType>(fileName, new Dictionary<string, string>(), "");
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
        /// <returns>Returns a Linq queryable interface to an Excel sheet</returns>
        public static IQueryable<SheetDataType> GetSheet<SheetDataType>(string fileName, Dictionary<string, string> columnMapping)
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
            return GetSheet<SheetDataType>(fileName, new Dictionary<string, string>(), worksheetName);
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
        public static IQueryable<SheetDataType> GetSheet<SheetDataType>(string fileName, Dictionary<string, string> columnMapping, string worksheetName)
        {
            return new QueryableExcelSheet<SheetDataType>(fileName, columnMapping, worksheetName);
        }
    }
}
