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
            return new QueryableExcelSheet<SheetDataType>(this.FileName);
        }

        /// <summary>
        /// Creates a Linq queryable interface to an Excel sheet
        /// </summary>
        /// <typeparam name="SheetDataType">A Class representing the data in the Excel sheet</typeparam>
        /// <returns>Returns a Linq queryable interface to an Excel sheet</returns>
        public static IQueryable<SheetDataType> GetSheet<SheetDataType>(string fileName)
        {
            return new QueryableExcelSheet<SheetDataType>(fileName);
        }
    }
}
