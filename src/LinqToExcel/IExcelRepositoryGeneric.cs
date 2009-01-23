using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;

namespace LinqToExcel
{
    /// <typeparam name="TSheetData">The data type that the sheet contains</typeparam>
    public interface IExcelRepository<SheetDataType>
    {
        /// <summary>
        /// Full path to the Excel document
        /// </summary>
        string FileName { get; set; }

        /// <summary>
        /// Add a property to column name mapping
        /// 
        /// Example
        /// AddMapping(x => x.Name, "FullName")
        /// </summary>
        /// <param name="property">Property to map</param>
        /// <param name="column">Name of the column in the Excel worksheet</param>
        void AddMapping(Expression<Func<SheetDataType, object>> property, string column);
        
        /// <summary>
        /// Worksheet (Sheet1) to perform the Linq query against
        /// </summary>
        IQueryable<SheetDataType> Worksheet();

        /// <summary>
        /// Worksheet to perform the Linq query against
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        IQueryable<SheetDataType> Worksheet(string worksheetName);
    }
}
