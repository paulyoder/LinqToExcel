using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel
{
    public interface IExcelRepository
    {
        /// <summary>
        /// Full path to the Excel document
        /// </summary>
        string FileName { get; set; }

        /// <summary>
        /// Excel document type
        /// </summary>
        ExcelVersion FileType { get; set; }

        /// <summary>
        /// Worksheet (Sheet1) to perform the Linq query against
        /// </summary>
        IQueryable<Row> Worksheet();

        /// <summary>
        /// Worksheet to perform the Linq query against
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        IQueryable<Row> Worksheet(string worksheetName);
    }
}
