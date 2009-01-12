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
        /// Name of the worksheet
        /// 
        /// Default is "Sheet1"
        /// </summary>
        string WorksheetName { get; set; }

        /// <summary>
        /// Worksheet to perform the Linq query against
        /// </summary>
        IQueryable<Row> Worksheet { get; }
    }
}
