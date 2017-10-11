using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel.Exceptions
{
    public class ExcelException : Exception
    {
        /// <summary>
        /// Row index where exception occours
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// Column name where exception occours
        /// </summary>
        public string ColumnName { get; set; }

        
        public ExcelException(int row, string columnName, Exception innerException)
            : base(string.Format("Error on row {0} and column name '{1}'.", row, columnName), innerException)
        {
            Row = row;
            ColumnName = columnName;
        }
    }
}
