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
        public int Row { get; private set; }

        /// <summary>
        /// Column name where exception occours
        /// </summary>
        public string ColumnName { get; private set; }

        /// <summary>
        /// Column index where exception occours
        /// </summary>
        public int Column { get; private set; }

        public ExcelException(int row, string columnName, Exception innerException)
            : base(string.Format("Error on row {0} and column name '{1}'.", row, columnName), innerException)
        {
            Row = row;
            ColumnName = columnName;
        }

        public ExcelException(int row, int column, Exception innerException)
            : base(string.Format("Error on row {0} and column '{1}'.", row, Query.ExcelUtilities.ColumnIndexToExcelColumnName(column)), innerException)
        {
            Row = row;
            Column = column;
            ColumnName = Query.ExcelUtilities.ColumnIndexToExcelColumnName(column);
        }

        public ExcelException(int row, int column, string columnName, Exception innerException)
            : base(string.Format("Error on row {0} and column name '{1}'.", row, columnName), innerException)
        {
            Row = row;
            Column = column;
            ColumnName = columnName ?? Query.ExcelUtilities.ColumnIndexToExcelColumnName(column);
        }
    }
}
