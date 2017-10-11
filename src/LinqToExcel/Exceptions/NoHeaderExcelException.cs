using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel.Exceptions
{
    public class NoHeaderExcelException : Exception
    {
        /// <summary>
        /// Row index where exception occours
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// Column index where exception occours
        /// </summary>
        public int Column { get; set; }

        public NoHeaderExcelException(int row, int column, Exception innerException)
            : base(string.Format("Error on row {0} and column {1}.", row, column), innerException)
        {
            Row = row;
            Column = column;
        }
    }
}
