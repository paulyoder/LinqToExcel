using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel
{
    /// <summary>
    /// Represents a cell and its value in an excel spreadsheet
    /// </summary>
    public class Cell
    {
        /// <summary>
        /// Cell's value
        /// </summary>
        public object Value { get; private set; }

        /// <param name="value">Cell's value</param>
        public Cell(object value)
        {
            Value = value;
        }

        /// <summary>
        /// Returns the cell's value converted as the generic argument type
        /// </summary>
        /// <typeparam name="T">Object type to convert to</typeparam>
        public T ValueAs<T>()
        {
            return (Value == null || Value is DBNull) ?
                default(T) :
                (T)Convert.ChangeType(Value, typeof(T));
        }
    }
}
