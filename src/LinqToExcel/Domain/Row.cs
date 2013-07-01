﻿using System;
using System.Collections.Generic;
using LinqToExcel.Extensions;

namespace LinqToExcel
{
    public class Row : List<Cell>
    {
        IDictionary<string, int> _columnIndexMapping;

        public Row() : 
            this(new List<Cell>(),new Dictionary<string, int>())
        { }

        /// <param name="cells">Cells contained within the row</param>
        /// <param name="columnIndexMapping">Column name to cell index mapping</param>
        public Row(IList<Cell> cells, IDictionary<string, int> columnIndexMapping)
        {
            for (int i = 0; i < cells.Count; i++)
                this.Insert(i, cells[i]);
            _columnIndexMapping = columnIndexMapping;
        }

        /// <param name="columnName">Column Name</param>
        public Cell this[string columnName]
        {
            get 
            {
                if (!_columnIndexMapping.ContainsKey(columnName))
                    throw new ArgumentException(string.Format("'{0}' column name does not exist. Valid column names are '{1}'", 
                        columnName, string.Join("', '", _columnIndexMapping.Keys.ToArray())));
                return base[_columnIndexMapping[columnName]]; 
            }
        }

        /// <summary>
        /// List of column names in the row object
        /// </summary>
        public IEnumerable<string> ColumnNames
        {
            get { return _columnIndexMapping.Keys; }
        }
    }
}
