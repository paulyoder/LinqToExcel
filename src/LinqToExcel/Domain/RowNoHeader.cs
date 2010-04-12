using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel
{
    public class RowNoHeader : List<Cell>
    {
        /// <param name="cells">Cells contained within the row</param>
        public RowNoHeader(IEnumerable<Cell> cells)
        {
            base.AddRange(cells);
        }
    }
}
