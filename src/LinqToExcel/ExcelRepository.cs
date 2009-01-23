using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel
{
    public class ExcelRepository : ExcelRepository<Row>, IExcelRepository
    {
        public ExcelRepository()
            : base("")
        { }

        /// <param name="fileName">Full path to Excel file</param>
        public ExcelRepository(string fileName)
            : base(fileName)
        { }
    }
}
