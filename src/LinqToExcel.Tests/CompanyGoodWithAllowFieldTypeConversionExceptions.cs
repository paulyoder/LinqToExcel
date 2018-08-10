using System.Collections.Generic;
using LinqToExcel.Exceptions;

namespace LinqToExcel.Tests
{
    class CompanyGoodWithAllowFieldTypeConversionExceptions : IAllowFieldTypeConversionExceptions
    {
        public CompanyGoodWithAllowFieldTypeConversionExceptions()
        {
            FieldTypeConversionExceptions = new List<ExcelException>();
        }

        public string Name { get; set; }
        public string CEO { get; set; }
        public int EmployeeCount { get; set; }

        public IList<ExcelException> FieldTypeConversionExceptions { get; private set; }
    }
}
