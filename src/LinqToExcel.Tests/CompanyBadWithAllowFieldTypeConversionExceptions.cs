using System;
using System.Collections.Generic;
using LinqToExcel.Exceptions;

namespace LinqToExcel.Tests
{
    class CompanyBadWithAllowFieldTypeConversionExceptions : IAllowFieldTypeConversionExceptions
    {
        public CompanyBadWithAllowFieldTypeConversionExceptions()
        {
            FieldTypeConversionExceptions = new List<ExcelException>();
        }

        public string Name { get; set; }

        /// <summary>
        /// The CEO column is actually a string and should fail type conversion to int.
        /// </summary>
        public int CEO { get; set; }

        /// <summary>
        /// The EmployeeCount column is actually a string and should fail type conversion to DateTime.
        /// </summary>
        public DateTime EmployeeCount { get; set; }

        public IList<ExcelException> FieldTypeConversionExceptions { get; private set; }
    }
}
