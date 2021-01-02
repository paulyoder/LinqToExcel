using System;
using LinqToExcel.Attributes;

namespace LinqToExcel.Tests
{
    public class CompanyWithSimilarColumnAnnotations
    {
        [ExcelColumn(new[] { "Company Title", "Company Name" })]
        public string Name { get; set; }

        [ExcelColumn(new[] { "CEO", "Boss" })]
        public string CEO { get; set; }
        
        [ExcelColumn(new[] { "Employee Count", "Number People" })]
        public int EmployeeCount { get; set; }

        [ExcelColumn(new[] { "Initiation Date", "Init Date" })]
        public DateTime StartDate { get; set; }

        [ExcelColumn(new[] { "IsActive", "Active" })]
        public string IsActive { get; set; }
    }
}
