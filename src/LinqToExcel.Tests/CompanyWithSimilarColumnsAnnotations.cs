using System;
using LinqToExcel.Attributes;

namespace LinqToExcel.Tests
{
    public class CompanyWithSimilarColumnsAnnotations
    {
        [ExcelColumn(new[] { "Company Title", "Company", "Business Title" })]
        public string Name { get; set; }

        [ExcelColumn(new[] { "Boss", "Head", "Executive" })]
        public string CEO { get; set; }
        
        [ExcelColumn(new[] { "Number People", "People Count" })]
        public int EmployeeCount { get; set; }

        [ExcelColumn(new[] { "Initiation Date", "Init Date" })]
        public DateTime StartDate { get; set; }

        [ExcelColumn(new[] { "Active", "Is Active" })]
        public string IsActive { get; set; }
    }
}
