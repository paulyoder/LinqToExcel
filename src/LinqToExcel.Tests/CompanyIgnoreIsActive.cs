using LinqToExcel.Attributes;

namespace LinqToExcel.Tests
{
    class CompanyIgnoreIsActive : Company
    {
        [ExcelIgnore]
        new public bool IsActive { get; set; }
    }
}
