using System;
using System.Linq;
using MbUnit.Framework;
using System.IO;
using System.Data;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Integration")]
    [TestFixture]
    public class InvalidColumnNamesUsed
    {
        private string _excelFileName;

        [TestFixtureSetUp]
        public void fs()
        {
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
        }

        [Test]
        [ExpectedException(typeof(DataException), "'Bad Column' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'")]
        public void row_column_in_where_clause()
        {
            var list = (from x in ExcelQueryFactory.Worksheet("Sheet1", _excelFileName)
                        where x["Bad Column"].ToString() == "nothing"
                        select x).ToList();
        }

        [Test]
        [ExpectedException(typeof(DataException), "'Bad Column' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'")]
        public void row_column_in_orderby_clause()
        {
            var list = (from x in ExcelQueryFactory.Worksheet("Sheet1", _excelFileName)
                        select x)
                        .OrderBy(x => x["Bad Column"])
                        .ToList();
        }

        [Test]
        [ExpectedException(typeof(DataException), "'City' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'")]
        public void bad_column_in_where_clause()
        {
            var list = (from x in ExcelQueryFactory.Worksheet<CompanyWithCity>("Sheet1", _excelFileName)
                        where x.City == "Omaha"
                        select x).ToList();
        }

        [Test]
        [ExpectedException(typeof(DataException), "'Town' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'")]
        public void bad_column_mapping_in_where_clause()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.AddMapping<CompanyWithCity>(x => x.City, "Town");
            var list = (from x in excel.Worksheet<CompanyWithCity>("Sheet1")
                        where x.City == "Omaha"
                        select x).ToList();
        }

        [Test]
        [ExpectedException(typeof(DataException), "'City' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'")]
        public void bad_column_in_orderby_clause()
        {
            var list = (from x in ExcelQueryFactory.Worksheet<CompanyWithCity>("Sheet1", _excelFileName)
                        select x)
                        .OrderBy(x => x.City)
                        .ToList();
        }

        [Test]
        [ExpectedException(typeof(DataException), "'Town' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'")]
        public void bad_column_mapping_in_orderby_clause()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.AddMapping<CompanyWithCity>(x => x.City, "Town");
            var list = (from x in excel.Worksheet<CompanyWithCity>("Sheet1")
                        select x)
                        .OrderBy(x => x.City)
                        .ToList();
        }

        [Test]
        [ExpectedException(typeof(DataException), "'Employees' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'")]
        public void bad_column_in_average_aggregate()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.AddMapping<CompanyWithCity>(x => x.EmployeeCount, "Employees");
            var list = (from x in excel.Worksheet<CompanyWithCity>("Sheet1")
                        select x)
                        .Average(x => x.EmployeeCount);
        }

        [Test]
        [ExpectedException(typeof(DataException), "'Employees' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'")]
        public void bad_column_in_max_aggregate()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.AddMapping<CompanyWithCity>(x => x.EmployeeCount, "Employees");
            var list = (from x in excel.Worksheet<CompanyWithCity>("Sheet1")
                        select x)
                        .Max(x => x.EmployeeCount);
        }

        [Test]
        [ExpectedException(typeof(DataException), "'Employees' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'")]
        public void bad_column_in_min_aggregate()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.AddMapping<CompanyWithCity>(x => x.EmployeeCount, "Employees");
            var list = (from x in excel.Worksheet<CompanyWithCity>("Sheet1")
                        select x)
                        .Min(x => x.EmployeeCount);
        }

        [Test]
        [ExpectedException(typeof(DataException), "'Employees' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'")]
        public void bad_column_in_sum_aggregate()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.AddMapping<CompanyWithCity>(x => x.EmployeeCount, "Employees");
            var list = (from x in excel.Worksheet<CompanyWithCity>("Sheet1")
                        select x)
                        .Sum(x => x.EmployeeCount);
        }
    }
}
