using System;
using System.Linq;
using NUnit.Framework;
using System.IO;
using System.Data;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Integration")]
    [TestFixture]
    public class InvalidColumnNamesUsed
    {
        private string _excelFileName;

        [SetUp]
        public void fs()
        {
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
        }

        [Test]
        public void row_column_in_where_clause()
        {
            Assert.That(() => (from x in ExcelQueryFactory.Worksheet("Sheet1", _excelFileName, new LogManagerFactory())
                               where x["Bad Column"].ToString() == "nothing"
                               select x).ToList(),
           Throws.TypeOf<DataException>(), "'Bad Column' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'");
        }

        [Test]
        public void row_column_in_orderby_clause()
        {
            Assert.That(() => (from x in ExcelQueryFactory.Worksheet("Sheet1", _excelFileName, new LogManagerFactory())
                               select x)
                        .OrderBy(x => x["Bad Column"])
                        .ToList(),
           Throws.TypeOf<DataException>(), "'Bad Column' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'");
        }

        [Test]
        public void bad_column_in_where_clause()
        {
            Assert.That(() => (from x in ExcelQueryFactory.Worksheet<CompanyWithCity>("Sheet1", _excelFileName, new LogManagerFactory())
                               where x.City == "Omaha"
                               select x).ToList(),
           Throws.TypeOf<DataException>(), "'City' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'");
        }

        [Test]
        public void bad_column_mapping_in_where_clause()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.AddMapping<CompanyWithCity>(x => x.City, "Town");
            Assert.That(() => (from x in excel.Worksheet<CompanyWithCity>("Sheet1")
                               where x.City == "Omaha"
                               select x).ToList(),
            Throws.TypeOf<DataException>(), "'Town' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'");
        }

        [Test]
        public void bad_column_in_orderby_clause()
        {
            var list = (from x in ExcelQueryFactory.Worksheet<CompanyWithCity>("Sheet1", _excelFileName, new LogManagerFactory())
                select x)
                .OrderBy(x => x.City);
            Assert.That(() => list.ToList(),
            Throws.TypeOf<DataException>(), "'City' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'");
        }

        [Test]
        public void bad_column_mapping_in_orderby_clause()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.AddMapping<CompanyWithCity>(x => x.City, "Town");
            Assert.That(() => (from x in excel.Worksheet<CompanyWithCity>("Sheet1")
                               select x)
                        .OrderBy(x => x.City)
                        .ToList(),
            Throws.TypeOf<DataException>(), "'Town' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'");
        }

        [Test]
        public void bad_column_in_average_aggregate()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.AddMapping<CompanyWithCity>(x => x.EmployeeCount, "Employees");
            Assert.That(() => (from x in excel.Worksheet<CompanyWithCity>("Sheet1")
                               select x)
                        .Average(x => x.EmployeeCount),
            Throws.TypeOf<DataException>(), "'Employees' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'");
        }

        [Test]
        public void bad_column_in_max_aggregate()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.AddMapping<CompanyWithCity>(x => x.EmployeeCount, "Employees");
            Assert.That(() => (from x in excel.Worksheet<CompanyWithCity>("Sheet1")
                               select x)
                        .Max(x => x.EmployeeCount),
            Throws.TypeOf<DataException>(), "'Employees' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'");
        }

        [Test]
        public void bad_column_in_min_aggregate()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.AddMapping<CompanyWithCity>(x => x.EmployeeCount, "Employees");
            Assert.That(() => (from x in excel.Worksheet<CompanyWithCity>("Sheet1")
                               select x)
                        .Min(x => x.EmployeeCount),
            Throws.TypeOf<DataException>(), "'Employees' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'");
        }

        [Test]
        public void bad_column_in_sum_aggregate()
        {
            var excel = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
            excel.AddMapping<CompanyWithCity>(x => x.EmployeeCount, "Employees");
         
            Assert.That(() => (from x in excel.Worksheet<CompanyWithCity>("Sheet1")
                               select x)
                        .Sum(x => x.EmployeeCount),
            Throws.TypeOf<DataException>(), "'Employees' is not a valid column name. " +
            "Valid column names are: 'Name', 'CEO', 'EmployeeCount', 'StartDate'");
        }
    }
}
