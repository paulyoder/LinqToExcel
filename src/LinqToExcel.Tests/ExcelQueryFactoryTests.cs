using System.Linq;
using MbUnit.Framework;
using System;
using System.IO;
using System.Data;
using LinqToExcel.Domain;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Unit")]
    [TestFixture]
    public class ExcelQueryFactoryTests
    {
        private string _excelFileName;

        [SetUp]
        public void s()
        {
            var excelFilesDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
        }

        [Test]
        [ExpectedArgumentNullException]
        public void throw_argumentnullexception_when_filename_is_null()
        {
            var repo = new ExcelQueryFactory();
            var first = (from r in repo.Worksheet() select r).First();
        }

        [Test]
        public void Constructor_sets_filename()
        {
            var repo = new ExcelQueryFactory(@"C:\spreadsheet.xls");
            Assert.AreEqual(@"C:\spreadsheet.xls", repo.FileName);
        }

        [Test]
        [ExpectedException(typeof(NullReferenceException), "FileName property is not set")]
        public void GetWorksheetNames_throws_exception_when_filename_not_set()
        {
            var factory = new ExcelQueryFactory();
            factory.GetWorksheetNames();
        }

        [Test]
        [ExpectedException(typeof(NullReferenceException), "FileName property is not set")]
        public void GetColumnNames_throws_exception_when_filename_not_set()
        {
            var factory = new ExcelQueryFactory();
            factory.GetColumnNames("");
        }

        [Test]
        public void GetWorksheetNames_returns_worksheet_names()
        {
            var excel = new ExcelQueryFactory(_excelFileName);

            var worksheetNames = excel.GetWorksheetNames();
            Assert.AreEqual(
                "ColumnMappings, IMEX Table, More Companies, Null Dates, Range1, Sheet1", 
                string.Join(", ", worksheetNames.ToArray()));
        }

        [Test]
        public void GetColumnNames_returns_column_names()
        {
            var excel = new ExcelQueryFactory(_excelFileName);

            var columnNames = excel.GetColumnNames("Sheet1");
            Assert.AreEqual(
                "Name, CEO, EmployeeCount, StartDate",
                string.Join(", ", columnNames.ToArray()));
        }

        [Test]
        [ExpectedException(typeof(StrictMappingException), "'City' property is not mapped to a column")]
        public void StrictMapping_throws_StrictMappingException_when_property_is_not_mapped_to_column()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.StrictMapping = true;
            var companies = (from x in excel.Worksheet<CompanyWithCity>()
                             select x).ToList();
        }

        [Test]
        [ExpectedException(typeof(StrictMappingException), "'City' column is not mapped to a property")]
        public void StrictMapping_throws_StrictMappingException_when_column_is_not_mapped_to_property()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.StrictMapping = true;
            var companies = (from x in excel.Worksheet<Company>("Null Dates")
                             select x).ToList();
        }

        [Test]
        public void StrictMapping_with_column_mappings_doesnt_throw_exception()
        {
            var excel = new ExcelQueryFactory(_excelFileName);
            excel.StrictMapping = true;
            excel.AddMapping<Company>(x => x.IsActive, "Active");

            var companies = (from c in excel.Worksheet<Company>("More Companies")
                             where c.Name == "ACME"
                             select c).ToList();

            Assert.AreEqual(1, companies.Count);
        }
    }
}
