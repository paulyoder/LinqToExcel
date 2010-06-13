using System.Linq;
using MbUnit.Framework;
using System;
using System.IO;
using System.Data;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Unit")]
    [TestFixture]
    public class ExcelQueryFactoryTests
    {
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
            var excelFilesDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelFiles");
            var excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
            var excel = new ExcelQueryFactory(excelFileName);

            var worksheetNames = excel.GetWorksheetNames();
            Assert.AreEqual(
                "ColumnMappings, IMEX Table, More Companies, Null Dates, Range1, Sheet1", 
                string.Join(", ", worksheetNames.ToArray()));
        }

        [Test]
        public void GetColumnNames_returns_column_names()
        {
            var excelFilesDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelFiles");
            var excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
            var excel = new ExcelQueryFactory(excelFileName);

            var columnNames = excel.GetColumnNames("Sheet1");
            Assert.AreEqual(
                "Name, CEO, EmployeeCount, StartDate",
                string.Join(", ", columnNames.ToArray()));
        }
    }
}
