using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using System.IO;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Integration")]
    [TestFixture]
    public class Range_IntegrationTests
    {
        ExcelQueryFactory _factory;
        string _excelFileName;

        [OneTimeSetUp]
        public void fs()
        {
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
        }

        [SetUp]
        public void s()
        {
            _factory = new ExcelQueryFactory(_excelFileName);
        }

        [Test]
        public void use_sheetData_and_worksheetIndex()
        {
            var companies = from c in _factory.WorksheetRange<Company>("C6", "F13", 4)
                            select c;

            Assert.AreEqual(7, companies.Count(), "Count");
            Assert.AreEqual("ACME", companies.First().Name, "First Company Name");
        }

        [Test]
        public void use_row_and_worksheetIndex()
        {
            var companies = from c in _factory.WorksheetRange("c6", "f13", 4)
                            select c;

            Assert.AreEqual(7, companies.Count(), "Count");
            Assert.AreEqual("Ontario Systems", companies.Last()["Name"].ToString(), "Last Company Name");
        }


        [Test]
        public void use_sheetData_where_null()
        {
            var factory = new ExcelQueryFactory(_excelFileName + "x");
            var companies = from c in factory.WorksheetRange<Company>("A1", "D4", "NullCells")
                            where c.EmployeeCount == null
                            select c;

            Assert.AreEqual(2, companies.Count(), "Count");
        }

        [Test]
        public void use_row_where_null()
        {
            var factory = new ExcelQueryFactory(_excelFileName + "x");
            var companies = from c in factory.WorksheetRange("A1", "D4", "NullCells")
                            where c["EmployeeCount"] == null
                            select c;

            Assert.AreEqual(2, companies.Count(), "Count");
        }

        [Test]
        public void use_row_no_header_where_null()
        {
            var factory = new ExcelQueryFactory(_excelFileName + "x");
            var companies = from c in factory.WorksheetRangeNoHeader("A1", "D4", "NullCells")
                            where c[2] == null
                            select c;

            Assert.AreEqual(2, companies.Count(), "Count");
        }
    }
}
