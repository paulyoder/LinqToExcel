using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;
using System.IO;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Integration")]
    [TestFixture]
    public class Range_IntegrationTests
    {
        ExcelQueryFactory _factory;
        string _excelFileName;
        string _worksheetName;

        [TestFixtureSetUp]
        public void fs()
        {
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
            _worksheetName = "Range1";
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
    }
}
