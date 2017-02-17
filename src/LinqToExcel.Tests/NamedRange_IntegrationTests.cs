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
    public class NamedRange_IntegrationTests
    {
        ExcelQueryFactory _factory;
        string _excelFileName;

        [TestFixtureSetUp]
        public void fs()
        {
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
        }

        [SetUp]
        public void s()
        {
            _factory = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
        }

        [Test]
        public void GetNamedRanges()
        {
            var namedRanges = from c in _factory.GetNamedRanges()
                              select c;

            Assert.AreEqual("Companies, CompaniesRange, MoreCompanies", string.Join(", ", namedRanges.ToArray()), "Workbook-scope");
        }

        [Test]
        public void GetCompaniesRangeNamedRanges()
        {
            var namedRanges = from c in _factory.GetNamedRanges("Range1")
                            select c;

            Assert.AreEqual("CompaniesRange", string.Join(", ", namedRanges.ToArray()), "Worksheet-scope");
        }

        [Test]
        public void use_sheetData_and_worksheetIndex()
        {
            var companies = from c in _factory.NamedRange<Company>(4, "CompaniesRange")
                            select c;

            Assert.AreEqual(7, companies.Count(), "Count");
            Assert.AreEqual("ACME", companies.First().Name, "First Company Name");
        }

        [Test]
        public void use_row_and_worksheetIndex()
        {
            var companies = from c in _factory.NamedRange(4, "CompaniesRange")
                            select c;

            Assert.AreEqual(7, companies.Count(), "Count");
            Assert.AreEqual("Ontario Systems", companies.Last()["Name"].ToString(), "Last Company Name");
        }


        [Test]
        public void use_sheetData_where_null()
        {
            var factory = new ExcelQueryFactory(_excelFileName + "x", new LogManagerFactory());
            var companies = from c in factory.NamedRange<Company>("NullCellCompanies")
                            where c.EmployeeCount == null
                            select c;

            Assert.AreEqual(2, companies.Count(), "Count");
        }

        [Test]
        public void use_row_where_null()
        {
            var factory = new ExcelQueryFactory(_excelFileName + "x", new LogManagerFactory());
            var companies = from c in factory.NamedRange("NullCellCompanies")
                            where c["EmployeeCount"] == null
                            select c;

            Assert.AreEqual(2, companies.Count(), "Count");
        }

        [Test]
        public void use_row_no_header_where_null()
        {
            var factory = new ExcelQueryFactory(_excelFileName + "x", new LogManagerFactory());
            var companies = from c in factory.NamedRangeNoHeader("NullCellCompanies")
                            where c[2] == null
                            select c;

            Assert.AreEqual(2, companies.Count(), "Count");
        }
    }
}
