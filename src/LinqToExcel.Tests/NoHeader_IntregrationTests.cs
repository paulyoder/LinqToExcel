using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NUnit.Framework;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Integration")]
    [TestFixture]
    public class NoHeader_IntregrationTests
    {
        private IExcelQueryFactory _factory;
        private string _excelFileName;

        [OneTimeSetUp]
        public void ts()
        {
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "NoHeader.xls");
        }

        [SetUp]
        public void s()
        {
            _factory = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
        }

        [Test]
        public void NoHeader_no_args()
        {
            var companies = from c in _factory.WorksheetNoHeader()
                            select c;

            Assert.AreEqual(7, companies.Count(), "Company Count");
            Assert.AreEqual("ACME", companies.First()[0].ToString(), "First Company Name");
            Assert.AreEqual(455, companies.Last()[2].Cast<int>(), "Last Company Employee Count");
        }

        [Test]
        public void csv_file()
        {
            var csvFile = Path.Combine(
                Path.GetDirectoryName(_excelFileName),
                "NoHeader.csv");

            _factory = new ExcelQueryFactory(csvFile, new LogManagerFactory());
            var companies = from c in _factory.WorksheetNoHeader()
                            select c;

            Assert.AreEqual(7, companies.Count(), "Company Count");
            Assert.AreEqual("ACME2", companies.First()[0].ToString(), "First Company Name");
            Assert.AreEqual(4554, companies.Last()[2].Cast<int>(), "Last Company Employee Count");
        }

        [Test]
        public void where_clause()
        {
            var oldCompanies = from c in _factory.WorksheetNoHeader()
                               where c[3].Cast<DateTime>() < new DateTime(1955, 1, 1)
                               select c;

            Assert.AreEqual(3, oldCompanies.Count(), "Company Count");
            Assert.AreEqual("McDonalds", oldCompanies.Last()[0].ToString(), "Last OldCompany Name");
        }

        [Test]
        public void where_is_null()
        {
            var factory = new ExcelQueryFactory(_excelFileName.Replace("NoHeader.xls", "Companies.xlsx"), new LogManagerFactory());
            var oldCompanies = from c in factory.WorksheetNoHeader("NullCells")
                               where c[2] == null
                               select c;

            Assert.AreEqual(2, oldCompanies.Count(), "Company Count");
        }

        [Test]
        public void no_header_range_where_clause()
        {
            var ACompanies = from c in _factory.WorksheetRangeNoHeader("C5", "F11", "Range")
                             where c[0].ToString().StartsWith("A")
                             select c;

            Assert.AreEqual(2, ACompanies.Count(), "Company Count");
            Assert.AreEqual(new DateTime(1917, 9, 1).ToShortDateString(), 
                ACompanies.Last()[3].Cast<DateTime>().ToShortDateString(), 
                "Last ACompany Date");
        }

        [Test]
        public void WorksheetNoHeader_WorksheetName_arg()
        {
            var companies = from c in _factory.WorksheetNoHeader("Sheet1")
                            select c;

            Assert.AreEqual(7, companies.Count(), "Company Count");
            Assert.AreEqual("ACME", companies.First()[0].ToString(), "First Company Name");
            Assert.AreEqual(455, companies.Last()[2].Cast<int>(), "Last Company Employee Count");
        }

        [Test]
        public void WorksheetNoHeader_WorksheetIndex_arg()
        {
            var companies = from c in _factory.WorksheetNoHeader(1)
                            select c;

            Assert.AreEqual(7, companies.Count(), "Company Count");
            Assert.AreEqual("ACME", companies.First()[0].ToString(), "First Company Name");
            Assert.AreEqual(455, companies.Last()[2].Cast<int>(), "Last Company Employee Count");
        }

        [Test]
        public void WorksheetRangeNoHeader_no_arg()
        {
            var companies = from c in _factory.WorksheetRangeNoHeader("A1", "D7")
                            select c;

            Assert.AreEqual(7, companies.Count(), "Company Count");
            Assert.AreEqual("ACME", companies.First()[0].ToString(), "First Company Name");
            Assert.AreEqual(455, companies.Last()[2].Cast<int>(), "Last Company Employee Count");
        }

        [Test]
        public void WorksheetRangeNoHeader_WorksheetName_arg()
        {
            var companies = from c in _factory.WorksheetRangeNoHeader("C5", "F11", "Range")
                            select c;

            Assert.AreEqual(7, companies.Count(), "Company Count");
            Assert.AreEqual("ACME", companies.First()[0].ToString(), "First Company Name");
            Assert.AreEqual(455, companies.Last()[2].Cast<int>(), "Last Company Employee Count");
        }

        [Test]
        public void WorksheetRangeNoHeader_WorksheetIndex_arg()
        {
            var companies = from c in _factory.WorksheetRangeNoHeader("C5", "F11", 0)
                            select c;

            Assert.AreEqual(7, companies.Count(), "Company Count");
            Assert.AreEqual("ACME", companies.First()[0].ToString(), "First Company Name");
            Assert.AreEqual(455, companies.Last()[2].Cast<int>(), "Last Company Employee Count");
        }
    }
}
