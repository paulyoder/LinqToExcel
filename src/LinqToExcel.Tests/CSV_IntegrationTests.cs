using System;
using System.Linq;
using NUnit.Framework;
using System.IO;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Integration")]
    [TestFixture]
    public class CSV_IntegrationTests
    {
        string _fileName;

        [OneTimeSetUp]
        public void fs()
        {
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _fileName = Path.Combine(excelFilesDirectory, "Companies.csv");
        }

        [Test]
        public void select_all()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _fileName, null, new LogManagerFactory())
                            select c;

            Assert.AreEqual(7, companies.ToList().Count);
        }

        [Test]
        public void where_contains_string_criteria()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _fileName, null, new LogManagerFactory())
                            where c.Name == "ACME"
                            select c;

            Assert.AreEqual(1, companies.ToList().Count);
        }

        [Test]
        public void where_contains_int_criteria()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _fileName, null, new LogManagerFactory())
                            where c.EmployeeCount > 20
                            select c;

            Assert.AreEqual(5, companies.ToList().Count);
        }

        [Test]
        public void where_contains_datetime_criteria()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _fileName, null, new LogManagerFactory())
                            where c.StartDate == new DateTime(1980, 8, 23)
                            select c;

            Assert.AreEqual(1, companies.ToList().Count);
        }
    }
}
