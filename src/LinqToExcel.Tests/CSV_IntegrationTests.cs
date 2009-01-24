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
    public class CSV_IntegrationTests
    {
        IExcelRepository _repo;
        string _fileName;

        [TestFixtureSetUp]
        public void fs()
        {
            string testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _fileName = Path.Combine(excelFilesDirectory, "Companies.csv");
        }

        [SetUp]
        public void s()
        {
            _repo = new ExcelRepository(_fileName);
        }

        [Test]
        public void select_all()
        {
            var companies = from c in _repo.Worksheet()
                            select c;

            Assert.AreEqual(7, companies.ToList().Count);
        }

        [Test]
        public void where_contains_string_criteria()
        {
            var companies = from c in _repo.Worksheet()
                            where c["Name"].ValueAs<string>() == "ACME"
                            select c;

            Assert.AreEqual(1, companies.ToList().Count);
        }

        [Test]
        public void where_contains_int_criteria()
        {
            var companies = from c in _repo.Worksheet()
                            where c["EmployeeCount"].ValueAs<int>() > 20
                            select c;

            Assert.AreEqual(5, companies.ToList().Count);
        }

        [Test]
        public void where_contains_datetime_criteria()
        {
            var companies = from c in _repo.Worksheet()
                            where c["StartDate"].ValueAs<DateTime>() == new DateTime(1980, 8, 23)
                            select c;

            Assert.AreEqual(1, companies.ToList().Count);
        }
    }
}
