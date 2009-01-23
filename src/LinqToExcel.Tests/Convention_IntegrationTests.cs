using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;
using System.IO;
using System.Reflection;
using System.Data.SqlClient;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Integration")]
    [TestFixture]
    public class Convention_IntegrationTests
    {
        string _excelFileName;
        IExcelRepository<Company> _repo;

        [TestFixtureSetUp]
        public void fs()
        {
            string testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
        }

        [SetUp]
        public void s()
        {
            _repo = new ExcelRepository<Company>(_excelFileName);
        }

        [Test]
        public void select_all()
        {
            var companies = from c in _repo.Worksheet()
                            select c;

            Assert.AreEqual(7, companies.ToList().Count);
        }

        [Test]
        public void where_string_equals()
        {
            var companies = from c in _repo.Worksheet()
                            where c.CEO == "Paul Yoder"
                            select c;

            //Don't know why companies.Count() doesn't work. It throws an IndexOutOfRange exception
            Assert.AreEqual(1, companies.ToList().Count);
        }

        [Test]
        public void where_string_not_equal()
        {
            var companies = from c in _repo.Worksheet()
                            where c.CEO != "Bugs Bunny"
                            select c;

            Assert.AreEqual(6, companies.ToList().Count);
        }

        [Test]
        public void where_int_equals()
        {
            var companies = from c in _repo.Worksheet()
                            where c.EmployeeCount == 25
                            select c;

            Assert.AreEqual(1, companies.ToList().Count);
        }

        [Test]
        public void where_int_not_equal()
        {
            var companies = from c in _repo.Worksheet()
                            where c.EmployeeCount != 98
                            select c;

            Assert.AreEqual(6, companies.ToList().Count);
        }

        [Test]
        public void where_int_greater_than()
        {
            var companies = from c in _repo.Worksheet()
                            where c.EmployeeCount > 98
                            select c;

            Assert.AreEqual(3, companies.ToList().Count);
        }

        [Test]
        public void where_int_greater_than_or_equal()
        {
            var companies = from c in _repo.Worksheet()
                            where c.EmployeeCount >= 98
                            select c;

            Assert.AreEqual(4, companies.ToList().Count);
        }

        [Test]
        public void where_int_less_than()
        {
            var companies = from c in _repo.Worksheet()
                            where c.EmployeeCount < 300
                            select c;

            Assert.AreEqual(4, companies.ToList().Count);
        }

        [Test]
        public void where_int_less_than_or_equal()
        {
            var companies = from c in _repo.Worksheet()
                            where c.EmployeeCount <= 300
                            select c;
            
            Assert.AreEqual(5, companies.ToList().Count);
        }

        [Test]
        public void where_datetime_equals()
        {
            var companies = from c in _repo.Worksheet()
                            where c.StartDate == new DateTime(2008, 10, 9)
                            select c;

            Assert.AreEqual(1, companies.ToList().Count);
        }

        [Test]
        public void no_exception_on_property_not_used_in_where_clause_when_column_doesnt_exist()
        {
            IExcelRepository<CompanyWithCity> repo = new ExcelRepository<CompanyWithCity>(_excelFileName);
            var companies = from c in repo.Worksheet()
                            select c;

            foreach (CompanyWithCity company in companies)
                Assert.IsTrue(String.IsNullOrEmpty(company.City));
        }

        //Todo
        //It is desired to have the SqlException and message thrown instead of a general OleDbException when the
        //column name is incorrect, but I don't know how to do that yet
        //[ExpectedException(typeof(SqlException), "The 'City' column does not exist in the 'Sheet1' worksheet")]
        [ExpectedException(typeof(OleDbException))]
        [Test]
        public void exception_on_property_used_in_where_clause_when_column_doesnt_exist()
        {
            IExcelRepository<CompanyWithCity> repo = new ExcelRepository<CompanyWithCity>(_excelFileName);
            var companies = from c in repo.Worksheet()
                            where c.City == "Omaha"
                            select c;

            companies.GetEnumerator();
        }
    }
}
