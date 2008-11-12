using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;
using System.IO;
using System.Reflection;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Integration")]
    [TestFixture]
    public class Convention_IntegrationTests
    {
        private string _excelFileName;

        [TestFixtureSetUp]
        public void fs()
        {
            string testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "Simple_Companies.xls");
        }

        [Test]
        public void select_all()
        {
            var companies = from c in ExcelRepository.GetSheet<Company>(_excelFileName)
                            select c;

            Assert.AreEqual(7, companies.ToList().Count);
        }

        [Test]
        public void where_string_equals()
        {
            var companies = from c in ExcelRepository.GetSheet<Company>(_excelFileName)
                            where c.CEO == "Paul Yoder"
                            select c;

            //Don't know why companies.Count() doesn't work. It throws an IndexOutOfRange exception
            Assert.AreEqual(1, companies.ToList().Count);
        }

        [Test]
        public void where_string_not_equal()
        {
            var companies = from c in ExcelRepository.GetSheet<Company>(_excelFileName)
                            where c.CEO != "Bugs Bunny"
                            select c;

            Assert.AreEqual(6, companies.ToList().Count);
        }

        [Test]
        public void where_int_equals()
        {
            var companies = from c in ExcelRepository.GetSheet<Company>(_excelFileName)
                            where c.EmployeeCount == 25
                            select c;

            Assert.AreEqual(1, companies.ToList().Count);
        }

        [Test]
        public void where_int_not_equal()
        {
            var companies = from c in ExcelRepository.GetSheet<Company>(_excelFileName)
                            where c.EmployeeCount != 98
                            select c;

            Assert.AreEqual(6, companies.ToList().Count);
        }

        [Test]
        public void where_int_greater_than()
        {
            var companies = from c in ExcelRepository.GetSheet<Company>(_excelFileName)
                            where c.EmployeeCount > 98
                            select c;

            Assert.AreEqual(3, companies.ToList().Count);
        }

        [Test]
        public void where_int_greater_than_or_equal()
        {
            var companies = from c in ExcelRepository.GetSheet<Company>(_excelFileName)
                            where c.EmployeeCount >= 98
                            select c;

            Assert.AreEqual(4, companies.ToList().Count);
        }

        [Test]
        public void where_int_less_than()
        {
            var companies = from c in ExcelRepository.GetSheet<Company>(_excelFileName)
                            where c.EmployeeCount < 300
                            select c;

            Assert.AreEqual(4, companies.ToList().Count);
        }

        [Test]
        public void where_int_less_than_or_equal()
        {
            var companies = from c in ExcelRepository.GetSheet<Company>(_excelFileName)
                            where c.EmployeeCount <= 300
                            select c;
            
            Assert.AreEqual(5, companies.ToList().Count);
        }

        [Test]
        public void where_datetime_equals()
        {
            var companies = from c in ExcelRepository.GetSheet<Company>(_excelFileName)
                            where c.StartDate == new DateTime(2008, 10, 9)
                            select c;

            Assert.AreEqual(1, companies.ToList().Count);
        }

        [Test]
        public void more_properties_on_class_than_columns_on_sheet()
        {
            var companies = from c in ExcelRepository.GetSheet<CompanyWithCity>(_excelFileName)
                            select c;

            foreach (CompanyWithCity company in companies)
                Assert.IsTrue(String.IsNullOrEmpty(company.City));
        }
    }
}
