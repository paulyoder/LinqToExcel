using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;
using log4net;
using System.Reflection;
using log4net.Appender;
using log4net.Core;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Unit")]
    [TestFixture]
    public class Convention_SQLStatements_UnitTests : SQLLogStatements_Helper
    {
        IExcelRepository<Company> _repo;

        [TestFixtureSetUp]
        public void fs()
        {
            _repo = new ExcelRepository<Company>();
            InstantiateLogger();
        }

        [SetUp]
        public void s()
        {
            ClearLogEvents();
        }

        [Test]
        public void select_all()
        {
            var companies = from c in _repo.Worksheet()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void where_equals()
        {
            var companies = from p in _repo.Worksheet()
                         where p.Name == "Paul"
                         select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} = ?)", GetSQLFieldName("Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("Paul", GetSQLParameters()[0]);
        }

        [Test]
        public void where_not_equal()
        {
            var companies = from p in _repo.Worksheet()
                            where p.Name != "Paul"
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} <> ?)", GetSQLFieldName("Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("Paul", GetSQLParameters()[0]);
        }

        [Test]
        public void where_greater_than()
        {
            var companies = from p in _repo.Worksheet()
                            where p.EmployeeCount > 25
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} > ?)", GetSQLFieldName("EmployeeCount"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("25", GetSQLParameters()[0]);
        }

        [Test]
        public void where_greater_than_or_equal()
        {
            var companies = from p in _repo.Worksheet()
                            where p.EmployeeCount >= 25
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} >= ?)", GetSQLFieldName("EmployeeCount"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("25", GetSQLParameters()[0]);
        }

        [Test]
        public void where_lesser_than()
        {
            var companies = from p in _repo.Worksheet()
                            where p.EmployeeCount < 25
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} < ?)", GetSQLFieldName("EmployeeCount"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("25", GetSQLParameters()[0]);
        }

        [Test]
        public void where_lesser_than_or_equal()
        {
            var companies = from p in _repo.Worksheet()
                            where p.EmployeeCount <= 25
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} <= ?)", GetSQLFieldName("EmployeeCount"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("25", GetSQLParameters()[0]);
        }

        [Test]
        public void where_and()
        {
            var companies = from p in _repo.Worksheet()
                            where p.EmployeeCount > 5 && p.CEO == "Paul"
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE (({0} > ?) AND ({1} = ?))", 
                                                GetSQLFieldName("EmployeeCount"),
                                                GetSQLFieldName("CEO"));
            string[] parameters = GetSQLParameters();

            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("5", parameters[0]);
            Assert.AreEqual("Paul", parameters[1]);
        }

        [Test]
        public void where_or()
        {
            var companies = from p in _repo.Worksheet()
                            where p.EmployeeCount > 5 || p.CEO == "Paul"
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE (({0} > ?) OR ({1} = ?))",
                                                GetSQLFieldName("EmployeeCount"),
                                                GetSQLFieldName("CEO"));
            string[] parameters = GetSQLParameters();

            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("5", parameters[0]);
            Assert.AreEqual("Paul", parameters[1]);
        }

        [Test]
        public void local_field_used()
        {
            string desiredName = "Paul";
            var companies = from p in _repo.Worksheet()
                            where p.Name == desiredName
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} = ?)", GetSQLFieldName("Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("Paul", GetSQLParameters()[0]);
        }

        [Test]
        public void constructor_with_constant_value_arguments()
        {
            var companies = from p in _repo.Worksheet()
                            where p.StartDate == new DateTime(2008, 10, 9)
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} = ?)", GetSQLFieldName("StartDate"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("10/9/2008", GetSQLParameters()[0]);
        }

        [Test]
        public void constructor_with_field_value_arguments()
        {
            int year = 1876;
            int month = 6;
            int day = 25;
            var companies = from p in _repo.Worksheet()
                            where p.StartDate == new DateTime(year, month, day)
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} = ?)", GetSQLFieldName("StartDate"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("6/25/1876", GetSQLParameters()[0]);
        }

        [Test]
        public void constructor_with_property_value_arguments()
        {
            var companies = from p in _repo.Worksheet()
                            where p.StartDate == new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day)
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} = ?)", GetSQLFieldName("StartDate"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual(DateTime.Now.ToShortDateString(), GetSQLParameters()[0]);
        }

        [Test]
        public void datetime_now_is_used_in_where_clause()
        {
            var companies = from p in _repo.Worksheet()
                            where p.StartDate == DateTime.Now
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} = ?)", GetSQLFieldName("StartDate"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual(DateTime.Now.ToShortDateString(), GetSQLParameters()[0]);
        }

        private string GetName(string name)
        {
            return name;
        }

        [Test]
        public void method_used_in_where_clause()
        {
            var companies = from p in _repo.Worksheet()
                            where p.Name == GetName("Paul")
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("Paul", GetSQLParameters()[0]);
        }
    }
}
