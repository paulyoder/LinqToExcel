﻿using System;
using System.Linq;
using MbUnit.Framework;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Unit")]
    [TestFixture]
    public class ColumnMappings_SQLStatements_UnitTests : SQLLogStatements_Helper
    {
        private ExcelQueryFactory _repo;

        [FixtureSetUp]
        public void fs()
        {
            InstantiateLogger();
        }

        [SetUp]
        public void Setup()
        {
            _repo = new ExcelQueryFactory();
            _repo.FileName = "";
            ClearLogEvents();
        }

        [Test]
        public void where_clause_contains_property_with_column_mapping()
        {
            _repo.AddMapping<Company>(x => x.CEO, "Boss");

            var companies = from c in _repo.Worksheet<Company>()
                            where c.CEO == "Paul"
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} = ?)", GetSQLFieldName("Boss"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        [Test]
        public void where_clause_contains_property_without_column_mapping()
        {
            _repo.AddMapping<Company>(x => x.CEO, "Boss");

            var companies = from c in _repo.Worksheet<Company>()
                            where c.Name == "ACME"
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} = ?)", GetSQLFieldName("Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        [Test]
        public void where_is_null()
        {
            _repo.AddMapping<Company>(x => x.CEO, "Boss");

            var companies = from c in _repo.Worksheet<Company>()
                            where c.CEO == null
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} IS NULL)", GetSQLFieldName("Boss"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        /// <summary>
        /// Expression body is UnaryExpression for non native types
        /// </summary>
        [Test]
        public void mapped_property_is_not_native_type()
        {
            _repo.AddMapping<Company>(x => x.StartDate, "Hired Date");

            var companies = from c in _repo.Worksheet<Company>()
                            where c.StartDate == new DateTime(2008, 1, 1)
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} = ?)", GetSQLFieldName("Hired Date"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        /// <summary>
        /// Expression body is MemberExpression for native types
        /// </summary>
        [Test]
        public void mapped_property_is_native_type()
        {
            _repo.AddMapping<Company>(x => x.CEO, "Da big Cheese");

            var companies = from c in _repo.Worksheet<Company>()
                            where c.CEO == "Paul"
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} = ?)", GetSQLFieldName("Da big Cheese"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        [Test]
        public void multiple_mapped_properties()
        {
            _repo.AddMapping<Company>(x => x.CEO, "Da big Cheese");
            _repo.AddMapping<Company>(x => x.Name, "Legal Name");

            var companies = from c in _repo.Worksheet<Company>()
                            where c.CEO == "Paul" && c.Name == "ACME"
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE (({0} = ?) AND ({1} = ?))",
               GetSQLFieldName("Da big Cheese"),
               GetSQLFieldName("Legal Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        [Test]
        public void mapped_property_is_passed_in_as_a_string()
        {
            _repo.AddMapping("CEO", "Da big Cheese");
            _repo.AddMapping("Name", "Legal Name");

            var companies = from c in _repo.Worksheet<Company>()
                            where c.CEO == "Paul" && c.Name == "ACME"
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE (({0} = ?) AND ({1} = ?))",
               GetSQLFieldName("Da big Cheese"),
               GetSQLFieldName("Legal Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        [Test]
        public void distinct()
        {
            _repo.AddMapping("Name", "FullName");
            var companies = (from c in _repo.Worksheet<Company>()
                             select c.Name).Distinct();

            try { companies.ToList(); }
            catch (OleDbException) { }
            var expectedSql = "SELECT DISTINCT(FullName) FROM [Sheet1$]";
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }
    }
}
