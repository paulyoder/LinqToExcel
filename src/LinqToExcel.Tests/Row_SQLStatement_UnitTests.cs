﻿using System;
using System.Linq;
using MbUnit.Framework;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Unit")]
    [TestFixture]
    public class Row_SQLStatement_UnitTests : SQLLogStatements_Helper
    {
        [FixtureSetUp]
        public void fs()
        {
            InstantiateLogger();
        }

        [SetUp]
        public void s()
        {
            ClearLogEvents();
        }

        [Test]
        public void no_where_clause()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>("", "", null)
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void column_name_used_in_where_clause()
        {
            var companies = from c in ExcelQueryFactory.Worksheet("", "", null)
                            where c["City"] == "Omaha"
                            select c;
            
            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} = ?)", GetSQLFieldName("City"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("Omaha", GetSQLParameters()[0]);
        }

        [Test]
        public void column_name_used_in_orderby_clause()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet("", "", null)
                             select c).OrderBy(x => x["City"]);

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] ORDER BY {0} ASC", GetSQLFieldName("City"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        [Test]
        public void column_name_is_cast_in_where_clause()
        {
            var companies = from c in ExcelQueryFactory.Worksheet("", "", null)
                            where c["Modified"].Cast<DateTime>() < new DateTime(2009, 11, 2)
                            select c;
            
            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} < ?)", GetSQLFieldName("Modified"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        [Test]
        public void where_is_null()
        {
            var companies = from c in ExcelQueryFactory.Worksheet("", "", null)
                            where c["City"] == null
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} IS NULL)", GetSQLFieldName("City"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        [Test]
        public void where_is_not_null()
        {
            var companies = from c in ExcelQueryFactory.Worksheet("", "", null)
                            where c["City"] != null
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} IS NOT NULL)", GetSQLFieldName("City"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        [Test]
        [ExpectedArgumentException("Can only use column indexes in WHERE clause when using WorksheetNoHeader")]
        public void argument_exception_thrown_when_column_indexes_used_in_worksheet_where_clause()
        {
            var companies = from c in ExcelQueryFactory.Worksheet("", "", null)
                            where c[0] == "Omaha"
                            select c;
            
            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
        }
    }
}
