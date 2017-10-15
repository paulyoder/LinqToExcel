using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Unit")]
    [TestFixture]
    public class NamedRange_SQLStatements_UnitTests : SQLLogStatements_Helper
    {
        private IExcelQueryFactory _factory;

        [OneTimeSetUp]
        public void fs()
        {
            InstantiateLogger();
        }

        [SetUp]
        public void s()
        {
            ClearLogEvents();
            _factory = new ExcelQueryFactory("", new LogManagerFactory());
        }

        [Test]
        public void Appends_named_range_info_to_table_name()
        {
            var companies = from c in _factory.NamedRange("Sheet1", "NamedRange")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [Sheet1$NamedRange]", GetSQLStatement());
        }

        [Test]
        public void use_sheetData_method()
        {
            var companies = from c in _factory.NamedRange<Company>("Sheet1", "NamedRange")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [Sheet1$NamedRange]", GetSQLStatement());
        }

        [Test]
        public void use_row_where_is_null()
        {
            var companies = from c in _factory.NamedRange("Sheet1", "NamedRange")
                            where c["City"] == null
                            select c;
            //System.Diagnostics.Debugger.Launch();
            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$NamedRange] WHERE ({0} IS NULL)", GetSQLFieldName("City"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        [Test]
        public void use_sheetData_where_is_null()
        {
            var companies = from c in _factory.NamedRange<Company>("worksheetName", "NamedRange")
                            where c.Name == null
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [worksheetName$NamedRange] WHERE ({0} IS NULL)", GetSQLFieldName("Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }
    }
}
