﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Unit")]
    [TestFixture]
    public class Range_SQLStatements_UnitTests : SQLLogStatements_Helper
    {
        private IExcelQueryFactory _factory;

        [FixtureSetUp]
        public void fs()
        {
            InstantiateLogger();
        }

        [SetUp]
        public void s()
        {
            ClearLogEvents();
            _factory = new ExcelQueryFactory("");
        }

        [Test]
        public void Appends_range_info_to_table_name()
        {
            var companies = from c in _factory.WorksheetRange("B2", "D4")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [Sheet1$B2:D4]", GetSQLStatement());
        }

        [Test]
        [ExpectedArgumentException("StartRange argument '22' is invalid format for cell name")]
        public void Throws_argument_exception_if_startRange_is_incorrect_format()
        {
            var companies = from c in _factory.WorksheetRange("22", "D4")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
        }

        [Test]
        [ExpectedArgumentException("EndRange argument 'DD' is invalid format for cell name")]
        public void Throws_argument_exception_if_endRange_is_incorrect_format()
        {
            var companies = from c in _factory.WorksheetRange("B2", "DD")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
        }

        [Test]
        public void use_sheetData_method()
        {
            var companies = from c in _factory.WorksheetRange<Company>("B2", "D4")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [Sheet1$B2:D4]", GetSQLStatement());
        }

        [Test]
        public void use_sheetData_worksheetName_method()
        {
            var companies = from c in _factory.WorksheetRange<Company>("B2", "D4", "worksheetName")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [worksheetName$B2:D4]", GetSQLStatement());
        }

        [Test]
        public void use_row_method()
        {
            var companies = from c in _factory.WorksheetRange("B2", "D4")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [Sheet1$B2:D4]", GetSQLStatement());
        }

        [Test]
        public void use_row_worksheetName_method()
        {
            var companies = from c in _factory.WorksheetRange("B2", "D4", "worksheetName")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [worksheetName$B2:D4]", GetSQLStatement());
        }

        [Test]
        public void use_row_where_is_null()
        {
            var companies = from c in _factory.WorksheetRange("B2", "D4", "worksheetName")
                            where c["City"] == null
                            select c;
            //System.Diagnostics.Debugger.Launch();
            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [worksheetName$B2:D4] WHERE ({0} IS NULL)", GetSQLFieldName("City"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        [Test]
        public void use_sheetData_where_is_null()
        {
            var companies = from c in _factory.WorksheetRange<Company>("B2", "D4")
                            where c.Name == null
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$B2:D4] WHERE ({0} IS NULL)", GetSQLFieldName("Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }
    }
}
