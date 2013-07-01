﻿using System.Linq;
using MbUnit.Framework;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Unit")]
    [TestFixture]
    public class ConfiguredWorksheetName_SQLStatements_UnitTests : SQLLogStatements_Helper
    {
        [FixtureSetUp]
        public void fs()
        {
            InstantiateLogger();
        }

        [SetUp]
        public void Setup()
        {
            ClearLogEvents();
        }

        [Test]
        public void table_name_in_sql_statement_matches_configured_table_name()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>("Company Worksheet", "")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = "SELECT * FROM [Company Worksheet$]";
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }
    }
}
