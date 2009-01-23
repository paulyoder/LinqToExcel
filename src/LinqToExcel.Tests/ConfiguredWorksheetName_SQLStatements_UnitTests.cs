using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Unit")]
    [TestFixture]
    public class ConfiguredWorksheetName_SQLStatements_UnitTests : SQLLogStatements_Helper
    {
        [TestFixtureSetUp]
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
            IExcelRepository<Company> repo = new ExcelRepository<Company>();
            var companies = from c in repo.Worksheet("Company Worksheet")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = "SELECT * FROM [Company Worksheet$]";
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }
    }
}
