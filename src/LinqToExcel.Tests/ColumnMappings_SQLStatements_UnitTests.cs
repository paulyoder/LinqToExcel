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
    public class ColumnMappings_SQLStatements_UnitTests : SQLLogStatements_Helper
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
        public void where_clause_contains_property_with_column_mapping()
        {
            Dictionary<string, string> mapping = new Dictionary<string, string>();
            mapping["CEO"] = "Boss";

            var companies = from c in ExcelRepository.GetSheet<Company>("", mapping)
                            where c.CEO == "Paul"
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] Where ({0} = ?)", GetSQLFieldName(mapping["CEO"]));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        [Test]
        public void where_clause_contains_property_without_column_mapping()
        {
            Dictionary<string, string> mapping = new Dictionary<string, string>();
            mapping["CEO"] = "Boss";

            var companies = from c in ExcelRepository.GetSheet<Company>("", mapping)
                            where c.Name == "ACME"
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] Where ({0} = ?)", GetSQLFieldName("Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }
    }
}
