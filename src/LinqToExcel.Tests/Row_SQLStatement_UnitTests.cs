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
    public class Row_SQLStatement_UnitTests : SQLLogStatements_Helper
    {
        IExcelRepository _repo;

        [TestFixtureSetUp]
        public void fs()
        {
            _repo = new ExcelRepository();
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
            var companies = from c in _repo.Worksheet
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void column_name_used_in_where_clause()
        {
            var companies = from c in _repo.Worksheet
                            where c["City"].ToString() == "Omaha"
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} = ?)", GetSQLFieldName("City"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("Omaha", GetSQLParameters()[0]);
        }

        [Test]
        [ExpectedArgumentException("Cannot use column indexes in where clause")]
        public void argument_thrown_when_column_indexes_used_in_where_clause()
        {
            var companies = from c in _repo.Worksheet
                            where c[0].ToString() == "Omaha"
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
        }
    }
}
