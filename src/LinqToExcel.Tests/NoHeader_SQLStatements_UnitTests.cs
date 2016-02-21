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
    public class NoHeader_SQLStatements_UnitTests : SQLLogStatements_Helper
    {
        private ExcelQueryFactory _factory;

        [OneTimeSetUp]
        public void fs()
        {
            InstantiateLogger();
        }

        [SetUp]
        public void s()
        {
            _factory = new ExcelQueryFactory("");
            ClearLogEvents();
        }

        [Test]
        public void range_csv_file_throws_exception()
        {
            var csvFile = @"C:\ExcelFiles\NoHeaderRange.csv";

            var excel = new ExcelQueryFactory(csvFile);
            Assert.That(() => (from c in excel.WorksheetRangeNoHeader("B9", "E15")
                               select c),
            Throws.TypeOf<ArgumentException>(), "Cannot use WorksheetRangeNoHeader on csv files");

        }

        [Test]
        public void where_clause()
        {
            var warnerCompany = from c in _factory.WorksheetNoHeader()
                               where c[1] == "Bugs Bunny"
                               select c;

            try { warnerCompany.GetEnumerator(); }
            catch (OleDbException) { }

            var expectedSQL = "SELECT * FROM [Sheet1$] WHERE (F2 = ?)";
            Assert.AreEqual(expectedSQL, GetSQLStatement());
        }

        [Test]
        public void where_is_null()
        {
            var warnerCompany = from c in _factory.WorksheetNoHeader()
                                where c[1] == null
                                select c;

            try { warnerCompany.GetEnumerator(); }
            catch (OleDbException) { }

            var expectedSQL = "SELECT * FROM [Sheet1$] WHERE (F2 IS NULL)";
            Assert.AreEqual(expectedSQL, GetSQLStatement());
        }
    }
}
