using System;
using System.Linq;
using MbUnit.Framework;
using System.IO;
using System.Data;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Integration")]
    [TestFixture]
    public class ConfiguredWorksheetName_IntegrationTests : SQLLogStatements_Helper
    {
        private string _excelFileName;

        [TestFixtureSetUp]
        public void fs()
        {
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
            InstantiateLogger();
        }

        [SetUp]
        public void s()
        {
            ClearLogEvents();
        }

        [Test]
        public void data_is_read_from_correct_worksheet()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>("More Companies", _excelFileName)
                            select c;

            Assert.AreEqual(3, companies.ToList().Count);
        }

        [Test]
        public void worksheetIndex_of_2_uses_third_table_name_orderedby_name()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet<Company>(2, _excelFileName)
                             select c).ToList();

            var expectedSql = "SELECT * FROM [More Companies$]";
            Assert.AreEqual(expectedSql, GetSQLStatement(), "SQL Statement");
        }

        [Test]
        [ExpectedException(typeof(System.Data.DataException), "Worksheet Index Out of Range")]
        public void worksheetIndex_too_high_throws_exception()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(8, _excelFileName)
                            select c;

            companies.ToList();
        }
    }
}
