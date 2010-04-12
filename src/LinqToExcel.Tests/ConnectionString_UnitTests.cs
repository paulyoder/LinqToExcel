using System.Linq;
using MbUnit.Framework;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Unit")]
    [TestFixture]
    public class ConnectionString_UnitTests : SQLLogStatements_Helper
    {
        [TestFixtureSetUp]
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
        public void xls_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet(null, "spreadsheet.xls", null)
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1""",
                "spreadsheet.xls");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void unknown_file_type_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet(null, "spreadsheet.dlo", null)
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1""",
                "spreadsheet.dlo");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void csv_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet(null, @"C:\Desktop\spreadsheet.csv", null)
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""text;HDR=YES;FMT=Delimited;IMEX=1""",
                @"C:\Desktop");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xlsx_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet(null, "spreadsheet.xlsx", null)
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""",
                "spreadsheet.xlsx");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xlsm_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet(null, "spreadsheet.xlsm", null)
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""",
                "spreadsheet.xlsm");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xlsb_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet(null, "spreadsheet.xlsb", null)
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                "spreadsheet.xlsb");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void no_header()
        {
            var excel = new ExcelQueryFactory("spreadsheet.xls");
            var companies = from c in excel.WorksheetNoHeader()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=NO;IMEX=1""",
                "spreadsheet.xls");
            Assert.AreEqual(expected, GetConnectionString());
        }
    }
}
