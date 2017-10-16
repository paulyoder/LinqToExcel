using System;
using System.Linq;
using NUnit.Framework;
using System.Data.OleDb;
using LinqToExcel.Domain;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Unit")]
    [TestFixture]
    public class ConnectionString_UnitTests : SQLLogStatements_Helper
    {

        private bool Is64BitProcess() {
            return IntPtr.Size == 8;
        }

        [OneTimeSetUp]
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
            var companies = from c in ExcelQueryFactory.Worksheet(null, "spreadsheet.xls", null, new LogManagerFactory())
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expected;

            if (Is64BitProcess()) {
                expected = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                    "spreadsheet.xls"
                );
            } else {
                expected = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1""",
                    "spreadsheet.xls"
                );
            }

            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xls_readonly_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.xls", new LogManagerFactory());
            excel.ReadOnly = true;

            var companies = from c in excel.Worksheet<Company>()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expected;

            if (Is64BitProcess()) {
                expected = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1;READONLY=TRUE""",
                    "spreadsheet.xls"
                );
            } else {
                expected = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;READONLY=TRUE""",
                    "spreadsheet.xls"
                );
            }

            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xls_with_Ace_DatabaseEngine_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.xls", new LogManagerFactory());
            excel.DatabaseEngine = DatabaseEngine.Ace;

            var companies = from c in excel.Worksheet()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                "spreadsheet.xls");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xls_readonly_with_Ace_DatabaseEngine_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.xls", new LogManagerFactory());
            excel.DatabaseEngine = DatabaseEngine.Ace;
            excel.ReadOnly = true;

            var companies = from c in excel.Worksheet()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1;READONLY=TRUE""",
                "spreadsheet.xls");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void unknown_file_type_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet(null, "spreadsheet.dlo", null, new LogManagerFactory())
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expected;

            if (Is64BitProcess()) {
                expected = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                    "spreadsheet.dlo"
                );
            } else {
                expected = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1""",
                    "spreadsheet.dlo"
                );
            }

            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void unknown_file_type_readonly_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.dlo", new LogManagerFactory());
            excel.ReadOnly = true;

            var companies = from c in excel.Worksheet()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expected;

            if (Is64BitProcess()) {
                expected = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1;READONLY=TRUE""",
                    "spreadsheet.dlo"
                );
            } else {
                expected = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;READONLY=TRUE""",
                    "spreadsheet.dlo"
                );
            }

            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void unknown_file_type_with_Ace_DatabaseEngine_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.dlo", new LogManagerFactory());
            excel.DatabaseEngine = DatabaseEngine.Ace;

            var companies = from c in excel.Worksheet()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                "spreadsheet.dlo");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void unknown_file_type_readonly_with_Ace_DatabaseEngine_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.dlo", new LogManagerFactory());
            excel.DatabaseEngine = DatabaseEngine.Ace;
            excel.ReadOnly = true;

            var companies = from c in excel.Worksheet()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1;READONLY=TRUE""",
                "spreadsheet.dlo");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void csv_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet(null, @"C:\Desktop\spreadsheet.csv", null, new LogManagerFactory())
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }

            string expected;

            if (Is64BitProcess()) {
                expected = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""text;Excel 12.0;HDR=YES;IMEX=1""",
                    @"C:\Desktop"
                );
            } else {
                expected = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""text;HDR=YES;FMT=Delimited;IMEX=1""",
                    @"C:\Desktop"
                );
            }

            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void csv_readonly_connection_string()
        {
            var excel = new ExcelQueryFactory(@"C:\Desktop\spreadsheet.csv", new LogManagerFactory());
            excel.ReadOnly = true;

            var companies = from c in excel.Worksheet<Company>()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expected;

            if (Is64BitProcess()) {
                expected = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""text;Excel 12.0;HDR=YES;IMEX=1;READONLY=TRUE""",
                    @"C:\Desktop"
                );
            } else {
                expected = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""text;HDR=YES;FMT=Delimited;IMEX=1;READONLY=TRUE""",
                    @"C:\Desktop"
                );
            }

            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void csv_with_Ace_DatabaseEngine_connection_string()
        {
            var excel = new ExcelQueryFactory(@"C:\Desktop\spreadsheet.csv", new LogManagerFactory());
            excel.DatabaseEngine = DatabaseEngine.Ace;

            var companies = from c in excel.Worksheet()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""text;Excel 12.0;HDR=YES;IMEX=1""",
                @"C:\Desktop");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void csv_readonly_with_Ace_DatabaseEngine_connection_string()
        {
            var excel = new ExcelQueryFactory(@"C:\Desktop\spreadsheet.csv", new LogManagerFactory());
            excel.DatabaseEngine = DatabaseEngine.Ace;
            excel.ReadOnly = true;

            var companies = from c in excel.Worksheet()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""text;Excel 12.0;HDR=YES;IMEX=1;READONLY=TRUE""",
                @"C:\Desktop");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xlsx_connection_string()
        {
            var companies = from c in ExcelQueryFactory.Worksheet(null, "spreadsheet.xlsx", null, new LogManagerFactory())
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""",
                "spreadsheet.xlsx");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xlsx_readonly_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.xlsx", new LogManagerFactory());
            excel.ReadOnly = true;

            var companies = from c in excel.Worksheet<Company>()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;READONLY=TRUE""",
                "spreadsheet.xlsx");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xlsx_with_Jet_DatabaseEngine_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.xlsx", new LogManagerFactory());
            excel.DatabaseEngine = DatabaseEngine.Jet;

            var companies = from c in excel.Worksheet()
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
            var companies = from c in ExcelQueryFactory.Worksheet(null, "spreadsheet.xlsm", null, new LogManagerFactory())
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""",
                "spreadsheet.xlsm");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xlsm_with_Jet_DatabaseEngine_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.xlsm", new LogManagerFactory());
            excel.DatabaseEngine = DatabaseEngine.Jet;

            var companies = from c in excel.Worksheet()
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
            var companies = from c in ExcelQueryFactory.Worksheet(null, "spreadsheet.xlsb", null, new LogManagerFactory())
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                "spreadsheet.xlsb");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xlsb_readonly_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.xlsb", new LogManagerFactory());
            excel.ReadOnly = true;

            var companies = from c in excel.Worksheet<Company>()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1;READONLY=TRUE""",
                "spreadsheet.xlsb");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void xlsb_with_Jet_DatabaseEngine_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.xlsb", new LogManagerFactory());
            excel.DatabaseEngine = DatabaseEngine.Jet;

            var companies = from c in excel.Worksheet()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""",
                "spreadsheet.xlsb");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void no_header_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.xls", new LogManagerFactory());
            var companies = from c in excel.WorksheetNoHeader()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expected;

            if (Is64BitProcess()) {
                expected = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=NO;IMEX=1""",
                    "spreadsheet.xls"
                );
            } else {
                expected = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=NO;IMEX=1""",
                    "spreadsheet.xls"
                );
            }

            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void no_header_readonly_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.xls", new LogManagerFactory());
            excel.ReadOnly = true;

            var companies = from c in excel.WorksheetNoHeader()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expected;

            if (Is64BitProcess()) {
                expected = string.Format(
                    @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=NO;IMEX=1;READONLY=TRUE""",
                    "spreadsheet.xls"
                );
            } else {
                expected = string.Format(
                    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=NO;IMEX=1;READONLY=TRUE""",
                    "spreadsheet.xls"
                );
            }

            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void no_header_with_Jet_DatabaseEngine_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.xls", new LogManagerFactory());
            excel.DatabaseEngine = DatabaseEngine.Ace;

            var companies = from c in excel.WorksheetNoHeader()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=NO;IMEX=1""",
                "spreadsheet.xls");
            Assert.AreEqual(expected, GetConnectionString());
        }

        [Test]
        public void no_header_readonly_with_Jet_DatabaseEngine_connection_string()
        {
            var excel = new ExcelQueryFactory("spreadsheet.xls", new LogManagerFactory());
            excel.DatabaseEngine = DatabaseEngine.Ace;
            excel.ReadOnly = true;

            var companies = from c in excel.WorksheetNoHeader()
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expected = string.Format(
                @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=NO;IMEX=1;READONLY=TRUE""",
                "spreadsheet.xls");
            Assert.AreEqual(expected, GetConnectionString());
        }
    }
}
