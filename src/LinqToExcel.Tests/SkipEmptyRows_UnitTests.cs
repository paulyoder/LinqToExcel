using System;
using System.IO;
using System.Linq;
using NUnit.Framework;

namespace LinqToExcel.Tests
{
    [Category("Unit")]
    [TestFixture]
    public class SkipEmptyRows_UnitTests
    {
        ExcelQueryFactory _factory;
        string _fileName;

        [OneTimeSetUp]
        public void fs()
        {
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _fileName = Path.Combine(excelFilesDirectory, "EmptyRows.xls");
        }

        [SetUp]
        public void s()
        {
            _factory = new ExcelQueryFactory(_fileName, new LogManagerFactory());
        }

        [Test]
        public void CanGetAllRows_GetTypeResults()
        {
            _factory.SkipEmptyRows = false;
            var skippedRows = from row in _factory.Worksheet<Company>("EmptyRows")
                select row;

            Assert.AreEqual(6, skippedRows.ToList().Count);
        }

        [Test]
        public void CanGetAllRows_GetRowResults()
        {
            _factory.SkipEmptyRows = false;
            var skippedRows = from row in _factory.Worksheet("EmptyRows")
                select row;

            Assert.AreEqual(6, skippedRows.ToList().Count);
        }

        [Test]
        public void CanGetAllRows_GetRowNoHeaderResults()
        {
            _factory.SkipEmptyRows = false;
            var skippedRows = from row in _factory.WorksheetNoHeader("EmptyRowsNoHeader")
                select row;

            Assert.AreEqual(6, skippedRows.ToList().Count);
        }

        [Test]
        public void CanSkipEmptyRows_GetTypeResults()
        {
            _factory.SkipEmptyRows = true;
            var skippedRows = from row in _factory.Worksheet<Company>("EmptyRows")
                select row;

            Assert.AreEqual(3, skippedRows.ToList().Count);
        }

        [Test]
        public void CanSkipEmptyRows_GetRowResults()
        {
            _factory.SkipEmptyRows = true;
            var skippedRows = from row in _factory.Worksheet("EmptyRows")
                select row;

            Assert.AreEqual(3, skippedRows.ToList().Count);
        }

        [Test]
        public void CanSkipEmptyRows_GetRowNoHeaderResults()
        {
            _factory.SkipEmptyRows = true;
            var skippedRows = from row in _factory.WorksheetNoHeader("EmptyRowsNoHeader")
                select row;

            Assert.AreEqual(3, skippedRows.ToList().Count);
        }
    }
}
