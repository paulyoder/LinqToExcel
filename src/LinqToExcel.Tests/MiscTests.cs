using System.Linq;
using MbUnit.Framework;
using System;
using System.IO;
using System.Data;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Unit")]
    [TestFixture]
    public class MiscTests
    {
        [Test]
        [ExpectedArgumentNullException]
        public void throw_argumentnullexception_when_filename_is_null()
        {
            var repo = new ExcelQueryFactory();
            var first = (from r in repo.Worksheet() select r).First();
        }

        [Test]
        public void excelqueryfactory_constructor_sets_filename()
        {
            var repo = new ExcelQueryFactory(@"C:\spreadsheet.xls");
            Assert.AreEqual(@"C:\spreadsheet.xls", repo.FileName);
        }
    }
}
