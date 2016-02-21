using System;
using System.Linq;
using NUnit.Framework;
using System.IO;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Integration")]
    [TestFixture]
    public class IMEX_Tests
    {
        string _excelFileName;

        [OneTimeSetUp]
        public void fs()
        {
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
        }

        [Test]
        public void date_and_text_column_values_are_not_null()
        {
            var sheet = new ExcelQueryFactory();
            sheet.FileName = _excelFileName;

            var names = (from x in sheet.Worksheet("IMEX Table")
                         select x).ToList();
            
            Assert.AreEqual("Bye", names.Last()["Date"].ToString());
        }
    }
}
