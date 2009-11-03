using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;
using System.IO;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Integration")]
    [TestFixture]
    public class IMEX_Tests
    {
        string _excelFileName;

        [TestFixtureSetUp]
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
