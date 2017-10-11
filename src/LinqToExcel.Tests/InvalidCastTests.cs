using MbUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LinqToExcel.Tests
{
    [Author("Alberto Chvaicer", "achvaicer@gmail.com")]
    [FixtureCategory("Integration")]
    [TestFixture]
    public class InvalidCastTests
    {
        private IExcelQueryFactory _factory;
        private string _excelFileName;

        [TestFixtureSetUp]
        public void fs()
        {
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
        }

        [SetUp]
        public void s()
        {
            _factory = new ExcelQueryFactory(_excelFileName, new LogManagerFactory());
        }

        [Test]
        [ExpectedException(typeof(LinqToExcel.Exceptions.ExcelException), "Error on row 8 and column name 'EmployeeCount'.")]
        public void invalid_number_cast_with_header()
        {
            var list = (from x in ExcelQueryFactory.Worksheet<Company>("Invalid Cast", _excelFileName, new LogManagerFactory())
                        select x).ToList();

        }

        [Test, Ignore]
        [ExpectedException(typeof(LinqToExcel.Exceptions.NoHeaderExcelException), "Error on row 8 and column 3.")]
        public void invalid_number_cast_without_header()
        {
            
            var list = (from x in _factory.WorksheetRangeNoHeader("A2", "D9", "Invalid Cast")
                        where x[2].Cast<double>() > 30.0
                        select x[2]).ToList();
        }


    }
}
