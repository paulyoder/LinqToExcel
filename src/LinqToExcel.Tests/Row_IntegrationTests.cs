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
    public class Row_IntegrationTests
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
        public void column_values()
        {
            var firstCompany = (from c in ExcelQueryFactory.Worksheet(null, _excelFileName, null)
                                select c).First();

            Assert.AreEqual("ACME", firstCompany["Name"].ToString());
            Assert.AreEqual("Bugs Bunny", firstCompany["CEO"].ToString());
            Assert.AreEqual(25, firstCompany["EmployeeCount"].Cast<int>());
            Assert.AreEqual(new DateTime(1918, 11, 11).Date, firstCompany["StartDate"].Cast<DateTime>().Date);
        }

        [Test]
        public void columnNames_returns_list_of_column_names()
        {
            var firstCompany = (from c in ExcelQueryFactory.Worksheet(null, _excelFileName, null)
                                select c).First();

            Assert.AreEqual(4, firstCompany.ColumnNames.Count());
            Assert.IsTrue(firstCompany.ColumnNames.Contains("Name"));
            Assert.IsTrue(firstCompany.ColumnNames.Contains("CEO"));
            Assert.IsTrue(firstCompany.ColumnNames.Contains("EmployeeCount"));
            Assert.IsTrue(firstCompany.ColumnNames.Contains("StartDate"));
        }
    }
}
