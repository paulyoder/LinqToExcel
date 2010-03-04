using System;
using System.Linq;
using MbUnit.Framework;
using System.IO;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Integration")]
    [TestFixture]
    public class Excel2007_IntegrationTests
    {
        string _filesDirectory;

        [TestFixtureSetUp]
        public void fs()
        {
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            _filesDirectory = Path.Combine(testDirectory, "ExcelFiles");
        }

        [Test]
        public void xlsx()
        {
            var fileName = Path.Combine(_filesDirectory, "Companies.xlsx");
            var companies = from c in ExcelQueryFactory.Worksheet<Company>("MoreCompanies", fileName)
                            select c;

            //Using ToList() because using Count() first would change the sql 
            //string to "SELECT COUNT(*)" which we're not testing here
            Assert.AreEqual(3, companies.ToList().Count);
        }

        [Test]
        public void xlsm()
        {
            var fileName = Path.Combine(_filesDirectory, "Companies.xlsm");
            var companies = from c in ExcelQueryFactory.Worksheet<Company>("MoreCompanies", fileName, null)
                            select c;

            //Using ToList() because using Count() first would change the sql 
            //string to "SELECT COUNT(*)" which we're not testing here
            Assert.AreEqual(3, companies.ToList().Count);
        }

        //don't know why xlsb isn't working. I believe it's a bug with the jet driver
        public void xlsb()
        {
            var fileName = Path.Combine(_filesDirectory, "Companies.xlsb");
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, fileName, null)
                            select c;

            //Using ToList() because using Count() first would change the sql 
            //string to "SELECT COUNT(*)" which we're not testing here
            Assert.AreEqual(7, companies.ToList().Count);
        }
    }
}
