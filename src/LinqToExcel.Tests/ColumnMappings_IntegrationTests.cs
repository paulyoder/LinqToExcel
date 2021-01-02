using System;
using System.Linq;
using NUnit.Framework;
using System.IO;
using log4net.Core;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [Category("Integration")]
    [TestFixture]
    public class ColumnMappings_IntegrationTests : SQLLogStatements_Helper
    {
        ExcelQueryFactory _repo;
        string _excelFileName;
        string _worksheetName;

        [OneTimeSetUp]
        public void fs()
        {
            InstantiateLogger();
            var testDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var excelFilesDirectory = Path.Combine(testDirectory, "ExcelFiles");
            _excelFileName = Path.Combine(excelFilesDirectory, "Companies.xls");
            _worksheetName = "ColumnMappings";
        }

        [SetUp]
        public void s()
        {
            _repo = new ExcelQueryFactory(new LogManagerFactory());
            _repo.FileName = _excelFileName;
        }

        [Test]
        public void all_properties_have_column_mappings()
        {
            _repo.AddMapping<Company>(x => x.Name, "Company Title");
            _repo.AddMapping<Company>(x => x.CEO, "Boss");
            _repo.AddMapping<Company>(x => x.EmployeeCount, "Number of People");
            _repo.AddMapping<Company>(x => x.StartDate, "Initiation Date");

            var companies = from c in _repo.Worksheet<Company>(_worksheetName)
                            where c.Name == "Taylor University"
                            select c;

            var rival = companies.ToList().First();
            Assert.AreEqual(1, companies.ToList().Count, "Result Count");
            Assert.AreEqual("Taylor University", rival.Name, "Name");
            Assert.AreEqual("Your Mom", rival.CEO, "CEO");
            Assert.AreEqual(400, rival.EmployeeCount, "EmployeeCount");
            Assert.AreEqual(new DateTime(1988, 7, 26), rival.StartDate, "StartDate");
        }

        [Test]
        public void some_properties_have_column_mappings()
        {
            _repo.AddMapping<Company>(x => x.CEO, "Boss");
            _repo.AddMapping<Company>(x => x.StartDate, "Initiation Date");

            var companies = from c in _repo.Worksheet<Company>(_worksheetName)
                            where c.Name == "Anderson University"
                            select c;

            Company rival = companies.ToList()[0];
            Assert.AreEqual(1, companies.ToList().Count, "Result Count");
            Assert.AreEqual("Anderson University", rival.Name, "Name");
            Assert.AreEqual("Your Mom", rival.CEO, "CEO");
            Assert.AreEqual(300, rival.EmployeeCount, "EmployeeCount");
            Assert.AreEqual(new DateTime(1988, 7, 26), rival.StartDate, "StartDate");
        }

        [Test]
        public void log_warning_when_property_with_column_mapping_not_in_where_clause_when_mapped_column_doesnt_exist()
        {
            _loggedEvents.Clear();
            _repo.AddMapping<Company>(x => x.CEO, "The Big Cheese");

            var companies = from c in _repo.Worksheet<Company>(_worksheetName)
                            select c;

            companies.GetEnumerator();
            int warningsLogged = 0;
            foreach (LoggingEvent logEvent in _loggedEvents.GetEvents())
            {
                if ((logEvent.Level == Level.Warn) &&
                    (logEvent.RenderedMessage == "'The Big Cheese' column that is mapped to the 'CEO' property does not exist in the 'ColumnMappings' worksheet"))
                    warningsLogged++;
            }
            Assert.AreEqual(1, warningsLogged);
        }

        [Test]
        public void column_mappings_with_transformation()
        {
            _repo.AddMapping<Company>(x => x.IsActive, "Active", x => x == "Y");
            var companies = from c in _repo.Worksheet<Company>(_worksheetName)
                            select c;

            foreach (var company in companies)
                Assert.AreEqual(company.StartDate > new DateTime(1980, 1, 1), company.IsActive);
        }

        [Test]
        public void transformation()
        {
            //Add transformation to change the Name value to 'Looney Tunes' if it is originally 'ACME'
            _repo.AddTransformation<Company>(p => p.Name, value => (value == "ACME") ? "Looney Tunes" : value);
            var firstCompany = (from c in _repo.Worksheet<Company>(_worksheetName)
                                select c).First();

            Assert.AreEqual("Looney Tunes", firstCompany.Name);
        }

        [Test]
        public void transformation_that_returns_null()
        {
            //Add transformation to change the Name value to 'Looney Tunes' if it is originally 'ACME'
            _repo.AddTransformation<Company>(p => p.Name, value => null);
            var firstCompany = (from c in _repo.Worksheet<Company>(_worksheetName)
                                select c).First();

            Assert.AreEqual(null, firstCompany.Name);
        }

        [Test]
        public void annotated_properties_map_to_columns()
        {
            var companies = from c in _repo.Worksheet<CompanyWithColumnAnnotations>(_worksheetName)
                            where c.Name == "Taylor University"
                            select c;

            var rival = companies.ToList().First();
            Assert.AreEqual(1, companies.ToList().Count, "Result Count");
            Assert.AreEqual("Taylor University", rival.Name, "Name");
            Assert.AreEqual("Your Mom", rival.CEO, "CEO");
            Assert.AreEqual(400, rival.EmployeeCount, "EmployeeCount");
            Assert.AreEqual(new DateTime(1988, 7, 26), rival.StartDate, "StartDate");
            Assert.AreEqual("N", rival.IsActive, "IsActive");
        }

        [Test]
        public void annotated_properties_map_to_similar_columns()
        {
            var companies = from c in _repo.Worksheet<CompanyWithSimilarColumnAnnotations>("ColumnSimilarMappings")
                            where c.Name == "Taylor University"
                            select c;

            var rival = companies.ToList().First();
            Assert.AreEqual(1, companies.ToList().Count, "Result Count");
            Assert.AreEqual("Taylor University", rival.Name, "Name");
            Assert.AreEqual("Your Mom", rival.CEO, "CEO");
            Assert.AreEqual(300, rival.EmployeeCount, "EmployeeCount");
            Assert.AreEqual(new DateTime(1988, 7, 26), rival.StartDate, "StartDate");
            Assert.AreEqual("N", rival.IsActive, "IsActive");
        }

        [Test]
        public void two_transformations_with_the_same_property_but_in_different_types()
        {
            _repo.AddTransformation<Transformation1>(x => x.Value, x => x);
            _repo.AddTransformation<Transformation2>(x => x.Value, x => x);

            var value1 = _repo.Worksheet<Transformation1>("Transformation1").First().Value;
            var value2 = _repo.Worksheet<Transformation2>("Transformation2").First().Value;

            Assert.AreEqual(1, value1);
            Assert.AreEqual("some different value", value2);
        }

    }
}
