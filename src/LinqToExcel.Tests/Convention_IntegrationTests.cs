using System;
using System.Linq;
using MbUnit.Framework;
using System.IO;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Integration")]
    [TestFixture]
    public class Convention_IntegrationTests
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
        public void select_all()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                            select c;
            
            //Using ToList() because using Count() first would change the sql 
            //string to "SELECT COUNT(*)" which we're not testing here
            Assert.AreEqual(7, companies.ToList().Count);
        }

        [Test]
        public void where_string_equals()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                            where c.CEO == "Paul Yoder"
                            select c;

            Assert.AreEqual(1, companies.ToList().Count);
        }

        [Test]
        public void where_string_not_equal()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                            where c.CEO != "Bugs Bunny"
                            select c;

            Assert.AreEqual(6, companies.ToList().Count);
        }

        [Test]
        public void where_int_equals()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                            where c.EmployeeCount == 25
                            select c;

            Assert.AreEqual(1, companies.ToList().Count);
        }

        [Test]
        public void where_int_not_equal()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                            where c.EmployeeCount != 98
                            select c;

            Assert.AreEqual(6, companies.ToList().Count);
        }

        [Test]
        public void where_int_greater_than()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                            where c.EmployeeCount > 98
                            select c;

            Assert.AreEqual(3, companies.ToList().Count);
        }

        [Test]
        public void where_int_greater_than_or_equal()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                            where c.EmployeeCount >= 98
                            select c;

            Assert.AreEqual(4, companies.ToList().Count);
        }

        [Test]
        public void where_int_less_than()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                            where c.EmployeeCount < 300
                            select c;

            Assert.AreEqual(4, companies.ToList().Count);
        }

        [Test]
        public void where_int_less_than_or_equal()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                            where c.EmployeeCount <= 300
                            select c;

            Assert.AreEqual(5, companies.ToList().Count);
        }

        [Test]
        public void where_datetime_equals()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                            where c.StartDate == new DateTime(2008, 10, 9)
                            select c;

            Assert.AreEqual(1, companies.ToList().Count);
        }

        [Test]
        public void no_exception_on_property_not_used_in_where_clause_when_column_doesnt_exist()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<CompanyWithCity>(null, _excelFileName, null)
                            select c;

            foreach (var company in companies)
                Assert.IsTrue(String.IsNullOrEmpty(company.City));
        }

        [Test]
        public void where_contains()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                            where c.CEO.Contains("Paul")
                            select c;

            Assert.AreEqual(1, companies.ToList().Count);
        }

        [Test]
        public void selection_projection()
        {
            var companyCities = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                                select new CompanyWithCity
                                {
                                    CEO = c.CEO,
                                    City = "Omaha",
                                };
            foreach (var company in companyCities)
            {
                Assert.AreEqual("Omaha", company.City);
                Assert.AreEqual("Bugs Bunny", company.CEO);
                break;
            }
        }

        [Test]
        public void selection_row_projection()
        {
            var companyCities = from c in ExcelQueryFactory.Worksheet(null, _excelFileName, null)
                                select new Company
                                {
                                    CEO = c["CEO"],
                                    Name = c["Name"]
                                };
            
            foreach (var company in companyCities)
            {
                Assert.AreEqual("ACME", company.Name);
                Assert.AreEqual("Bugs Bunny", company.CEO);
                break;
            }
        }

        [Test]
        public void selection_row_to_anonymous_projection()
        {
            var companyCities = from c in ExcelQueryFactory.Worksheet(null, _excelFileName, null)
                                select new 
                                {
                                    Employees = c["EmployeeCount"].Cast<double>(),
                                    Name = c["Name"].ToString()
                                };

            foreach (var company in companyCities)
            {
                Assert.AreEqual("ACME", company.Name);
                Assert.AreEqual(25, company.Employees);
                break;
            }
        }

        [Test]
        public void first()
        {
            var firstCompany = (from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                                select c).First();

            Assert.AreEqual("ACME", firstCompany.Name);
        }

        [Test]
        public void FirstOrDefault_returns_null_when_no_rows_returned()
        {
            var noCompany = (from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                                where c.CEO == "Nobody"
                                select c).FirstOrDefault();

            Assert.IsNull(noCompany);
        }

        [Test]
        public void count()
        {
            var companyCount = (from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                                select c).Count();

            Assert.AreEqual(7, companyCount);
        }

        [Test]
        public void long_count()
        {
            var companyCount = (from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                                select c).LongCount();

            Assert.AreEqual(7, companyCount);
        }

        [Test]
        public void sum()
        {
            var companySum = (from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                              select c).Sum(x => x.EmployeeCount);

            Assert.AreEqual(30723, companySum);
        }

        [Test]
        public void average()
        {
            var averageEmployees = (from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                                    select c).Average(x => x.EmployeeCount);

            Assert.AreEqual(4389, averageEmployees);
        }

        [Test]
        public void max()
        {
            var maxEmployees = (from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                                select c).Max(x => x.EmployeeCount);

            Assert.AreEqual(29839, maxEmployees);
        }

        [Test]
        public void min()
        {
            var minEmployees = (from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                                select c).Min(x => x.EmployeeCount);

            Assert.AreEqual(1, minEmployees);
        }

        [Test]
        public void oderby()
        {
            var minEmployees = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                               orderby c.EmployeeCount ascending
                               select c;

            Assert.AreEqual(1, minEmployees.First().EmployeeCount);
        }

        [Test]
        public void oderby_desc()
        {
            var minEmployees = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                               orderby c.EmployeeCount descending
                               select c;

            Assert.AreEqual(29839, minEmployees.First().EmployeeCount);
        }

        [Test]
        public void last()
        {
            var minEmployees = from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                               select c;

            Assert.AreEqual(455, minEmployees.Last().EmployeeCount);
        }

        [Test]
        public void LastOrDefault_returns_null_when_no_rows_returned()
        {
            var noCompany = (from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                             where c.CEO == "Nobody"
                             select c).LastOrDefault();

            Assert.IsNull(noCompany);
        }

        [Test]
        public void take()
        {
            var threeEmployees = (from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                                  select c).Take(3);

            Assert.AreEqual(3, threeEmployees.ToList().Count);
        }

        [Test]
        public void skip()
        {
            var threeEmployees = (from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                                  select c).Skip(3);

            Assert.AreEqual(4, threeEmployees.ToList().Count);
        }

        [Test]
        public void reverse()
        {
            var reverse = (from c in ExcelQueryFactory.Worksheet<Company>(null, _excelFileName, null)
                           select c).Reverse().ToList();

            Assert.AreEqual("Ontario Systems", reverse.First().Name);
        }


        [Test]
        public void convert_nullable_properties()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<CompanyNullable>(null, _excelFileName, null)
                            select c;

            //Using ToList() because using Count() first would change the sql 
            //string to "SELECT COUNT(*)" which we're not testing here
            Assert.AreEqual(7, companies.ToList().Count);
        }

        [Test]
        public void dbnull_fields_return_null_and_dont_throw_exception()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<CompanyNullable>("Null Dates", _excelFileName, null)
                            select c;

            //Using ToList() because using Count() first would change the sql 
            //string to "SELECT COUNT(*)" which we're not testing here
            Assert.IsNull(companies.ToList()[3].StartDate);
        }
    }
}
