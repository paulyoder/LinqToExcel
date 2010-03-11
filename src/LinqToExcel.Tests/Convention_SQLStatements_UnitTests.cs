using System.Linq;
using MbUnit.Framework;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [FixtureCategory("Unit")]
    [TestFixture]
    public class Convention_SQLStatements_UnitTests : SQLLogStatements_Helper
    {
        [TestFixtureSetUp]
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
        public void select_all()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void where_equals()
        {
            var companies = from p in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            where p.Name == "Paul"
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} = ?)", GetSQLFieldName("Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("Paul", GetSQLParameters()[0]);
        }

        [Test]
        public void where_not_equal()
        {
            var companies = from p in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            where p.Name != "Paul"
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} <> ?)", GetSQLFieldName("Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("Paul", GetSQLParameters()[0]);
        }

        [Test]
        public void where_greater_than()
        {
            var companies = from p in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            where p.EmployeeCount > 25
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} > ?)", GetSQLFieldName("EmployeeCount"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("25", GetSQLParameters()[0]);
        }

        [Test]
        public void where_greater_than_or_equal()
        {
            var companies = from p in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            where p.EmployeeCount >= 25
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} >= ?)", GetSQLFieldName("EmployeeCount"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("25", GetSQLParameters()[0]);
        }

        [Test]
        public void where_lesser_than()
        {
            var companies = from p in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            where p.EmployeeCount < 25
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} < ?)", GetSQLFieldName("EmployeeCount"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("25", GetSQLParameters()[0]);
        }

        [Test]
        public void where_lesser_than_or_equal()
        {
            var companies = from p in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            where p.EmployeeCount <= 25
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} <= ?)", GetSQLFieldName("EmployeeCount"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("25", GetSQLParameters()[0]);
        }

        [Test]
        public void where_and()
        {
            var companies = from p in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            where p.EmployeeCount > 5 && p.CEO == "Paul"
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE (({0} > ?) AND ({1} = ?))",
                                                GetSQLFieldName("EmployeeCount"),
                                                GetSQLFieldName("CEO"));
            var parameters = GetSQLParameters();

            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("5", parameters[0]);
            Assert.AreEqual("Paul", parameters[1]);
        }

        [Test]
        public void where_or()
        {
            var companies = from p in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            where p.EmployeeCount > 5 || p.CEO == "Paul"
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE (({0} > ?) OR ({1} = ?))",
                                                GetSQLFieldName("EmployeeCount"),
                                                GetSQLFieldName("CEO"));
            var parameters = GetSQLParameters();

            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("5", parameters[0]);
            Assert.AreEqual("Paul", parameters[1]);
        }

        private string GetName(string name)
        {
            return name;
        }

        [Test]
        public void method_used_in_where_clause()
        {
            var companies = from p in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            where p.Name == GetName("Paul")
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("Paul", GetSQLParameters()[0]);
        }

        [Test]
        public void where_contains()
        {
            var companies = from p in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            where p.Name.Contains("Paul")
                            select p;
            
            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} LIKE ?)", GetSQLFieldName("Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("%Paul%", GetSQLParameters()[0]);
        }

        [Test]
        public void where_startswith()
        {
            var companies = from p in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            where p.Name.StartsWith("Paul")
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} LIKE ?)", GetSQLFieldName("Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("Paul%", GetSQLParameters()[0]);
        }

        [Test]
        public void where_endswith()
        {
            var companies = from p in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            where p.Name.EndsWith("Paul")
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] WHERE ({0} LIKE ?)", GetSQLFieldName("Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("%Paul", GetSQLParameters()[0]);
        }

        [Test]
        public void first()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            select c;
            
            try { companies.First(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT TOP 1 * FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void count()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            select c;

            try { companies.Count(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT COUNT(*) FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void long_count()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            select c;

            try { companies.LongCount(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT COUNT(*) FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void sum()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            select c;

            try { companies.Sum(x => x.EmployeeCount); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT SUM(EmployeeCount) FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void average()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            select c;

            try { companies.Average(x => x.EmployeeCount); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT AVG(EmployeeCount) FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void max()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            select c;

            try { companies.Max(x => x.EmployeeCount); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT MAX(EmployeeCount) FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void min()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            select c;

            try { companies.Min(x => x.EmployeeCount); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT MIN(EmployeeCount) FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void orderby()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            orderby c.StartDate ascending
                            select c;

            try { companies.ToList(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] ORDER BY {0} ASC", GetSQLFieldName("StartDate"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        [Test]
        public void orderby_desc()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            orderby c.StartDate descending 
                            select c;

            try { companies.ToList(); }
            catch (OleDbException) { }
            var expectedSql = string.Format("SELECT * FROM [Sheet1$] ORDER BY {0} DESC", GetSQLFieldName("StartDate"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
        }

        [Test]
        public void last()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            select c;

            try { companies.Last(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void take()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             select c).Take(3);

            try { companies.ToList(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT TOP 3 * FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void skip()
        {
            var companies = (from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                             select c).Skip(3);

            try { companies.ToList(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void worksheetName_is_set_to_Sheet1_when_null_worksheetName_argument()
        {
            var companies = from c in ExcelQueryFactory.Worksheet<Company>(null, "", null)
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [Sheet1$]", GetSQLStatement());
        }
    }
}
