using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;
using log4net;
using System.Reflection;
using log4net.Appender;
using log4net.Core;
using System.Data.OleDb;

namespace LinqToExcel.Tests
{
    [Author("Paul Yoder", "paulyoder@gmail.com")]
    [TestsOn(typeof(ExpressionToSQL))]
    [FixtureCategory("Unit")]
    [TestFixture]
    public class ExpressionToSQLUnitTests
    {
        /// <summary>
        /// This memory appender catches all the logged messages 
        /// which the unit tests can use for their assertions
        /// </summary>
        private MemoryAppender _loggedEvents;

        [TestFixtureSetUp]
        public void fs()
        {
            _loggedEvents = new MemoryAppender();
            log4net.Config.BasicConfigurator.Configure(_loggedEvents);
        }

        [SetUp]
        public void s()
        {
            _loggedEvents.Clear();
        }

        /// <summary>
        /// Retrieves the SQL statement from the log events
        /// </summary>
        /// <remarks>
        /// The SQL statement log message is in the following format
        /// SQL: {sql statement}
        /// </remarks>
        private string GetSQLStatement()
        {
            LoggingEvent[] loggingEvents = _loggedEvents.GetEvents();
            foreach (LoggingEvent logEvent in loggingEvents)
            {
                string message = logEvent.RenderedMessage;
                if (message.Length > 5 && message.Substring(0, 4) == "SQL:")
                    return logEvent.RenderedMessage.Substring(5);
            }
            return "";
        }

        /// <summary>
        /// Returns the SQL Parameters
        /// </summary>
        /// <remarks>
        /// The SQL Parameters log messages are in the following format
        /// Param[{param #}]: {parameter value}
        /// </remarks>
        private string[] GetSQLParameters()
        {
            LoggingEvent[] loggingEvents = _loggedEvents.GetEvents();
            List<string> parameters = new List<string>();
            foreach (LoggingEvent logEvent in loggingEvents)
            {
                string message = logEvent.RenderedMessage;
                if (message.Length > 5 && message.Substring(0, 6) == "Param[")
                    parameters.Add(logEvent.RenderedMessage.Split(" ".ToCharArray())[1]);
            }
            return parameters.ToArray();
        }

        /// <summary>
        /// Returns the sql formatted column name
        /// </summary>
        /// <param name="columnName">Name of column</param>
        private string GetSQLFieldName(string columnName)
        {
            return string.Format("[{0}]", columnName);
        }

        [Test]
        public void no_attribute_class_and_no_where_clause()
        {
            var companies = from c in ExcelRepository.GetSheet<Company>("")
                            select c;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void field_used_in_where_clause()
        {
            string desiredName = "Paul";
            var companies = from p in ExcelRepository.GetSheet<Company>("")
                            where p.Name == desiredName
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            Assert.AreEqual("SELECT * FROM [Sheet1$] Where (Name = ?)", GetSQLStatement());
            Assert.AreEqual("Paul", GetSQLParameters()[0]);
        }

        [Test]
        public void where_equals()
        {
            var companies = from p in ExcelRepository.GetSheet<Company>("")
                         where p.Name == "Paul"
                         select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] Where ({0} = ?)", GetSQLFieldName("Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("Paul", GetSQLParameters()[0]);
        }

        [Test]
        public void where_not_equal()
        {
            var companies = from p in ExcelRepository.GetSheet<Company>("")
                            where p.Name != "Paul"
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] Where ({0} <> ?)", GetSQLFieldName("Name"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("Paul", GetSQLParameters()[0]);
        }

        [Test]
        public void where_greater_than()
        {
            var companies = from p in ExcelRepository.GetSheet<Company>("")
                            where p.EmployeeCount > 25
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] Where ({0} > ?)", GetSQLFieldName("EmployeeCount"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("25", GetSQLParameters()[0]);
        }

        [Test]
        public void where_greater_than_or_equal()
        {
            var companies = from p in ExcelRepository.GetSheet<Company>("")
                            where p.EmployeeCount >= 25
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] Where ({0} >= ?)", GetSQLFieldName("EmployeeCount"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("25", GetSQLParameters()[0]);
        }

        [Test]
        public void where_lesser_than()
        {
            var companies = from p in ExcelRepository.GetSheet<Company>("")
                            where p.EmployeeCount < 25
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] Where ({0} < ?)", GetSQLFieldName("EmployeeCount"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("25", GetSQLParameters()[0]);
        }

        [Test]
        public void where_lesser_than_or_equal()
        {
            var companies = from p in ExcelRepository.GetSheet<Company>("")
                            where p.EmployeeCount <= 25
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] Where ({0} <= ?)", GetSQLFieldName("EmployeeCount"));
            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("25", GetSQLParameters()[0]);
        }

        [Test]
        public void where_and()
        {
            var companies = from p in ExcelRepository.GetSheet<Company>("")
                            where p.EmployeeCount > 5 && p.CEO == "Paul"
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] Where (({0} > ?) AND ({1} = ?))", 
                                                GetSQLFieldName("EmployeeCount"),
                                                GetSQLFieldName("CEO"));
            string[] parameters = GetSQLParameters();

            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("5", parameters[0]);
            Assert.AreEqual("Paul", parameters[1]);
        }

        [Test]
        public void where_or()
        {
            var companies = from p in ExcelRepository.GetSheet<Company>("")
                            where p.EmployeeCount > 5 || p.CEO == "Paul"
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] Where (({0} > ?) OR ({1} = ?))",
                                                GetSQLFieldName("EmployeeCount"),
                                                GetSQLFieldName("CEO"));
            string[] parameters = GetSQLParameters();

            Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("5", parameters[0]);
            Assert.AreEqual("Paul", parameters[1]);
        }

        [Test]
        public void where_new_datetime()
        {
            var companies = from p in ExcelRepository.GetSheet<Company>("")
                            where p.StartDate == new DateTime(2008, 10, 9)
                            select p;

            try { companies.GetEnumerator(); }
            catch (OleDbException) { }
            string expectedSql = string.Format("SELECT * FROM [Sheet1$] Where ({0} <= ?)", GetSQLFieldName("StartDate"));
            //Assert.AreEqual(expectedSql, GetSQLStatement());
            Assert.AreEqual("2008-10-09", GetSQLParameters()[0]);
        }
    }
}
