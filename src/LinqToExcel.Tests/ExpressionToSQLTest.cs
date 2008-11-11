using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MbUnit.Framework;
using log4net;
using System.Reflection;
using log4net.Appender;
using log4net.Core;

namespace LinqToExcel.Tests
{
    [TestFixture]
    public class ExpressionToSQLTest
    {
        private static ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
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

        private string GetSQLStatement()
        {
            LoggingEvent[] loggingEvents = _loggedEvents.GetEvents();
            foreach (LoggingEvent logEvent in loggingEvents)
            {
                string message = logEvent.RenderedMessage;
                if (message.Length > 5 && message.Substring(0, 4) == "SQL:")
                {
                    return logEvent.RenderedMessage.Substring(5);
                }
            }
            return "";
        }

        [Test]
        public void no_attribute_class_and_no_where_clause()
        {
            ExcelRepository<Company> repo = new ExcelRepository<Company>();
            var companies = from c in repo
                            select c;
            
            companies.GetEnumerator();
            Assert.AreEqual("SELECT * FROM [Sheet1$]", GetSQLStatement());
        }

        [Test]
        public void where_string_equals()
        {
            ExcelRepository<Company> repo = new ExcelRepository<Company>();
            var companies = from p in repo
                         where p.Name == "Paul"
                         select p;

            companies.GetEnumerator();
            Assert.AreEqual("SELECT * FROM [Sheet1$] Where (Name = 'Paul')", GetSQLStatement());
        }

        [Test]
        public void field_used_in_where_clause()
        {
            string desiredName = "Paul";
            ExcelRepository<Company> repo = new ExcelRepository<Company>();
            var companies = from p in repo
                            where p.Name == desiredName
                            select p;

            companies.GetEnumerator();
            Assert.AreEqual("SELECT * FROM [Sheet1$] Where (Name = 'Paul')", GetSQLStatement());
        }

        [Test]
        public void where_int_equals()
        {
            ExcelRepository<Company> repo = new ExcelRepository<Company>();
            var companies = from p in repo
                         where p.EmployeeCount == 25
                         select p;

            companies.GetEnumerator();
            Assert.AreEqual("SELECT * FROM [Sheet1$] Where (EmployeeCount = 25)", GetSQLStatement());
        }

        [Test]
        public void where_int_greater()
        {
            ExcelRepository<Company> repo = new ExcelRepository<Company>();
            var companies = from p in repo
                            where p.EmployeeCount > 25
                            select p;

            companies.GetEnumerator();
            Assert.AreEqual("SELECT * FROM [Sheet1$] Where (EmployeeCount > 25)", GetSQLStatement());
        }

        [Test]
        public void where_int_greater_than_or_equal()
        {
            ExcelRepository<Company> repo = new ExcelRepository<Company>();
            var companies = from p in repo
                            where p.EmployeeCount >= 25
                            select p;

            companies.GetEnumerator();
            Assert.AreEqual("SELECT * FROM [Sheet1$] Where (EmployeeCount >= 25)", GetSQLStatement());
        }
    }
}
