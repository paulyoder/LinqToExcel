using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using log4net.Appender;
using log4net.Core;

namespace LinqToExcel.Tests
{
    public class SQLLogStatements_Helper
    {
        /// <summary>
        /// This memory appender catches all the logged messages 
        /// which the unit tests can use for their assertions
        /// </summary>
        protected MemoryAppender _loggedEvents;

        protected void InstantiateLogger()
        {
            _loggedEvents = new MemoryAppender();
            log4net.Config.BasicConfigurator.Configure(_loggedEvents);
        }

        protected void ClearLogEvents()
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
        protected string GetSQLStatement()
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
        protected string[] GetSQLParameters()
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
        protected string GetSQLFieldName(string columnName)
        {
            return string.Format("[{0}]", columnName);
        }

        /// <summary>
        /// Returns the connection string
        /// </summary>
        protected string GetConnectionString()
        {
            LoggingEvent[] loggingEvents = _loggedEvents.GetEvents();
            foreach (LoggingEvent logEvent in loggingEvents)
            {
                string message = logEvent.RenderedMessage;
                if (message.Length > 5 && message.Substring(0, 18) == "Connection String:")
                    return logEvent.RenderedMessage.Substring(19);
            }
            return "";
        }

        protected string GetDataSource()
        {
            string[] conProps = GetConnectionString().Split(";".ToCharArray());
            foreach (string conProp in conProps)
            {
                if (conProp.Substring(0, 11) == "Data Source")
                    return conProp.Substring(12);
            }
            return "";
        }

        protected string GetExtendedProperties()
        {
            string conString = GetConnectionString();
            int location = conString.IndexOf("Extended Properties=");
            return conString.Substring(location + 20);
        }
    }
}
