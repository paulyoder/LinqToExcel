using System.Collections.Generic;
using log4net.Appender;
using log4net.Core;
using System.Linq;

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
            var loggingEvents = _loggedEvents.GetEvents();
            foreach (LoggingEvent logEvent in loggingEvents)
            {
                if (logEvent.LoggerName == "LinqToExcel.SQL")
                    return logEvent.RenderedMessage.Split(";".ToCharArray())[0];
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
            var loggingEvents = _loggedEvents.GetEvents();
            var parameters = new List<string>();
            foreach (LoggingEvent logEvent in loggingEvents)
            {
                if (logEvent.LoggerName == "LinqToExcel.SQL")
                {
                    var splitMessage = logEvent.RenderedMessage.Split(";".ToCharArray());
                    for (var i = 1; i < splitMessage.Length - 1; i++)
                    {
                        parameters.Add(
                            splitMessage[i]
                                .Split("=".ToCharArray())
                                .Last()
                                .Replace("'", "")
                                .Substring(1));
                    }
                }
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
            var loggingEvents = _loggedEvents.GetEvents();
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
            var conProps = GetConnectionString().Split(";".ToCharArray());
            foreach (string conProp in conProps)
            {
                if (conProp.Substring(0, 11) == "Data Source")
                    return conProp.Substring(12);
            }
            return "";
        }

        protected string GetExtendedProperties()
        {
            var conString = GetConnectionString();
            var location = conString.IndexOf("Extended Properties=");
            return conString.Substring(location + 20);
        }
    }
}
