// copyright(c) 2016 Stephen Workman (workman.stephen@gmail.com)

using System;

using log4net;
using LinqToExcel.Logging;

namespace LinqToExcel.Tests {
   public class LogManagerFactory : ILogManagerFactory {
      public ILogProvider GetLogger(String name) {
         return new LogProvider(LogManager.GetLogger(name));
      }

      public ILogProvider GetLogger(Type type) {
         return new LogProvider(LogManager.GetLogger(type));
      }
   }

   public class LogProvider : ILogProvider {
      private ILog _iLog;

      public LogProvider(ILog iLog) {
         _iLog = iLog;
      }

      public Boolean IsDebugEnabled {
         get {
            if (_iLog != null)
               return _iLog.IsDebugEnabled;
            return false;
         }
      }

      public void Debug(Object message) {
         if (_iLog != null)
            _iLog.Debug(message);
      }

      public void DebugFormat(String format, Object arg) {
         if (_iLog != null)
            _iLog.DebugFormat(format, arg);
      }

      public void Error(Object message, Exception exception) {
         if (_iLog != null)
            _iLog.Error(message, exception);
      }

      public void WarnFormat(String format, Object arg0, Object arg1, Object arg2) {
         if (_iLog != null)
            _iLog.WarnFormat(format, arg0, arg1, arg2);
      }
   }
}
