// copyright(c) 2016 Stephen Workman (workman.stephen@gmail.com)

using System;

namespace LinqToExcel.Logging {
   public interface ILogProvider {
      bool IsDebugEnabled { get; }
      void Debug(Object message);
      void DebugFormat(String format, Object arg);
      void Error(Object message, Exception exception);
      void WarnFormat(String format, Object arg0, Object arg1, Object arg2);
   }
}
