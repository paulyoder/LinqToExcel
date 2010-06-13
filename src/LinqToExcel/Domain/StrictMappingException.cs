using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel.Domain
{
    public class StrictMappingException : Exception
    {
        public StrictMappingException(string Message)
            : base(Message)
        { }

        public StrictMappingException(string formatMessage, params object[] args)
            : base(string.Format(formatMessage, args))
        { }
    }
}
