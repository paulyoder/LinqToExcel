using LinqToExcel.Exceptions;
using System.Collections.Generic;

namespace LinqToExcel
{
    /// <summary>
    /// Implement this interface to bypass the default thrown exception
    /// on a field parse error. All exceptions will instead be placed in
    /// this list. Be aware that you will still get a typed row back but
    /// failed column values should be untrusted.
    /// </summary>
    public interface IAllowFieldTypeConversionExceptions
    {
        IList<ExcelException> FieldTypeConversionExceptions { get; }
    }
}
