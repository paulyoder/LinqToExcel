using System;

namespace LinqToExcel.Attributes
{
    /// <summary>
    /// Ignores attribute during column map creation. Allows property to be safely ignored
    /// when using StrictMappingType
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
    public sealed class ExcelIgnore : Attribute
    {
        public ExcelIgnore()
        { }
    }
}
