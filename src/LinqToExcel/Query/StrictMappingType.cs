namespace LinqToExcel.Query
{
    /// <summary>
    /// Class property and worksheet mapping enforcemment type.
    /// </summary>
    public enum StrictMappingType
    {
        /// <summary>
        /// All worksheet columns must map to a class property; all class properties must map to a worksheet columm.
        /// </summary>
        Both,

        /// <summary>
        /// All class properties must map to a worksheet column; other worksheet columns are ignored.
        /// </summary>
        ClassStrict,

        /// <summary>
        /// No checks are made to enforce worksheet column or class property mappings.
        /// </summary>
        None,

        /// <summary>
        /// All worksheet columns must map to a class property; other class properties are ignored.
        /// </summary>
        WorksheetStrict
    }
}