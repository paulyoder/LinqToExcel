namespace LinqToExcel.Query
{
    /// <summary>
    /// Indicates how to treat leading and trailing spaces in string values.
    /// </summary>
    public enum TrimSpacesType
    {
        /// <summary>
        /// Do not perform any trimming.
        /// </summary>
        None, 
        
        /// <summary>
        /// Trim leading spaces from strings.
        /// </summary>
        Start, 
        
        /// <summary>
        /// Trim trailing spaces from strings.
        /// </summary>
        End, 
        
        /// <summary>
        /// Trim leading and trailing spaces from strings. 
        /// </summary>
        Both
    }
}
