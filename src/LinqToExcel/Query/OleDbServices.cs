namespace LinqToExcel.Query
{
    /// <summary>
    /// Describes which services the OLE DB connection will use.
    /// </summary>
    /// <remarks>
    /// This allows you to change the OLE DB Services value in the connection string, which among other
    /// features, will allow you to opt out of implicit transactions (e.g. those created using
    /// TransactionScope).
    ///
    /// <code>
    /// Services enabled                                | Value in connection string
    /// ============================================================================
    /// All services (the default)                      | "OLE DB Services = -1;"
    /// All services except pooling                     | "OLE DB Services = -2;"
    /// All services except pooling and auto-enlistment | "OLE DB Services = -4;"
    /// All services except client cursor               | "OLE DB Services = -5;"
    /// All services except client cursor and pooling   | "OLE DB Services = -6;"
    /// No services                                     | "OLE DB Services = 0;"
    /// </code>
    ///
    /// See https://msdn.microsoft.com/en-us/library/ms810829.aspx for more information.
    /// </remarks>
    public enum OleDbServices
    {
        /// <summary>
        /// This is the default value for OLE DB Services in connection strings where it
        /// is not explicitly specified. Sets OLE DB Services to -1 in the connection string.
        /// </summary>
        AllServices = -1,

        /// <summary>
        /// Sets OLE DB Services to -2 in the connection string.
        /// </summary>
        AllServicesExceptPooling = -2,

        /// <summary>
        /// This will disable auto-enlistment in TransactionScope.
        /// Sets OLE DB Services to -4 in the connection string.
        /// </summary>
        AllServicesExceptPoolingAndAutoEnlistment = -4,

        /// <summary>
        /// Sets OLE DB Services to -5 in the connection string.
        /// </summary>
        AllServicesExceptClientCursor = -5,

        /// <summary>
        /// Sets OLE DB Services to -6 in the connection string.
        /// </summary>
        AllServicesExceptClientCursorAndPooling = -6,

        /// <summary>
        /// Sets OLE DB Services to 0 in the connection string.
        /// </summary>
        NoServices = 0,
    }
}
