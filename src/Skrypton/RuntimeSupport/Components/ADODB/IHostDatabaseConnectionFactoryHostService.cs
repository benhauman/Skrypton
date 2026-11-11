using System.Data.Common;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The host service that the built-in ADODB.Connection / ADODB.Recordset / ADODB.Command open their database connections through. The host
    /// decides which provider (SqlClient, OleDb, ..) a connection string maps to and may restrict which databases scripts can reach - there is
    /// no default, so opening a connection raises an error if the host has not registered one.
    /// </summary>
    public interface IHostDatabaseConnectionFactoryHostService
    {
        /// <summary>
        /// Create a connection for the ADO connection string (eg. "Provider=SQLOLEDB;Data Source=..;Initial Catalog=..") - it may be returned open or
        /// closed (it is opened if it is closed)
        /// </summary>
        DbConnection CreateAndOpenDatabaseConnectionString(string connectionString, string userName, string password);
    }
}
