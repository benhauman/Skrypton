using System;
using System.Data.Common;
using System.Data.OleDb;
using Skrypton.RuntimeSupport.Components;

namespace Skrypton.Windows
{
    /// <summary>
    /// ADODB connections through the OLE DB providers of Windows - the ADO connection string is an OLE DB connection string ("Provider=MSOLEDBSQL;
    /// Data Source=..;Initial Catalog=..", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=x.accdb", ..); without a provider the ODBC provider is
    /// used, as with ADO
    /// </summary>
    public sealed class WindowsHostOleDbConnectionFactoryHostService : IHostDatabaseConnectionFactoryHostService
    {
        public DbConnection CreateAndOpenDatabaseConnectionString(string connectionString, string userName, string password)
        {
            var builder = new OleDbConnectionStringBuilder(connectionString ?? "");
            if (string.IsNullOrEmpty(builder.Provider))
                builder.Provider = "MSDASQL"; // ADO's default provider (ODBC)
            if (!string.IsNullOrEmpty(userName))
                builder["User ID"] = userName;
            if (!string.IsNullOrEmpty(password))
                builder["Password"] = password;
            var connection = new OleDbConnection(builder.ConnectionString);
            try
            {
                connection.Open();
                return connection;
            }
            catch
            {
                connection.Dispose();
                throw;
            }
        }
    }
}
