using System;
using System.Data;
using System.Data.Common;
using Skrypton.Tests.RuntimeSupport.Implementations;

namespace Skrypton.Tests.Application
{
    internal sealed class TestDatabaseConnectionFactoryHostService : IHostDatabaseConnectionFactoryHostService
    {
        public DbConnection CreateAndOpenDatabaseConnectionString(string connectionString, string userName, string password)
        {
            return new MyOleDbConnection(connectionString);
            //throw new NotImplementedException($"connectionString:'{connectionString}', user:'{userName}', pwd:{password}");
        }
    }

    internal sealed class MyOleDbConnection : DbConnection
    {
        private readonly string _connectionString;

        public MyOleDbConnection(string connectionString)
        {
            _connectionString = connectionString;
        }

        protected override DbTransaction BeginDbTransaction(IsolationLevel isolationLevel)
        {
            throw new NotImplementedException();
        }

        public override void ChangeDatabase(string databaseName)
        {
            throw new NotImplementedException();
        }

        public override void Close()
        {
            if (_state == ConnectionState.Closed)
                throw new InvalidOperationException("Connection is already closed.");
            if (_state == ConnectionState.Open)
                throw new InvalidOperationException("Connection is not open.");
            _state = ConnectionState.Closed;
        }

        private ConnectionState _state =  ConnectionState.Closed;
        public override void Open()
        {
            if (_state == ConnectionState.Open)
                throw new InvalidOperationException("Connection is already open.");
            _state = ConnectionState.Open;
        }

        public override string ConnectionString { get; set; }
        public override string Database { get; }
        public override ConnectionState State { get; }
        public override string DataSource { get; }
        public override string ServerVersion { get; }

        protected override DbCommand CreateDbCommand()
        {
            throw new NotImplementedException();
        }
    }
}