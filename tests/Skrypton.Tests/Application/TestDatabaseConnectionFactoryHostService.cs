#nullable enable
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using Skrypton.RuntimeSupport.Components;

namespace Skrypton.Tests.Application
{
    /// <summary>
    /// A test IHostDatabaseConnectionFactoryHostService: every connection answers every statement from the results function (an empty result
    /// with no rows by default) and records the executed statements - no database is needed
    /// </summary>
    internal sealed class TestDatabaseConnectionFactoryHostService : IHostDatabaseConnectionFactoryHostService
    {
        private readonly Func<string, DataTable?> _results;

        public TestDatabaseConnectionFactoryHostService(Func<string, DataTable?>? results = null)
        {
            _results = results ?? (_ => null);
        }

        public List<string> ConnectionStrings { get; } = new List<string>();
        public List<TestDbCommand> ExecutedCommands { get; } = new List<TestDbCommand>();

        public DbConnection CreateAndOpenDatabaseConnectionString(string connectionString, string userName, string password)
        {
            ConnectionStrings.Add(connectionString);
            return new TestOleDbConnection(connectionString, this);
        }

        internal DataTable GetResult(TestDbCommand command)
        {
            ExecutedCommands.Add(command);
            return _results(command.CommandText) ?? new DataTable();
        }
    }

    internal sealed class TestOleDbConnection : DbConnection
    {
        private readonly TestDatabaseConnectionFactoryHostService _service;
        private ConnectionState _state = ConnectionState.Closed;

        public TestOleDbConnection(string connectionString, TestDatabaseConnectionFactoryHostService service)
        {
            ConnectionString = connectionString;
            _service = service ?? throw new ArgumentNullException(nameof(service));
        }

        protected override DbTransaction BeginDbTransaction(IsolationLevel isolationLevel) => new TestDbTransaction(this, isolationLevel);
        public override void ChangeDatabase(string databaseName) { }
        public override void Close() => _state = ConnectionState.Closed;
        public override void Open() => _state = ConnectionState.Open;
#pragma warning disable CS8765 // Nullability of type of parameter doesn't match overridden member
        public override string ConnectionString { get; set; }
#pragma warning restore CS8765
        public override string Database => "";
        public override ConnectionState State => _state;
        public override string DataSource => "";
        public override string ServerVersion => "";
        protected override DbCommand CreateDbCommand() => new TestDbCommand(_service) { Connection = this };
    }

    internal sealed class TestDbTransaction : DbTransaction
    {
        public TestDbTransaction(DbConnection connection, IsolationLevel isolationLevel)
        {
            DbConnection = connection;
            IsolationLevel = isolationLevel;
        }

        public string State { get; private set; } = "open";
        public override IsolationLevel IsolationLevel { get; }
        protected override DbConnection DbConnection { get; }
        public override void Commit() => State = "committed";
        public override void Rollback() => State = "rolled back";
    }

    internal sealed class TestDbCommand : DbCommand
    {
        private readonly TestDatabaseConnectionFactoryHostService _service;
        private readonly TestDbParameterCollection _parameters = new TestDbParameterCollection();

        public TestDbCommand(TestDatabaseConnectionFactoryHostService service)
        {
            _service = service;
        }

#pragma warning disable CS8765 // Nullability of type of parameter doesn't match overridden member
        public override string CommandText { get; set; } = "";
#pragma warning restore CS8765
        public override int CommandTimeout { get; set; }
        public override CommandType CommandType { get; set; }
        public override bool DesignTimeVisible { get; set; }
        public override UpdateRowSource UpdatedRowSource { get; set; }
        protected override DbConnection? DbConnection { get; set; }
        protected override DbParameterCollection DbParameterCollection => _parameters;
        protected override DbTransaction? DbTransaction { get; set; }

        public IReadOnlyList<DbParameter> ExecutedParameters => _parameters.Items;

        public override void Cancel() { }
        public override int ExecuteNonQuery() => _service.GetResult(this).Rows.Count;
        public override object? ExecuteScalar() { DataTable result = _service.GetResult(this); return result.Rows.Count > 0 ? result.Rows[0][0] : null; }
        public override void Prepare() { }
        protected override DbParameter CreateDbParameter() => new TestDbParameter();
        protected override DbDataReader ExecuteDbDataReader(CommandBehavior behavior)
        {
            DataTable result = _service.GetResult(this);
            foreach (DbParameter parameter in _parameters.Items.Where(p => p.Direction != ParameterDirection.Input))
                parameter.Value = "out:" + parameter.ParameterName;
            return result.CreateDataReader();
        }
    }

    internal sealed class TestDbParameter : DbParameter
    {
        public override DbType DbType { get; set; }
        public override ParameterDirection Direction { get; set; } = ParameterDirection.Input;
        public override bool IsNullable { get; set; }
#pragma warning disable CS8765 // Nullability of type of parameter doesn't match overridden member
        public override string ParameterName { get; set; } = "";
        public override string SourceColumn { get; set; } = "";
#pragma warning restore CS8765
        public override object? Value { get; set; }
        public override bool SourceColumnNullMapping { get; set; }
        public override int Size { get; set; }
        public override void ResetDbType() { }
    }

    internal sealed class TestDbParameterCollection : DbParameterCollection
    {
        internal List<DbParameter> Items { get; } = new List<DbParameter>();

        public override int Count => Items.Count;
        public override object SyncRoot => Items;
        public override int Add(object value) { Items.Add((DbParameter)value); return Items.Count - 1; }
        public override void AddRange(Array values) { foreach (object value in values) Add(value); }
        public override void Clear() => Items.Clear();
        public override bool Contains(object value) => Items.Contains((DbParameter)value);
        public override bool Contains(string value) => IndexOf(value) >= 0;
        public override void CopyTo(Array array, int index) => ((ICollection)Items).CopyTo(array, index);
        public override IEnumerator GetEnumerator() => Items.GetEnumerator();
        public override int IndexOf(object value) => Items.IndexOf((DbParameter)value);
        public override int IndexOf(string parameterName) => Items.FindIndex(p => p.ParameterName == parameterName);
        public override void Insert(int index, object value) => Items.Insert(index, (DbParameter)value);
        public override void Remove(object value) => Items.Remove((DbParameter)value);
        public override void RemoveAt(int index) => Items.RemoveAt(index);
        public override void RemoveAt(string parameterName) => Items.RemoveAt(IndexOf(parameterName));
        protected override DbParameter GetParameter(int index) => Items[index];
        protected override DbParameter GetParameter(string parameterName) => Items[IndexOf(parameterName)];
        protected override void SetParameter(int index, DbParameter value) => Items[index] = value;
        protected override void SetParameter(string parameterName, DbParameter value) => Items[IndexOf(parameterName)] = value;
    }
}
