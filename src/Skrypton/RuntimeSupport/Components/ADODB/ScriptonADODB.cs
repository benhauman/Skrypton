using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The built-in ADODB.Connection. Connections are created by the host's IHostDatabaseConnectionFactoryHostService (there is no default).
    /// </summary>
    [SourceClassName("Connection")] // TypeName(CreateObject("ADODB.Connection"))
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonADODBConnection : ScriptonIReflectOnClrType
    {
        public const string ProgId = "ADODB.Connection";

        private readonly IRuntimeHost _runtimeHost;
        private DbConnection? _connection;
        private DbTransaction? _transaction;

        public ScriptonADODBConnection(IRuntimeHost runtimeHost) : base(typeof(ScriptonADODBConnection))
        {
            _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
        }

        [DispId(0)]
        public string ConnectionString { get; set; } = "";

        [DispId(2)]
        public int CommandTimeout { get; set; } = 30;

        [DispId(3)]
        public int ConnectionTimeout { get; set; } = 15;

        /// <summary>
        /// 0 (adStateClosed) or 1 (adStateOpen)
        /// </summary>
        [DispId(1006)]
        public int State => ((_connection != null) && (_connection.State != ConnectionState.Closed)) ? 1 : 0;

        [DispId(10)]
        public void Open(
            [Optional, DefaultParameterValue("")] string? ConnectionString,
            [Optional, DefaultParameterValue("")] string? UserID,
            [Optional, DefaultParameterValue("")] string? Password,
            [Optional, DefaultParameterValue(-1)] int Options)
        {
            if (State == 1)
                throw new InvalidOperationException("ADODB.Connection.Open: the connection is already open");
            if (!string.IsNullOrEmpty(ConnectionString))
                this.ConnectionString = ConnectionString!;
            if (string.IsNullOrWhiteSpace(this.ConnectionString))
                throw new InvalidOperationException("ADODB.Connection.Open: no connection string was specified");

            IHostDatabaseConnectionFactoryHostService factory = _runtimeHost.TryGetRuntimeHostService<IHostDatabaseConnectionFactoryHostService>()
                ?? throw new InvalidOperationException($"ADODB.Connection.Open: opening database connections requires the host to register an {nameof(IHostDatabaseConnectionFactoryHostService)}");
            DbConnection connection = factory.CreateAndOpenDatabaseConnectionString(this.ConnectionString, UserID ?? "", Password ?? "")
                ?? throw new InvalidOperationException($"ADODB.Connection.Open: the {nameof(IHostDatabaseConnectionFactoryHostService)} did not return a connection");
            if (connection.State == ConnectionState.Closed)
                connection.Open();
            _connection = connection;
        }

        [DispId(5)]
        public void Close()
        {
            if (State == 0)
                throw new InvalidOperationException("ADODB.Connection.Close: the connection is not open");
            _transaction?.Dispose();
            _transaction = null;
            _connection!.Dispose();
            _connection = null;
        }

        /// <summary>
        /// Execute(CommandText, [RecordsAffected], [Options]) - returns a Recordset (closed if the statement returns no rows)
        /// </summary>
        [DispId(6)]
        public object Execute(
            [In, MarshalAs(UnmanagedType.BStr)] string CommandText,
            [Optional] out object RecordsAffected,
            [In, Optional, DefaultParameterValue(-1)] int Options)
        {
            using DbCommand command = CreateCommand(CommandText, ((Options & 4) == 4) ? CommandType.StoredProcedure : CommandType.Text, CommandTimeout);
            ScriptonADODBRecordset recordset = ScriptonADODBRecordset.FromCommand(_runtimeHost, command, out int recordsAffected);
            RecordsAffected = recordsAffected;
            return recordset;
        }

        [DispId(7)]
        public int BeginTrans()
        {
            if (_transaction != null)
                throw new InvalidOperationException("ADODB.Connection.BeginTrans: nested transactions are not supported");
            _transaction = OpenConnection("BeginTrans").BeginTransaction();
            return 1;
        }

        [DispId(8)]
        public void CommitTrans()
        {
            (_transaction ?? throw new InvalidOperationException("ADODB.Connection.CommitTrans: there is no transaction")).Commit();
            _transaction.Dispose();
            _transaction = null;
        }

        [DispId(9)]
        public void RollbackTrans()
        {
            (_transaction ?? throw new InvalidOperationException("ADODB.Connection.RollbackTrans: there is no transaction")).Rollback();
            _transaction.Dispose();
            _transaction = null;
        }

        internal DbCommand CreateCommand(string commandText, CommandType commandType, int commandTimeout)
        {
            if (string.IsNullOrWhiteSpace(commandText))
                throw new ArgumentException("ADODB: a command text is required", nameof(commandText));
            DbCommand command = OpenConnection("Execute").CreateCommand();
#pragma warning disable CA2100 // The SQL is the script's own command text (as with ADO); the host controls which database it reaches
            command.CommandText = commandText;
#pragma warning restore CA2100
            command.CommandType = commandType;
            if (commandTimeout > 0)
                command.CommandTimeout = commandTimeout;
            command.Transaction = _transaction;
            return command;
        }

        private DbConnection OpenConnection(string member)
        {
            return (State == 1) ? _connection! : throw new InvalidOperationException($"ADODB.Connection.{member}: the connection is not open");
        }
    }

    /// <summary>
    /// The built-in ADODB.Recordset - the rows are read when it is opened and then navigated in memory (adOpenStatic, read-only)
    /// </summary>
    [SourceClassName("Recordset")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [DefaultMember("Fields")] // rs("name") is rs.Fields("name")
    internal sealed class ScriptonADODBRecordset : ScriptonIReflectOnClrType
    {
        public const string ProgId = "ADODB.Recordset";

        private readonly IRuntimeHost _runtimeHost;
        private string[] _fieldNames = Array.Empty<string>();
        private List<object?[]> _rows = new List<object?[]>();
        private int _position;
        private bool _open;

        public ScriptonADODBRecordset(IRuntimeHost runtimeHost) : base(typeof(ScriptonADODBRecordset))
        {
            _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
        }

        internal static ScriptonADODBRecordset FromCommand(IRuntimeHost runtimeHost, DbCommand command, out int recordsAffected)
        {
            var recordset = new ScriptonADODBRecordset(runtimeHost);
            recordset.Load(command, out recordsAffected);
            return recordset;
        }

        private void Load(DbCommand command, out int recordsAffected)
        {
            using DbDataReader reader = command.ExecuteReader();
            _fieldNames = Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToArray();
            _rows = new List<object?[]>();
            while (reader.Read())
            {
                var values = new object?[reader.FieldCount];
                reader.GetValues(values!);
                _rows.Add(values);
            }
            recordsAffected = reader.RecordsAffected;
            _position = 0;
            _open = reader.FieldCount > 0; // a statement that returns no rows gives a closed Recordset (as with ADO)
        }

        /// <summary>
        /// Open(Source, ActiveConnection, [CursorType], [LockType], [Options]) - Source is the SQL, ActiveConnection a Connection or a connection string
        /// </summary>
        [DispId(1022)]
        public void Open(
            [Optional] object? Source,
            [Optional] object? ActiveConnection,
            [Optional, DefaultParameterValue(-1)] int CursorType,
            [Optional, DefaultParameterValue(-1)] int LockType,
            [Optional, DefaultParameterValue(-1)] int Options)
        {
            if (_open)
                throw new InvalidOperationException("ADODB.Recordset.Open: the recordset is already open");
            string sql = Convert.ToString(Unwrap(Source), CultureInfo.InvariantCulture) ?? "";

            ScriptonADODBConnection connection;
            switch (Unwrap(ActiveConnection))
            {
                case ScriptonADODBConnection existing:
                    connection = existing;
                    break;
                case string connectionString:
                    connection = new ScriptonADODBConnection(_runtimeHost);
                    connection.Open(connectionString, "", "", -1);
                    break;
                default:
                    throw new ArgumentException("ADODB.Recordset.Open: ActiveConnection must be an ADODB.Connection or a connection string", nameof(ActiveConnection));
            }

            using DbCommand command = connection.CreateCommand(sql, ((Options & 4) == 4) ? CommandType.StoredProcedure : CommandType.Text, connection.CommandTimeout);
            Load(command, out _);
            _open = true;
        }

        [DispId(1014)]
        public void Close()
        {
            if (!_open)
                throw new InvalidOperationException("ADODB.Recordset.Close: the recordset is not open");
            _open = false;
            _rows = new List<object?[]>();
            _fieldNames = Array.Empty<string>();
        }

        /// <summary>
        /// 0 (adStateClosed) or 1 (adStateOpen)
        /// </summary>
        [DispId(1006)]
        public int State => _open ? 1 : 0;

        [DispId(1007)]
        public bool EOF => OpenRows("EOF").Count == 0 || (_position >= _rows.Count);

        [DispId(1002)]
        public bool BOF => OpenRows("BOF").Count == 0 || (_position < 0);

        [DispId(1005)]
        public int RecordCount => OpenRows("RecordCount").Count;

        /// <summary>
        /// The 1-based position of the current record
        /// </summary>
        [DispId(1003)]
        public int AbsolutePosition => _position + 1;

        [DispId(1018)]
        public void MoveNext()
        {
            if (EOF)
                throw new InvalidOperationException("ADODB.Recordset.MoveNext: either BOF or EOF is True (there is no current record)");
            _position++;
        }

        [DispId(1019)]
        public void MovePrevious()
        {
            if (BOF)
                throw new InvalidOperationException("ADODB.Recordset.MovePrevious: either BOF or EOF is True (there is no current record)");
            _position--;
        }

        [DispId(1020)]
        public void MoveFirst()
        {
            OpenRows("MoveFirst");
            _position = 0;
        }

        [DispId(1021)]
        public void MoveLast()
        {
            _position = OpenRows("MoveLast").Count - 1;
        }

        /// <summary>
        /// rs.Fields - the Fields collection of the current record
        /// </summary>
        [DispId(0)]
        public object Fields() => new ScriptonADODBFields(this);

        /// <summary>
        /// rs("name") / rs(index) / rs.Fields("name") - the Field of the current record
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public object Fields(object Index) => GetField(Index);

        internal int FieldCount => OpenRows("Fields").Count >= 0 ? _fieldNames.Length : 0;

        internal ScriptonADODBField GetField(object index)
        {
            OpenRows("Fields");
            index = Unwrap(index)!;
            int fieldIndex;
            if (index is string name)
            {
                fieldIndex = Array.FindIndex(_fieldNames, n => n.Equals(name, StringComparison.OrdinalIgnoreCase));
                if (fieldIndex < 0)
                    throw new ArgumentException($"ADODB.Recordset: the field '{name}' does not exist (the fields are: {string.Join(", ", _fieldNames)})", nameof(index));
            }
            else
            {
                fieldIndex = Convert.ToInt32(index, CultureInfo.InvariantCulture);
                if ((fieldIndex < 0) || (fieldIndex >= _fieldNames.Length))
                    throw new ArgumentOutOfRangeException(nameof(index), $"ADODB.Recordset: there is no field {fieldIndex} (there are {_fieldNames.Length})");
            }
            return new ScriptonADODBField(this, fieldIndex, _fieldNames[fieldIndex]);
        }

        internal object? GetValue(int fieldIndex)
        {
            if (EOF || BOF)
                throw new InvalidOperationException("ADODB.Field.Value: either BOF or EOF is True (there is no current record)");
            return _rows[_position][fieldIndex];
        }

        private List<object?[]> OpenRows(string member)
        {
            return _open ? _rows : throw new InvalidOperationException($"ADODB.Recordset.{member}: the recordset is closed");
        }

        private static object? Unwrap(object? value) => (value is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : value;
    }

    [SourceClassName("Fields")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [DefaultMember("Item")]
    internal sealed class ScriptonADODBFields : ScriptonIReflectOnClrType, IEnumerable
    {
        private readonly ScriptonADODBRecordset _recordset;

        public ScriptonADODBFields(ScriptonADODBRecordset recordset) : base(typeof(ScriptonADODBFields))
        {
            _recordset = recordset ?? throw new ArgumentNullException(nameof(recordset));
        }

        [DispId(1)]
        public int Count => _recordset.FieldCount;

        [DispId(0)]
        [IsDefault]
        public object Item(object Index) => _recordset.GetField(Index);

        [DispId(-4)]
        public IEnumerator GetEnumerator() => Enumerable.Range(0, Count).Select(i => (object)_recordset.GetField(i)).GetEnumerator();
    }

    [SourceClassName("Field")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [DefaultMember("Value")]
    internal sealed class ScriptonADODBField : ScriptonIReflectOnClrType
    {
        private readonly ScriptonADODBRecordset _recordset;
        private readonly int _index;

        public ScriptonADODBField(ScriptonADODBRecordset recordset, int index, string name) : base(typeof(ScriptonADODBField))
        {
            _recordset = recordset ?? throw new ArgumentNullException(nameof(recordset));
            _index = index;
            Name = name ?? "";
        }

        [DispId(1100)]
        public string Name { get; }

        /// <summary>
        /// The value of the field in the current record (a database NULL is VBScript Null)
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public object? Value => _recordset.GetValue(_index);
    }

    /// <summary>
    /// The built-in ADODB.Command
    /// </summary>
    [SourceClassName("Command")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [DefaultMember("Parameters")]
    internal sealed class ScriptonADODBCommand : ScriptonIReflectOnClrType
    {
        public const string ProgId = "ADODB.Command";

        private readonly IRuntimeHost _runtimeHost;
        private readonly ScriptonADODBParameters _parameters;

        public ScriptonADODBCommand(IRuntimeHost runtimeHost) : base(typeof(ScriptonADODBCommand))
        {
            _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
            _parameters = new ScriptonADODBParameters();
        }

        /// <summary>
        /// An ADODB.Connection (Set cmd.ActiveConnection = conn) or a connection string
        /// </summary>
        [DispId(1)]
        public object? ActiveConnection { get; set; }

        [DispId(2)]
        public string CommandText { get; set; } = "";

        [DispId(3)]
        public int CommandTimeout { get; set; } = 30;

        /// <summary>
        /// 1 (adCmdText, default), 4 (adCmdStoredProc)
        /// </summary>
        [DispId(7)]
        public int CommandType { get; set; } = 1;

        [DispId(0)]
        public object Parameters() => _parameters;

        [DispId(0)]
        [IsDefault]
        public object Parameters(object Index) => _parameters.Item(Index);

        [DispId(6)]
#pragma warning disable CA1822 // Bound by name from scripts, so it must be an instance member
        public object CreateParameter(
            [Optional, DefaultParameterValue("")] string? Name,
            [Optional, DefaultParameterValue(0)] int Type,
            [Optional, DefaultParameterValue(1)] int Direction,
            [Optional, DefaultParameterValue(0)] int Size,
            [Optional, DefaultParameterValue(null)] object? Value)
        {
            return new ScriptonADODBParameter { Name = Name ?? "", Type = Type, Direction = Direction, Size = Size, Value = Value };
        }
#pragma warning restore CA1822

        /// <summary>
        /// Execute([RecordsAffected], [Parameters], [Options]) - Parameters may be an array of values for the input parameters, in order
        /// </summary>
        [DispId(5)]
        public object Execute([Optional] out object RecordsAffected, [Optional] object? Parameters, [Optional, DefaultParameterValue(-1)] int Options)
        {
            ScriptonADODBConnection connection;
            switch ((ActiveConnection is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : ActiveConnection)
            {
                case ScriptonADODBConnection existing:
                    connection = existing;
                    break;
                case string connectionString:
                    connection = new ScriptonADODBConnection(_runtimeHost);
                    connection.Open(connectionString, "", "", -1);
                    break;
                default:
                    throw new InvalidOperationException("ADODB.Command.Execute: ActiveConnection must be set to an ADODB.Connection or a connection string");
            }

            if (Parameters is object[] values)
            {
                for (int i = 0; (i < values.Length) && (i < _parameters.Count); i++)
                    _parameters.GetParameter(i).Value = values[i];
            }

            System.Data.CommandType commandType = (CommandType == 4) || ((Options & 4) == 4) ? System.Data.CommandType.StoredProcedure : System.Data.CommandType.Text;
            using DbCommand command = connection.CreateCommand(CommandText, commandType, CommandTimeout);
            var dbParameters = new List<KeyValuePair<ScriptonADODBParameter, DbParameter>>();
            foreach (ScriptonADODBParameter parameter in _parameters)
            {
                DbParameter dbParameter = command.CreateParameter();
                parameter.ApplyTo(dbParameter);
                command.Parameters.Add(dbParameter);
                dbParameters.Add(new KeyValuePair<ScriptonADODBParameter, DbParameter>(parameter, dbParameter));
            }

            ScriptonADODBRecordset recordset = ScriptonADODBRecordset.FromCommand(_runtimeHost, command, out int recordsAffected);
            foreach (KeyValuePair<ScriptonADODBParameter, DbParameter> parameter in dbParameters.Where(p => p.Key.Direction != 1))
                parameter.Key.Value = parameter.Value.Value; // output / return values
            RecordsAffected = recordsAffected;
            return recordset;
        }
    }

    [SourceClassName("Parameters")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [DefaultMember("Item")]
    internal sealed class ScriptonADODBParameters : ScriptonIReflectOnClrType, IEnumerable
    {
        private readonly List<ScriptonADODBParameter> _parameters = new List<ScriptonADODBParameter>();

        public ScriptonADODBParameters() : base(typeof(ScriptonADODBParameters))
        {
        }

        [DispId(1)]
        public int Count => _parameters.Count;

        [DispId(3)]
        public void Append(object Object)
        {
            object? value = (Object is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : Object;
            _parameters.Add(value as ScriptonADODBParameter ?? throw new ArgumentException("ADODB.Parameters.Append: the object is not a Parameter (use Command.CreateParameter)", nameof(Object)));
        }

        [DispId(2)]
        public void Delete(object Index) => _parameters.Remove((ScriptonADODBParameter)Item(Index));

        /// <summary>
        /// By 0-based index or by name
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public object Item(object Index)
        {
            if (Index is string name)
            {
                return _parameters.FirstOrDefault(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    ?? throw new ArgumentException($"ADODB.Parameters: there is no parameter named '{name}'", nameof(Index));
            }
            return GetParameter(Convert.ToInt32(Index, CultureInfo.InvariantCulture));
        }

        internal ScriptonADODBParameter GetParameter(int index)
        {
            return ((index >= 0) && (index < _parameters.Count)) ? _parameters[index] : throw new ArgumentOutOfRangeException(nameof(index), $"ADODB.Parameters: there is no parameter {index} (there are {_parameters.Count})");
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator() => _parameters.ToList().GetEnumerator();
    }

    [SourceClassName("Parameter")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [DefaultMember("Value")]
    internal sealed class ScriptonADODBParameter : ScriptonIReflectOnClrType
    {
        public ScriptonADODBParameter() : base(typeof(ScriptonADODBParameter))
        {
        }

        [DispId(1)]
        public string Name { get; set; } = "";

        [DispId(0)]
        [IsDefault]
        public object? Value { get; set; }

        /// <summary>
        /// The ADO DataTypeEnum value (eg. 3 adInteger, 200 adVarChar, 202 adVarWChar, 7 adDate, 11 adBoolean)
        /// </summary>
        [DispId(2)]
        public int Type { get; set; }

        /// <summary>
        /// 1 adParamInput, 2 adParamOutput, 3 adParamInputOutput, 4 adParamReturnValue
        /// </summary>
        [DispId(3)]
        public int Direction { get; set; } = 1;

        [DispId(6)]
        public int Size { get; set; }

        internal void ApplyTo(DbParameter parameter)
        {
            if (!string.IsNullOrEmpty(Name))
                parameter.ParameterName = Name;
            parameter.Direction = Direction switch
            {
                2 => ParameterDirection.Output,
                3 => ParameterDirection.InputOutput,
                4 => ParameterDirection.ReturnValue,
                _ => ParameterDirection.Input
            };
            DbType? dbType = ToDbType(Type);
            if (dbType != null)
                parameter.DbType = dbType.Value;
            if (Size > 0)
                parameter.Size = Size;
            object? value = (Value is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : Value;
            parameter.Value = value ?? DBNull.Value;
        }

        private static DbType? ToDbType(int adoType) => adoType switch
        {
            16 => DbType.SByte,
            17 => DbType.Byte,
            2 => DbType.Int16,
            3 => DbType.Int32,
            20 => DbType.Int64,
            4 => DbType.Single,
            5 => DbType.Double,
            6 => DbType.Currency,
            14 or 131 or 139 => DbType.Decimal,
            11 => DbType.Boolean,
            72 => DbType.Guid,
            7 or 135 => DbType.DateTime,
            133 => DbType.Date,
            134 => DbType.Time,
            129 or 200 or 201 => DbType.AnsiString,
            8 or 130 or 202 or 203 => DbType.String,
            128 or 204 or 205 => DbType.Binary,
            _ => null
        };
    }

    /// <summary>
    /// The built-in ADODB.Stream (in memory). LoadFromFile / SaveToFile go through the host's IHostFileSystemHostService (or the process file
    /// system if the host has not registered one, as with Scripting.FileSystemObject).
    /// </summary>
    [SourceClassName("Stream")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonADODBStream : ScriptonIReflectOnClrType, IDisposable
    {
        public const string ProgId = "ADODB.Stream";

        private readonly IHostFileSystemHostService _fileSystem;
        private System.IO.MemoryStream? _buffer;

        public ScriptonADODBStream(IRuntimeHost runtimeHost) : base(typeof(ScriptonADODBStream))
        {
            if (runtimeHost == null)
                throw new ArgumentNullException(nameof(runtimeHost));
            _fileSystem = runtimeHost.TryGetRuntimeHostService<IHostFileSystemHostService>() ?? DefaultHostFileSystemHostService.Instance;
        }

        /// <summary>
        /// 1 (adTypeBinary) or 2 (adTypeText, default)
        /// </summary>
        [DispId(4)]
        public int Type { get; set; } = 2;

        /// <summary>
        /// The text encoding (default "Unicode", as with ADO)
        /// </summary>
        [DispId(21)]
        public string Charset { get; set; } = "Unicode";

        [DispId(1)]
        public int Size => (int)Buffer("Size").Length;

        [DispId(3)]
        public int Position
        {
            get => (int)Buffer("Position").Position;
            set => Buffer("Position").Position = value;
        }

        [DispId(2)]
        public bool EOS => Buffer("EOS").Position >= Buffer("EOS").Length;

        [DispId(10)]
        public void Open([Optional] object? Source, [Optional, DefaultParameterValue(0)] int Mode, [Optional, DefaultParameterValue(-1)] int Options, [Optional, DefaultParameterValue("")] string? UserName, [Optional, DefaultParameterValue("")] string? Password)
        {
            if (_buffer != null)
                throw new InvalidOperationException("ADODB.Stream.Open: the stream is already open");
            _buffer = new System.IO.MemoryStream();
        }

        [DispId(11)]
        public void Close()
        {
            Buffer("Close").Dispose();
            _buffer = null;
        }

        [DispId(13)]
        public void Write(object Buffer)
        {
            byte[] bytes = Buffer as byte[] ?? throw new ArgumentException("ADODB.Stream.Write: the buffer must be a Byte array", nameof(Buffer));
            this.Buffer("Write").Write(bytes, 0, bytes.Length);
        }

        [DispId(14)]
        public void WriteText(string Data, [Optional, DefaultParameterValue(0)] int Options)
        {
            string text = (Options == 1) ? (Data ?? "") + "\r\n" : (Data ?? ""); // 1: adWriteLine
            byte[] bytes = GetEncoding().GetBytes(text);
            Buffer("WriteText").Write(bytes, 0, bytes.Length);
        }

        [DispId(12)]
        public object Read([Optional, DefaultParameterValue(-1)] int NumBytes) => ReadBytes(NumBytes);

        [DispId(15)]
        public string ReadText([Optional, DefaultParameterValue(-1)] int NumChars)
        {
            string text = GetEncoding().GetString(ReadBytes(-1));
            return ((NumChars >= 0) && (NumChars < text.Length)) ? text.Substring(0, NumChars) : text;
        }

        [DispId(16)]
        public void SaveToFile(string FileName, [Optional, DefaultParameterValue(1)] int Options)
        {
            bool overwrite = Options == 2; // 1: adSaveCreateNotExist, 2: adSaveCreateOverWrite
            using System.IO.Stream target = _fileSystem.OpenTextFileWrite(FileName, createIfNotExists: true, overwriteIfExists: overwrite, append: false);
            Buffer("SaveToFile").WriteTo(target);
        }

        [DispId(19)]
        public void LoadFromFile(string FileName)
        {
            using System.IO.StreamReader reader = _fileSystem.OpenTextFileRead(FileName);
            System.IO.MemoryStream buffer = Buffer("LoadFromFile");
            buffer.SetLength(0);
            reader.BaseStream.CopyTo(buffer);
            buffer.Position = 0;
        }

        private byte[] ReadBytes(int count)
        {
            System.IO.MemoryStream buffer = Buffer("Read");
            int available = (int)(buffer.Length - buffer.Position);
            int toRead = ((count < 0) || (count > available)) ? available : count;
            byte[] bytes = new byte[toRead];
            int read = buffer.Read(bytes, 0, toRead);
            return (read == toRead) ? bytes : bytes.Take(read).ToArray();
        }

        private Encoding GetEncoding()
        {
            if (string.IsNullOrEmpty(Charset) || Charset.Equals("Unicode", StringComparison.OrdinalIgnoreCase))
                return Encoding.Unicode;
            return Encoding.GetEncoding(Charset);
        }

        public void Dispose()
        {
            _buffer?.Dispose();
            _buffer = null;
        }

        private System.IO.MemoryStream Buffer(string member) => _buffer ?? throw new InvalidOperationException($"ADODB.Stream.{member}: the stream is not open");
    }
}
