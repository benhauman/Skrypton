using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using Skrypton.RuntimeSupport.Attributes;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    public interface IHostDatabaseConnectionFactoryHostService
    {
        DbConnection CreateAndOpenDatabaseConnectionString(string connectionString, string userName, string password);
    }
}

namespace Skrypton.Tests.RuntimeSupport.Implementations.ADODB
{
    [SourceClassName("Connection")] // for TYPENAME(CreateObject("ADODB.Connection"))
    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyADODBConnection : IReflectOnClrType, Connection
    {
        // csproj =>  Add → Reference =>  COM → Select Microsoft ActiveX Data Objects x.x Library => Interop.ADODB.dll # // Interop.ADODB, Version=6.1.0.0, Culture=neutral, PublicKeyToken=null
        private readonly IServiceProvider _hostServices;
        public MyADODBConnection(IServiceProvider hostServices) : base(typeof(MyADODBConnection))
        {
            _hostServices = hostServices ?? throw new ArgumentNullException(nameof(hostServices));
        }

        [DispId(0)]
        public string ConnectionString { get; set; }

        [DispId(3)]
        public int ConnectionTimeout { get; set; }

        private DbConnection _currentOpenConnectionOrNull;
        [DispId(10)]
        public void Open(string connectionString = "", string UserID = "", string Password = "", int Options = -1)
        {
            if (_currentOpenConnectionOrNull != null && _currentOpenConnectionOrNull.State == System.Data.ConnectionState.Open)
            {
                throw new InvalidOperationException("Connection is already open.");
            }

            // provider=sqloledb;data source=MSSQLB;initial catalog=HLData
            // helplinedata
            // helplinedata
            // options:
            //   -1 → “Use the default behavior”
            //    0 → adConnectUnspecified
            //    1: adConnectUnspecified
            //    2: adAsyncConnect
            var svc = _hostServices.GetRequiredService<IHostDatabaseConnectionFactoryHostService>();

            _currentOpenConnectionOrNull = svc.CreateAndOpenDatabaseConnectionString(connectionString: string.IsNullOrEmpty(connectionString) ? ConnectionString : connectionString, userName: UserID, password: Password);
            if (_currentOpenConnectionOrNull.State == System.Data.ConnectionState.Closed)
            {
                _currentOpenConnectionOrNull.Open(); // or later (on demand)
            }
        }

        [DispId(5)]
        public void Close()
        {
            if (_currentOpenConnectionOrNull == null || _currentOpenConnectionOrNull.State == System.Data.ConnectionState.Closed)
            {
                throw new InvalidOperationException("Connection is already closed.");
            }

            _currentOpenConnectionOrNull.Close();
            _currentOpenConnectionOrNull = null;
        }

        //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(6)]
        //[return: MarshalAs(UnmanagedType.Interface)]
        public Recordset Execute([In][MarshalAs(UnmanagedType.BStr)] string CommandText
            , [Optional][MarshalAs(UnmanagedType.Struct)] out object RecordsAffected
            , [In] int Options = -1)
        {
            Console.WriteLine($"SQL:{CommandText}");
            RecordsAffected = 0;
            return new MyADODBRecordSet();
        }

    }

    [SourceClassName("Command")] // for TYPENAME(CreateObject("ADODB.Command"))
    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyADODBCommand : IReflectOnClrType, Command
    {
        // ADODB.Command

        private readonly Dictionary<int, string> _parametersIndexMap = new Dictionary<int, string>(); // In Microsoft ActiveX Data Objects, the Parameters collection is 0-based, not 1-based.
        internal readonly Dictionary<string, MyADODBParameter> _parametersList = new Dictionary<string, MyADODBParameter>(StringComparer.OrdinalIgnoreCase);

        //private readonly IServiceProvider _hostServices;
        internal MyADODBCommand()
        {
            //_hostServices = hostServices ?? throw new ArgumentNullException(nameof(hostServices));
        }

        public Connection ActiveConnection { get; set; }

        [DispId(2)]
        public string CommandText { get; set; }

        [DispId(7)]
        public int CommandType { get; set; } // CommandTypeEnum

        [DispId(0)]
        //public Parameters Parameters => new MyADODBParameters(this);
        public Parameters Parameters()
        {
            return new MyADODBParameters(this);
        }

        [DispId(0)]  // This makes it the default VBScript property
        public object Parameters(object key)
        {
            return ((MyADODBParameters)Parameters()).Item(key);
        }

        [DispId(6)]
        //[return: MarshalAs(UnmanagedType.Interface)]
        public Parameter CreateParameter([In, MarshalAs(UnmanagedType.BStr), DefaultParameterValue("")] string Name,
            [In, DefaultParameterValue((short)202)] short Type, // 202:DataTypeEnum.adEmpty
            [In, DefaultParameterValue((short)1)] short Direction, // 1: ParameterDirectionEnum.adParamInput
            [In, DefaultParameterValue(9)] int Size,
            [In, MarshalAs(UnmanagedType.Struct), Optional, DefaultParameterValue(null)] object Value)
        {
            return new MyADODBParameter() { Name = Name, Type = (DataTypeEnum)Type, Direction = (ParameterDirectionEnum)Direction, Size = Size, Value = Value };
        }
        [DispId(6)]
        //[return: MarshalAs(UnmanagedType.Interface)]
        public Parameter CreateParameter([DefaultParameterValue("")] string Name,
            [DefaultParameterValue((short)0)] short Type, // 0:DataTypeEnum.adEmpty
            [DefaultParameterValue((short)1)] short Direction // 1:ParameterDirectionEnum.adParamInput
        //[In, DefaultParameterValue(9)] int Size,
        //[In, MarshalAs(UnmanagedType.Struct), Optional, DefaultParameterValue(null)] object Value
            )
        {
            // Creating parameters is purely local to the Command object and does not require a database connection.
            //var svc = _hostServices.GetRequiredService<IHostDatabaseConnectionFactoryHostService>();
            //svc. // svc

            return new MyADODBParameter() { Name = Name, Type = (DataTypeEnum)Type, Direction = (ParameterDirectionEnum)Direction };
        }

        [DispId(5)]
        [return: MarshalAs(UnmanagedType.Interface)]
        public Recordset Execute([MarshalAs(UnmanagedType.Struct), Optional] out object RecordsAffected,
            [In, MarshalAs(UnmanagedType.Struct), Optional] ref object Parameters,// Optional parameter values (array or missing)
            [In] int Options = -1 // Execution flags (CommandTypeEnum, ExecuteOptionEnum) 1:adCmdText, 4:adCmdStoredProcm, 5:adExecuteNoRecords
            )
        {
            var rset = new MyADODBRecordSet();
            RecordsAffected = null; // Returns number of rows affected
            return rset;
        }

        internal void AppendParameter(MyADODBParameter prm)
        {
            _parametersList.Add(prm.Name, prm);
            _parametersIndexMap.Add(_parametersList.Count - 1, prm.Name); // 0 based
        }

        internal MyADODBParameter GetParameterByIndex(int index)
        {
            if (!_parametersIndexMap.TryGetValue(index, out string parameterName))
                throw new ArgumentException($"No parameter at index {index}", nameof(index));
            return _parametersList.TryGetValue(parameterName, out MyADODBParameter prm)
                ? prm
                : throw new ArgumentException($"No parameter at index {index} with name '{parameterName}'", nameof(index));
        }
    }

    [SourceClassName("Parameters")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyADODBParameters : IReflectOnClrType, ADODB.Parameters
    {
        private readonly MyADODBCommand _cmd;

        internal MyADODBParameters(MyADODBCommand cmd)
        {
            _cmd = cmd ?? throw new ArgumentNullException(nameof(cmd));
        }

        [DispId(1)]
        public int Count => _cmd._parametersList.Count;
        [DispId(-4)]
        IEnumerator _Collection.GetEnumerator()
        {
            return _cmd._parametersList.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _cmd._parametersList.Values.GetEnumerator();
        }

        [DispId(3)]
        public void Append(object Val)
        {
            var prm = (MyADODBParameter)Val ?? throw new ArgumentNullException(nameof(Val));
            _cmd.AppendParameter(prm);
        }


        //[DispId(0)]
        //public Parameter this[[In] object Index] // needed to access output/return-value parameters. see ParameterDirectionEnum
        //{
        //    get
        //    {
        //        int indexInt = Convert.ToInt32(Index);
        //        return _cmd.GetParameterByIndex(indexInt);
        //    }
        //    set
        //    {
        //        throw new NotSupportedException();
        //    }
        //}

        [DispId(0)]  // This makes it the default VBScript property
        [IsDefault] // needed for enumeration : dict[currentkey]
        public object Item(object key)
        {
            if (key == null) throw new ArgumentNullException(nameof(key));
            string k = key.ToString() ?? throw new ArgumentException("Key cannot be null", nameof(key));

            int indexInt = Convert.ToInt32(key);
            return _cmd.GetParameterByIndex(indexInt);
        }

        [DispId(0)]
        public void Item(object key, object value)
        {
            if (key == null) throw new ArgumentNullException(nameof(key));
            string k = key.ToString() ?? throw new ArgumentException("Key cannot be null", nameof(key));
            throw new NotSupportedException();
        }
    }

    [SourceClassName("Parameter")]
    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyADODBParameter : IReflectOnClrType, ADODB.Parameter
    {
        public MyADODBParameter()
        {
        }

        [DispId(1)]
        public string Name { get; set; }

        [DispId(0)]
        public object Value { get; set; }

        [DispId(2)]
        public DataTypeEnum Type { get; set; }

        [DispId(3)]
        public ParameterDirectionEnum Direction { get; set; }

        [DispId(6)]
        public int Size { get; set; }
    }

    //[ComImport]
    //[CoClass(typeof(ConnectionClass))]
    //[Guid("00001550-0000-0010-8000-00AA006D2EA4")]
    internal interface Connection : _Connection//, ConnectionEvents_Event
    {
    }

    //[ComImport]
    //[Guid("00001550-0000-0010-8000-00AA006D2EA4")]
    [DefaultMember("ConnectionString")] // +[DispId(0)] +[IsDefault]
    //[TypeLibType(4160)]
    internal interface _Connection : Connection15
    {
    }

    //[ComImport]
    [DefaultMember("ConnectionString")] // +[DispId(0)] +[IsDefault]
    //[Guid("00001515-0000-0010-8000-00AA006D2EA4")]
    //[TypeLibType(4176)]
    internal interface Connection15 //: _ADO
    {
        [DispId(0)]
        string ConnectionString
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(0)]
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(0)]
            [param: In]
            [param: MarshalAs(UnmanagedType.BStr)]
            set;
        }

        //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(10)]
        void Open([In][MarshalAs(UnmanagedType.BStr)] string ConnectionString = "", [In][MarshalAs(UnmanagedType.BStr)] string UserID = "", [In][MarshalAs(UnmanagedType.BStr)] string Password = "", [In] int Options = -1);

    }


    //[ComImport]
    //[Guid("986761E8-7269-4890-AA65-AD7C03697A6D")]
    //[CoClass(typeof(CommandClass))]
    internal interface Command : _Command
    {
    }

    //[ComImport]
    //[TypeLibType(4288)]
    //[DefaultMember("Parameters")]
    //[Guid("986761E8-7269-4890-AA65-AD7C03697A6D")]
    internal interface _Command : Command25
    {
    }

    //[ComImport]
    //[Guid("0000154E-0000-0010-8000-00AA006D2EA4")]
    //[TypeLibType(4304)]
    [DefaultMember("Parameters")] // +[DispId(0)] +[IsDefault]
    internal interface Command25 : Command15
    {
    }

    //[ComImport]
    //[Guid("00001508-0000-0010-8000-00AA006D2EA4")]
    [DefaultMember("Parameters")] // +[DispId(0)] +[IsDefault]
    //[TypeLibType(4304)]
    internal interface Command15 //: _ADO
    {
        [DispId(1)]
        Connection ActiveConnection
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1)]
            [return: MarshalAs(UnmanagedType.Interface)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1)]
            [param: In]
            [param: MarshalAs(UnmanagedType.Interface)]
            set;
        }
        [DispId(2)]
        string CommandText
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(2)]
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(2)]
            [param: In]
            [param: MarshalAs(UnmanagedType.BStr)]
            set;
        }
        [DispId(7)]
        int CommandType// see 'CommandTypeEnum'
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(7)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(7)]
            [param: In]
            set;
        }

        //[DispId(0)]
        //Parameters Parameters
        //{
        //    [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        //    [DispId(0)]
        //    [return: MarshalAs(UnmanagedType.Interface)]
        //    get;
        //}


        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(6)]
        [return: MarshalAs(UnmanagedType.Interface)]
        Parameter CreateParameter([In] string Name,// = "",
            [In] short Type,// = DataTypeEnum.adEmpty,
            [In] short Direction,// = ParameterDirectionEnum.adParamInput,
            [In] int Size,/// = 0,
            [Optional][In] object Value
            );

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(5)]
        [return: MarshalAs(UnmanagedType.Interface)]
        Recordset Execute([Optional][MarshalAs(UnmanagedType.Struct)] out object RecordsAffected, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object Parameters, [In] int Options = -1);

    }

    //[ComImport]
    //[Guid("00001556-0000-0010-8000-00AA006D2EA4")]
    //[CoClass(typeof(RecordsetClass))]
    internal interface Recordset : _Recordset//, RecordsetEvents_Event
    {
    }

    //[ComImport]
    //[Guid("0000150C-0000-0010-8000-00AA006D2EA4")]
    //[CoClass(typeof(ParameterClass))]
    internal interface Parameter : _Parameter
    {
    }

    //[ComImport]
    //[TypeLibType(4304)]
    //[Guid("0000150C-0000-0010-8000-00AA006D2EA4")]
    [DefaultMember("Value")]
    internal interface _Parameter// : _ADO
    {
        [DispId(1)]
        string Name
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1)]
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1)]
            [param: In]
            [param: MarshalAs(UnmanagedType.BStr)]
            set;
        }

        [DispId(0)]
        object Value
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(0)]
            [return: MarshalAs(UnmanagedType.Struct)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(0)]
            [param: In]
            [param: MarshalAs(UnmanagedType.Struct)]
            set;
        }

        [DispId(2)]
        DataTypeEnum Type
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(2)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(2)]
            [param: In]
            set;
        }
        [DispId(3)]
        ParameterDirectionEnum Direction
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(3)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(3)]
            [param: In]
            set;
        }

        [DispId(6)]
        int Size { get; set; }
    }

    //[Guid("0000052C-0000-0010-8000-00AA006D2EA4")]
    internal enum ParameterDirectionEnum
    {
        adParamUnknown,
        adParamInput,
        adParamOutput,
        adParamInputOutput,
        adParamReturnValue
    }


    //[Guid("0000051F-0000-0010-8000-00AA006D2EA4")]
    internal enum DataTypeEnum
    {
        adEmpty = 0,
        adTinyInt = 16,
        adSmallInt = 2,
        adInteger = 3,
        adBigInt = 20,
        adUnsignedTinyInt = 17,
        adUnsignedSmallInt = 18,
        adUnsignedInt = 19,
        adUnsignedBigInt = 21,
        adSingle = 4,
        adDouble = 5,
        adCurrency = 6,
        adDecimal = 14,
        adNumeric = 131,
        adBoolean = 11,
        adError = 10,
        adUserDefined = 132,
        adVariant = 12,
        adIDispatch = 9,
        adIUnknown = 13,
        adGUID = 72,
        adDate = 7,
        adDBDate = 133,
        adDBTime = 134,
        adDBTimeStamp = 135,
        adBSTR = 8,
        adChar = 129,
        adVarChar = 200,
        adLongVarChar = 201,
        adWChar = 130,
        adVarWChar = 202,
        adLongVarWChar = 203,
        adBinary = 128,
        adVarBinary = 204,
        adLongVarBinary = 205,
        adChapter = 136,
        adFileTime = 64,
        adPropVariant = 138,
        adVarNumeric = 139,
        adArray = 8192
    }

    //[ComImport]
    //[TypeLibType(4288)]
    //[Guid("0000150D-0000-0010-8000-00AA006D2EA4")]
    internal interface Parameters : _DynaCollection
    {
        [DispId(3)]
        void Append([In][MarshalAs(UnmanagedType.Struct)] object Val);
    }

    //[ComImport]
    //[TypeLibType(4288)]
    //[Guid("00000513-0000-0010-8000-00AA006D2EA4")]
    internal interface _DynaCollection : _Collection
    {
    }

    //[ComImport]
    //[TypeLibType(4288)]
    //[Guid("00000512-0000-0010-8000-00AA006D2EA4")]
    internal interface _Collection : System.Collections.IEnumerable
    {
        [DispId(1)]
        int Count
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1)]
            get;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(-4)]
        //[TypeLibFunc(1)]
        //[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
        new System.Collections.IEnumerator GetEnumerator();
    }

    //[Guid("0000052E-0000-0010-8000-00AA006D2EA4")]
    internal enum CommandTypeEnum
    {
        //[TypeLibVar(64)]
        adCmdUnspecified = -1,
        adCmdUnknown = 8,
        adCmdText = 1,
        adCmdTable = 2,
        adCmdStoredProc = 4,
        adCmdFile = 256,
        adCmdTableDirect = 512
    }

    internal interface _Recordset : Recordset21
    {
    }
    internal interface Recordset21 : Recordset20
    {
    }
    internal interface Recordset20 : Recordset15
    {
    }

    //[ComImport]
    //[TypeLibType(4304)]
    //[Guid("0000150E-0000-0010-8000-00AA006D2EA4")]
    [DefaultMember("Fields")] // +[DispId(0)] +[IsDefault]
    internal interface Recordset15 //: _ADO
    {
        /*
        [DispId(500)]
        new Properties Properties
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(500)]
            [return: MarshalAs(UnmanagedType.Interface)]
            get;
        }

        [DispId(1000)]
        new PositionEnum AbsolutePosition
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1000)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1000)]
            [param: In]
            set;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [SpecialName]
        [DispId(1001)]
        new void let_ActiveConnection([In][MarshalAs(UnmanagedType.Struct)] object pvar);

        [DispId(1001)]
        new object ActiveConnection
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1001)]
            [return: MarshalAs(UnmanagedType.Struct)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1001)]
            [param: In]
            [param: MarshalAs(UnmanagedType.IDispatch)]
            set;
        }

        [DispId(1002)]
        new bool BOF
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1002)]
            get;
        }

        [DispId(1003)]
        new object Bookmark
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1003)]
            [return: MarshalAs(UnmanagedType.Struct)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1003)]
            [param: In]
            [param: MarshalAs(UnmanagedType.Struct)]
            set;
        }

        [DispId(1004)]
        new int CacheSize
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1004)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1004)]
            [param: In]
            set;
        }

        [DispId(1005)]
        new CursorTypeEnum CursorType
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1005)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1005)]
            [param: In]
            set;
        }
        */
        [DispId(1006)]
        bool EOF
        {
            //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1006)]
            get;
        }

        /*
        [DispId(0)]
        Fields Fields
        {
            //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(0)]
            //[return: MarshalAs(UnmanagedType.Interface)]
            get;
        }
        */
        /*
        [DispId(1008)]
        new LockTypeEnum LockType
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1008)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1008)]
            [param: In]
            set;
        }

        [DispId(1009)]
        new int MaxRecords
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1009)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1009)]
            [param: In]
            set;
        }

        [DispId(1010)]
        new int RecordCount
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1010)]
            get;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [SpecialName]
        [DispId(1011)]
        new void let_Source([In][MarshalAs(UnmanagedType.BStr)] string pvSource);

        [DispId(1011)]
        new object Source
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1011)]
            [return: MarshalAs(UnmanagedType.Struct)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1011)]
            [param: In]
            [param: MarshalAs(UnmanagedType.IDispatch)]
            set;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1012)]
        new void AddNew([Optional][In][MarshalAs(UnmanagedType.Struct)] object FieldList, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Values);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1013)]
        new void CancelUpdate();

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1014)]
        new void Close();

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1015)]
        new void Delete([In] AffectEnum AffectRecords = AffectEnum.adAffectCurrent);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1016)]
        [return: MarshalAs(UnmanagedType.Struct)]
        new object GetRows([In] int Rows = -1, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Start, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Fields);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1017)]
        new void Move([In] int NumRecords, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Start);
        */
        //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1018)]
        void MoveNext();
        /*
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1019)]
        new void MovePrevious();
        */
        //[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1020)]
        void MoveFirst();
        /*
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1021)]
        new void MoveLast();

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1022)]
        new void Open([Optional][In][MarshalAs(UnmanagedType.Struct)] object Source, [Optional][In][MarshalAs(UnmanagedType.Struct)] object ActiveConnection, [In] CursorTypeEnum CursorType = CursorTypeEnum.adOpenUnspecified, [In] LockTypeEnum LockType = LockTypeEnum.adLockUnspecified, [In] int Options = -1);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1023)]
        new void Requery([In] int Options = -1);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [TypeLibFunc(64)]
        [DispId(1610809378)]
        new void _xResync([In] AffectEnum AffectRecords = AffectEnum.adAffectAll);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1025)]
        new void Update([Optional][In][MarshalAs(UnmanagedType.Struct)] object Fields, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Values);

        [DispId(1047)]
        new PositionEnum AbsolutePage
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1047)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1047)]
            [param: In]
            set;
        }

        [DispId(1026)]
        new EditModeEnum EditMode
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1026)]
            get;
        }

        [DispId(1030)]
        new object Filter
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1030)]
            [return: MarshalAs(UnmanagedType.Struct)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1030)]
            [param: In]
            [param: MarshalAs(UnmanagedType.Struct)]
            set;
        }

        [DispId(1050)]
        new int PageCount
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1050)]
            get;
        }

        [DispId(1048)]
        new int PageSize
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1048)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1048)]
            [param: In]
            set;
        }

        [DispId(1031)]
        new string Sort
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1031)]
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1031)]
            [param: In]
            [param: MarshalAs(UnmanagedType.BStr)]
            set;
        }

        [DispId(1029)]
        new int Status
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1029)]
            get;
        }

        [DispId(1054)]
        new int State
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1054)]
            get;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1610809392)]
        [TypeLibFunc(64)]
        [return: MarshalAs(UnmanagedType.Interface)]
        new Recordset _xClone();

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1035)]
        new void UpdateBatch([In] AffectEnum AffectRecords = AffectEnum.adAffectAll);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1049)]
        new void CancelBatch([In] AffectEnum AffectRecords = AffectEnum.adAffectAll);

        [DispId(1051)]
        new CursorLocationEnum CursorLocation
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1051)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1051)]
            [param: In]
            set;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1052)]
        [return: MarshalAs(UnmanagedType.Interface)]
        new Recordset NextRecordset([Optional][MarshalAs(UnmanagedType.Struct)] out object RecordsAffected);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1036)]
        new bool Supports([In] CursorOptionEnum CursorOptions);

        [DispId(-8)]
        new object Collect
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(-8)]
            [TypeLibFunc(64)]
            [return: MarshalAs(UnmanagedType.Struct)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(-8)]
            [TypeLibFunc(64)]
            [param: In]
            [param: MarshalAs(UnmanagedType.Struct)]
            set;
        }

        [DispId(1053)]
        new MarshalOptionsEnum MarshalOptions
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1053)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1053)]
            [param: In]
            set;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1058)]
        new void Find([In][MarshalAs(UnmanagedType.BStr)] string Criteria, [In] int SkipRecords = 0, [In] SearchDirectionEnum SearchDirection = SearchDirectionEnum.adSearchForward, [Optional][In][MarshalAs(UnmanagedType.Struct)] object Start);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1055)]
        new void Cancel();

        [DispId(1056)]
        new object DataSource
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1056)]
            [return: MarshalAs(UnmanagedType.IUnknown)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1056)]
            [param: In]
            [param: MarshalAs(UnmanagedType.IUnknown)]
            set;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [TypeLibFunc(64)]
        [DispId(1610874883)]
        new void _xSave([In][MarshalAs(UnmanagedType.BStr)] string FileName = "", [In] PersistFormatEnum PersistFormat = PersistFormatEnum.adPersistADTG);

        [DispId(1061)]
        new object ActiveCommand
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1061)]
            [return: MarshalAs(UnmanagedType.IDispatch)]
            get;
        }

        [DispId(1063)]
        new bool StayInSync
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1063)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1063)]
            [param: In]
            set;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1062)]
        [return: MarshalAs(UnmanagedType.BStr)]
        new string GetString([In] StringFormatEnum StringFormat = StringFormatEnum.adClipString, [In] int NumRows = -1, [In][MarshalAs(UnmanagedType.BStr)] string ColumnDelimeter = "", [In][MarshalAs(UnmanagedType.BStr)] string RowDelimeter = "", [In][MarshalAs(UnmanagedType.BStr)] string NullExpr = "");

        [DispId(1064)]
        new string DataMember
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1064)]
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1064)]
            [param: In]
            [param: MarshalAs(UnmanagedType.BStr)]
            set;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1065)]
        new CompareEnum CompareBookmarks([In][MarshalAs(UnmanagedType.Struct)] object Bookmark1, [In][MarshalAs(UnmanagedType.Struct)] object Bookmark2);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1034)]
        [return: MarshalAs(UnmanagedType.Interface)]
        new Recordset Clone([In] LockTypeEnum LockType = LockTypeEnum.adLockUnspecified);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1024)]
        new void Resync([In] AffectEnum AffectRecords = AffectEnum.adAffectAll, [In] ResyncEnum ResyncValues = ResyncEnum.adResyncAllValues);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1066)]
        new void Seek([In][MarshalAs(UnmanagedType.Struct)] object KeyValues, [In] SeekEnum SeekOption = SeekEnum.adSeekFirstEQ);

        [DispId(1067)]
        new string Index
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1067)]
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(1067)]
            [param: In]
            [param: MarshalAs(UnmanagedType.BStr)]
            set;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [DispId(1057)]
        void Save([Optional][In][MarshalAs(UnmanagedType.Struct)] object Destination, [In] PersistFormatEnum PersistFormat = PersistFormatEnum.adPersistADTG);*/
    }

    [SourceClassName("Recordset")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [DefaultMember("Fields")] // +[DispId(0)] +[IsDefault]
    internal sealed class MyADODBRecordSet : IReflectOnClrType, Recordset
    {
        private readonly MyADODBFields _fields;
        public MyADODBRecordSet()
        {
            _fields = new MyADODBFields();
        }

        [DispId(0)]
        public Fields Fields()
        {
            return _fields;
        }

        [DispId(0)]  // This makes it the default VBScript property
        public Field Fields(object fieldIndexOrName)
        {
            return _fields.GetItem(fieldIndexOrName);
        }

        [DispId(1014)]
        public void Close()
        {

        }

        [DispId(1020)]
        public void MoveFirst()
        {

        }

        [DispId(1006)]
        public bool EOF
        {
            [DispId(1006)]
            get
            {
                return true;
            }
        }
        [DispId(1018)]
        public void MoveNext()
        {

        }

    }
    [SourceClassName("Fields")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyADODBFields : IReflectOnClrType, Fields
    {
        public MyADODBFields()
        {
        }

        public int Count { get; }

        [DispId(0)]  // This makes it the default VBScript property
        public Field GetItem(object fieldIndexOrName)
        {
            return new MyADODBField((string)fieldIndexOrName, null);
        }
    }

    [SourceClassName("Field")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyADODBField : IReflectOnClrType, Field
    {
        public MyADODBField(string name, object fieldValue)
        {
            Name = name;
            Value = fieldValue;
        }
        [DispId(1100)]
        public string Name { get; }

        [DispId(0)]
        public object Value { get; }
    }

    [DefaultMember("Value")] // +[DispId(0)] +[IsDefault]
    internal interface Field : Field20
    {
    }
    internal interface Field20
    {
    }

    internal interface Fields : Fields20
    {
    }
    internal interface Fields20 : Fields15
    {
    }
    internal interface Fields15 //
    {
        [DispId(1)]
        int Count { get; }
    }

}
