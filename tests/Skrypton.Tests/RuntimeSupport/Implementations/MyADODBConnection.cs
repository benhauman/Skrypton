using System;
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
        public string ConnectionString { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }


        private DbConnection _currentOpenConnectionOrNull;
        [DispId(10)]
        public void Open(string ConnectionString = "", string UserID = "", string Password = "", int Options = -1)
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

            _currentOpenConnectionOrNull = svc.CreateAndOpenDatabaseConnectionString(connectionString: ConnectionString, userName: UserID, password: Password);
        }
    }

    [SourceClassName("Command")] // for TYPENAME(CreateObject("ADODB.Command"))
    [ComVisible(true)] // Required because .NET can auto‑implement IDispatch when (1):COM‑visible:true, (2): interface mode:AutoDispatch and (3): DISPID(0) & DISPIDs used
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class MyADODBCommand : IReflectOnClrType, Command
    {
        // ADODB.Command
        public MyADODBCommand() : base(typeof(MyADODBCommand))
        {

        }

        public Connection ActiveConnection { get; set; }

        [DispId(2)]
        public string CommandText { get; set; }

        [DispId(7)]
        public int CommandType { get; set; } // CommandTypeEnum

        [DispId(0)]
        public Parameters Parameters => throw new NotImplementedException();
    }


    //[ComImport]
    //[CoClass(typeof(ConnectionClass))]
    //[Guid("00001550-0000-0010-8000-00AA006D2EA4")]
    internal interface Connection : _Connection//, ConnectionEvents_Event
    {
    }

    //[ComImport]
    //[Guid("00001550-0000-0010-8000-00AA006D2EA4")]
    [DefaultMember("ConnectionString")]
    //[TypeLibType(4160)]
    internal interface _Connection : Connection15
    {
    }

    //[ComImport]
    [DefaultMember("ConnectionString")]
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
    [DefaultMember("Parameters")]
    internal interface Command25 : Command15
    {
    }

    //[ComImport]
    //[Guid("00001508-0000-0010-8000-00AA006D2EA4")]
    [DefaultMember("Parameters")]
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

        [DispId(0)]
        Parameters Parameters
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [DispId(0)]
            [return: MarshalAs(UnmanagedType.Interface)]
            get;
        }
    }


    //[ComImport]
    //[TypeLibType(4288)]
    //[Guid("0000150D-0000-0010-8000-00AA006D2EA4")]
    internal interface Parameters : _DynaCollection
    {
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

}
