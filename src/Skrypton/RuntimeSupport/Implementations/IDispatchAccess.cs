using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Skrypton.RuntimeSupport.Implementations
{
    public static class IDispatchAccess
    {
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("oleaut32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        private static extern Int32 VariantClear(IntPtr pvarg);

        private const int LOCALE_SYSTEM_DEFAULT = 2048;
        private const int DISPID_PROPERTYPUT = -3;

        //private const int SizeOfNativeVariant = 16;
        //private static readonly ComTypes.DISPPARAMS EmptyDISPPARAMS = new ComTypes.DISPPARAMS()
        //{
        //    cArgs = 0,
        //    cNamedArgs = 0,
        //    rgdispidNamedArgs = IntPtr.Zero,
        //    rgvarg = IntPtr.Zero
        //};

        [Flags]
#pragma warning disable CA1028 // Enum Storage should be Int32
#pragma warning disable CA1711 // Identifiers should not have incorrect suffix
        public enum InvokeFlags : ushort
#pragma warning restore CA1711 // Identifiers should not have incorrect suffix
#pragma warning restore CA1028 // Enum Storage should be Int32
        {
            None = 0,
#pragma warning disable CA1707 // Identifiers should not contain underscores
            DISPATCH_METHOD = 1,
            DISPATCH_PROPERTYGET = 2,
            DISPATCH_PROPERTYPUT = 4,
            DISPATCH_PROPERTYPUTREF = 8
#pragma warning restore CA1707 // Identifiers should not contain underscores
        }

        public static bool ImplementsIDispatch(object source, out IDispatch? dispatch)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            dispatch = source as IDispatch;
            return dispatch != null;
        }

        //internal static void CallPropertySet(object source, string name, params object[] args)
        //{
        //    _ = Invoke<object>(source, InvokeFlags.DISPATCH_PROPERTYPUT, name, GetDispId(source, name), args);
        //}
        public static T CallMethodT<T>(IDispatch source, string name, params object[] args)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name), "Null/blank name specified");
            if (args == null)
                throw new ArgumentNullException(nameof(args));

            return Invoke<T>(source, InvokeFlags.DISPATCH_METHOD, name, GetDispId(source, name), args);
        }

        public static object CallMethodU(IDispatch source, string name, params object[] args)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name), "Null/blank name specified");
            if (args == null)
                throw new ArgumentNullException(nameof(args));

            return CallMethodT<object>(source, name, args);
        }

        public static T GetProperty<T>(IDispatch source, string name)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name), "Null/blank name specified");

            return Invoke<T>(source, InvokeFlags.DISPATCH_PROPERTYGET, name, GetDispId(source, name));
        }

        public static object GetProperty(IDispatch source, string name)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name), "Null/blank name specified");

            return GetProperty<object>(source, name);
        }

        /// <summary>
        /// This will raise an exception for a null source, null or blank name or if the name could not be matched to a DispId
        /// </summary>
        public static int GetDispId(IDispatch source, string name)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name), "Null/blank name specified");

            var IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
            var rgDispId = new int[1] { 0 }; // This will be populated with a the DispId of the named member (if available)
            var hrRet = ((IDispatch)source).GetIDsOfNames
            (
                ref IID_NULL,
                new string[1] { name },
                1, // number of names to get ids for
                LOCALE_SYSTEM_DEFAULT,
                rgDispId
            );
            if (hrRet != 0)
            {
                var message = "Invalid member \"" + name + "\"";
                var errorType = GetErrorMessageForHResult(hrRet);
                if (errorType != CommonErrors.Unknown)
                    message += " [" + errorType.ToString() + "]";
                throw new IDispatchAccessException(message, source, memberNameIfSpecified: name, dispIdIfKnown: null, errorType: errorType);
            }
            return rgDispId[0];
        }

        public static T Invoke<T>(object source, InvokeFlags invokeFlags, string memberName, int dispId, params object[] args)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if ((invokeFlags != InvokeFlags.DISPATCH_METHOD)
            && (invokeFlags != InvokeFlags.DISPATCH_PROPERTYGET)
            && (invokeFlags != InvokeFlags.DISPATCH_PROPERTYPUT)
            && (invokeFlags != InvokeFlags.DISPATCH_PROPERTYPUTREF)
            && (invokeFlags != (InvokeFlags.DISPATCH_METHOD | InvokeFlags.DISPATCH_PROPERTYGET)))
            {
                // It's acceptable to specify both DISPATCH_METHOD and DISPATCH_PROPERTYGET together
                throw new ArgumentOutOfRangeException(nameof(invokeFlags));
            }
            if (args == null)
                throw new ArgumentNullException(nameof(args));

            var memoryAllocationsToFree = new List<IntPtr>();
            IntPtr rgdispidNamedArgs;
            int cNamedArgs;
            if ((invokeFlags == InvokeFlags.DISPATCH_PROPERTYPUT) || (invokeFlags == InvokeFlags.DISPATCH_PROPERTYPUTREF))

            {
                // There must be at least one argument specified; only one if it is a non-indexed property and multiple if there are index values as well as the
                // value to set to
                if (args.Length < 1)
                    throw new ArgumentException("At least one argument must be specified when DISPATCH_PROPERTYPUT(REF) is requested");

                var pdPutID = Marshal.AllocCoTaskMem(sizeof(Int64));
                Marshal.WriteInt64(pdPutID, DISPID_PROPERTYPUT);
                memoryAllocationsToFree.Add(pdPutID);

                rgdispidNamedArgs = pdPutID;
                cNamedArgs = 1;
            }
            else
            {
                rgdispidNamedArgs = IntPtr.Zero;
                cNamedArgs = 0;
            }

            int SizeOfNativeVariant = VARIANT.MarshalSizeOf; // 24 and not 16

            var variantsToClear = new List<IntPtr>();
            IntPtr rgvarg;
            if (args.Length == 0)
            {
                rgvarg = IntPtr.Zero;
            }
            else
            {
                // We need to allocate enough memory to store a variant for each argument (and then populate this memory)
                rgvarg = Marshal.AllocCoTaskMem(SizeOfNativeVariant * args.Length);
                memoryAllocationsToFree.Add(rgvarg);
                for (var index = 0; index < args.Length; index++)
                {
                    // Note: The "IDispatch::Invoke method (Automation)" page (http://msdn.microsoft.com/en-us/library/windows/desktop/ms221479(v=vs.85).aspx)
                    // states that "Arguments are stored in pDispParams->rgvarg in reverse order" so we'll reverse them here
                    var arg = args[(args.Length - 1) - index];

                    // According to http://stackoverflow.com/a/1866268 it seems like using ToInt64 here will be valid for both 32 and 64 bit machines. While
                    // this may apparently not be the most performant approach, it should do the job.
                    // Don't think we have to worry about pinning any references when we do this manipulation here since we are allocating the array in
                    // unmanaged memory and so the garbage collector won't be moving anything around (and GetNativeVariantForObject copies the reference
                    // and automatic pinning will prevent the GC from interfering while this is happening).
                    var pVariant = new IntPtr(
                        rgvarg.ToInt64() + (SizeOfNativeVariant * index)
                    );
                    Marshal.GetNativeVariantForObject(arg, pVariant);
                    variantsToClear.Add(pVariant);
                }
            }

            var dispParams = new ComTypes.DISPPARAMS()
            {
                cArgs = args.Length,
                rgvarg = rgvarg,
                cNamedArgs = cNamedArgs,
                rgdispidNamedArgs = rgdispidNamedArgs
            };

            try
            {
                var IID_NULL = new Guid("00000000-0000-0000-0000-000000000000");
                UInt32 pArgErr = 0;
                object varResult;
                var excepInfo = new ComTypes.EXCEPINFO();
                var hrRet = ((IDispatch)source).Invoke
                (
                    dispId,
                    ref IID_NULL,
                    LOCALE_SYSTEM_DEFAULT,
                    (ushort)invokeFlags,
                    ref dispParams,
                    out varResult,
                    ref excepInfo,
                    out pArgErr
                );
                if (hrRet != 0)
                {
                    var errorType = GetErrorMessageForHResult(hrRet);
                    if (errorType == CommonErrors.DISP_E_MEMBERNOTFOUND)
                        throw new IDispatchAccessException("Member not found", source, memberNameIfSpecified: null, dispIdIfKnown: dispId, errorType: errorType);

                    string memberNameX = memberName ?? Skrypton.RuntimeSupport.Information.TryGetIDispatchMemberName(source, dispId);

                    var message = $"Failing attempting to invoke member '{memberNameX}' with DispId '" + dispId + "': ";
#pragma warning disable CA1820 // Test for empty strings using string length
                    if ((excepInfo.bstrDescription ?? "").Trim() == "")
#pragma warning restore CA1820 // Test for empty strings using string length

                    {
                        message += "Unspecified error";
                    }
                    else
                        message += excepInfo.bstrDescription;
                    if (errorType != CommonErrors.Unknown)
                        message += " [" + errorType.ToString() + "]";
                    if (args.Length > 0)
                    {
#pragma warning disable CA1031 // Do not catch general exception types
                        try { Marshal.GetObjectsForNativeVariants(rgvarg, args.Length); }
                        catch (Exception exx)
                        {
                            message += " - " + exx.Message;
                            // In Visual Studio 2012, if the "Prefer 32-bit" build option is not enabled then arguments do not get written into memory
                            // correctly which will result in the request failing. If an exception has been raised then we can confirm if this is the
                            // problem by trying to pull the arguments back out of the "rgvarg", if this operation fails then the absence of this
                            // build option is most likely the cause.
                            message += $" - this may be due to the \"Prefer 32-bit\" option not being enabled. (Is64BitProcess:{Environment.Is64BitProcess})";// in Visual Studio";
                        }
#pragma warning restore CA1031 // Do not catch general exception types
                    }
                    throw new IDispatchAccessException(message, source, memberNameIfSpecified: null, dispIdIfKnown: dispId, errorType: errorType);
                }
                return (T)varResult;
            }
            finally
            {
                foreach (var variantToClear in variantsToClear)
                {
                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) // Ubuntu: [DllNotFoundException] Unable to load shared library 'OleAuth32.dll' or one of its dependencies.
                    {
                        _ = VariantClear(variantToClear); //  Clear the VARIANT contents, Without this, BSTRs/SAFEARRAYs leak. VariantClear does not free the memory block that holds the VARIANT struct.
                    }
                    Marshal.FreeHGlobal(variantToClear); // Free the memory block
                }

                foreach (var memoryAllocationToFree in memoryAllocationsToFree)
                    Marshal.FreeCoTaskMem(memoryAllocationToFree);
            }
        }

        public static void FreeDISPPARAMS(ComTypes.DISPPARAMS dispParams)
        {
            FreeArgs(dispParams.rgvarg, dispParams.cArgs);
        }
        // Free VARIANT array
        private static void FreeArgs(IntPtr ptr, int count)
        {
            int SizeOfNativeVariant = VARIANT.MarshalSizeOf; // 24 and not 16
            for (int i = 0; i < count; i++)
            {
                IntPtr variantToClear = IntPtr.Add(ptr, i * SizeOfNativeVariant);
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) // Ubuntu: [DllNotFoundException] Unable to load shared library 'OleAuth32.dll' or one of its dependencies.
                {
                    _ = VariantClear(variantToClear); //  Clear the VARIANT contents, Without this, BSTRs/SAFEARRAYs leak. VariantClear does not free the memory block that holds the VARIANT struct.
                }
                Marshal.FreeHGlobal(variantToClear); // Free the memory block
            }
            Marshal.FreeCoTaskMem(ptr);
        }

        private static CommonErrors GetErrorMessageForHResult(int hrRet)
        {
            if (Enum.IsDefined(typeof(CommonErrors), hrRet))
                return (CommonErrors)hrRet;

            return CommonErrors.Unknown;
        }

        internal static VarEnum GetVariantType(object value)
        {
            /*
                    if (value == null)
                   return VarEnum.VT_EMPTY;

               switch (value)
               {
                   case bool _:               return VarEnum.VT_BOOL;
                   case byte _:               return VarEnum.VT_UI1;
                   case sbyte _:              return VarEnum.VT_I1;
                   case short _:              return VarEnum.VT_I2;
                   case ushort _:             return VarEnum.VT_UI2;
                   case int _:                return VarEnum.VT_I4;
                   case uint _:               return VarEnum.VT_UI4;
                   case long _:               return VarEnum.VT_I8;
                   case ulong _:              return VarEnum.VT_UI8;
                   case float _:              return VarEnum.VT_R4;   // System.Single
                   case double _:             return VarEnum.VT_R8;
                   case decimal _:            return VarEnum.VT_DECIMAL;
                   case string _:             return VarEnum.VT_BSTR;
                   case DateTime _:           return VarEnum.VT_DATE;
                   case DBNull _:             return VarEnum.VT_NULL;

                   // COM objects
                   case System.__ComObject _: return VarEnum.VT_DISPATCH;

                   // Arrays
                   case Array _:              return VarEnum.VT_ARRAY | VarEnum.VT_VARIANT;

                   default:
                       return VarEnum.VT_VARIANT;
               }

             */
            IntPtr ptr = Marshal.AllocCoTaskMem(24); // VARIANT is 16 bytes on Windows

            try
            {
                // Write the object into a native VARIANT
                Marshal.GetNativeVariantForObject(value, ptr);

                // Read the vt field (first 2 bytes)
                short vt = Marshal.ReadInt16(ptr);

                return (VarEnum)vt;
            }
            finally
            {
                // Free the VARIANT
                Marshal.FreeCoTaskMem(ptr);
            }

        }

        // http://blogs.msdn.com/b/eldar/archive/2007/04/03/a-lot-of-hresult-codes.aspx
        public enum CommonErrors
        {
            Unknown = 0,
#pragma warning disable CA1707 // Identifiers should not contain underscores

            E_UNEXPECTED = -2147418113,
            E_NOTIMPL = -2147467263,
            E_OUTOFMEMORY = -2147024882,
            E_INVALIDARG = -2147024809,
            E_NOINTERFACE = -2147467262,
            E_POINTER = -2147467261,
            E_HANDLE = -2147024890,
            E_ABORT = -2147467260,
            E_FAIL = -2147467259,
            E_ACCESSDENIED = -2147024891,
            E_PENDING = -2147483638,

            DISP_E_UNKNOWNINTERFACE = -2147352575,
            DISP_E_MEMBERNOTFOUND = -2147352573,
            DISP_E_PARAMNOTFOUND = -2147352572,
            DISP_E_TYPEMISMATCH = -2147352571,
            DISP_E_UNKNOWNNAME = -2147352570,
            DISP_E_NONAMEDARGS = -2147352569,
            DISP_E_BADVARTYPE = -2147352568,
            DISP_E_EXCEPTION = -2147352567,
            DISP_E_OVERFLOW = -2147352566,
            DISP_E_BADINDEX = -2147352565,
            DISP_E_UNKNOWNLCID = -2147352564,
            DISP_E_ARRAYISLOCKED = -2147352563,
            DISP_E_BADPARAMCOUNT = -2147352562,
            DISP_E_PARAMNOTOPTIONAL = -2147352561,
            DISP_E_BADCALLEE = -2147352560,
            DISP_E_NOTACOLLECTION = -2147352559,
            DISP_E_DIVBYZERO = -2147352558,
            DISP_E_BUFFERTOOSMALL = -2147352557
#pragma warning restore CA1707 // Identifiers should not contain underscores
        }

        [ComImport()]
        [Guid("00020400-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
#pragma warning disable CA1034 // Nested types should not be visible
        public interface IDispatch
#pragma warning restore CA1034 // Nested types should not be visible
        {
            [Obsolete("Bad signature. Fix and verify signature before use.", true)]//[System.Security.SecurityCritical]
            [PreserveSig]
            int GetTypeInfoCount(out int Count);

            [PreserveSig]
            int GetTypeInfo
            (
                [In, MarshalAs(UnmanagedType.U4)] int iTInfo,
                [In, MarshalAs(UnmanagedType.U4)] int lcid,
                out ComTypes.ITypeInfo typeInfo
            );

            [PreserveSig]
            int GetIDsOfNames
            (
                ref Guid riid,
                [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)]
                string[] rgsNames,
                int cNames,
                int lcid,
                [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId
            );

            [PreserveSig]
            int Invoke
            (
                int dispIdMember,
                ref Guid riid,
                uint lcid,
                ushort wFlags,
                ref ComTypes.DISPPARAMS pDispParams,
                out object pVarResult,
                ref ComTypes.EXCEPINFO pExcepInfo,
                out UInt32 pArgErr
            );
        }

#pragma warning disable CA1034 // Nested types should not be visible
        public sealed class IDispatchAccessException : Exception
#pragma warning restore CA1034 // Nested types should not be visible
        {
            [Obsolete("do not use it")] private IDispatchAccessException() : this(null, innerException: null) { }
            [Obsolete("do not use it")] private IDispatchAccessException(string message) : this(message, innerException: null) { }
            [Obsolete("do not use it")]
            private IDispatchAccessException(string message, Exception innerException) : this(message, null, null, null, CommonErrors.Unknown, innerException: null) { }

            public IDispatchAccessException(string message, object target, string memberNameIfSpecified, int? dispIdIfKnown, CommonErrors errorType, Exception innerException = null)
                : base(message, innerException)
            {
                if (string.IsNullOrWhiteSpace(memberNameIfSpecified) && (dispIdIfKnown == null))
                    throw new ArgumentException("At least one of memberNameIfSpecified and dispIdIfKnown must be specified");
                if (!Enum.IsDefined(typeof(CommonErrors), errorType))
                    ErrorType = errorType;

                Target = target ?? throw new ArgumentNullException(nameof(target));
                MemberNameIfSpecified = memberNameIfSpecified;
                DispIdIfKnown = dispIdIfKnown;
                ErrorType = errorType;
            }

            private IDispatchAccessException(SerializationInfo info, StreamingContext context) : base(info, context) {
                Target = new object();
            }

            /// <summary>
            /// This will never be null
            /// </summary>
            public object Target { get; private set; }

            /// <summary>
            /// At least one of MemberNameIfSpecified and DispIdIfKnown will have a value
            /// </summary>
            public string? MemberNameIfSpecified { get; private set; }

            /// <summary>
            /// At least one of MemberNameIfSpecified and DispIdIfKnown will have a value
            /// </summary>
            public int? DispIdIfKnown { get; private set; }

            public CommonErrors ErrorType { get; private set; }
        }
    }

    /*
    [StructLayout(LayoutKind.Sequential)]
    public struct zzzzVARIANT
    {
        public ushort vt;
        public ushort reserved1;
        public ushort reserved2;
        public ushort reserved3;
        public IntPtr data1;
        public IntPtr data2;

        public static VARIANT FromObject(object value)
        {
            VARIANT v = new VARIANT();

            if (value == null)
            {
                v.vt = 1; // VT_NULL
                return v;
            }

            switch (Type.GetTypeCode(value.GetType()))
            {
                case TypeCode.String:
                    v.vt = 8; // VT_BSTR
                    v.data1 = Marshal.StringToBSTR((string)value);
                    break;

                case TypeCode.Boolean:
                    v.vt = 11; // VT_BOOL
                    v.data1 = (bool)value ? (IntPtr)(-1) : IntPtr.Zero;
                    break;

                case TypeCode.Int32:
                    v.vt = 3; // VT_I4
                    v.data1 = (IntPtr)(int)value;
                    break;

                default:
                    throw new NotSupportedException("Unsupported VARIANT type: " + value.GetType());
            }

            return v;
        }

        public long llVal =>
            MemoryMarshal.Read<long>(MemoryMarshal.AsBytes(MemoryMarshal.CreateSpan(ref data1, 1)));

        public double dblVal =>
            MemoryMarshal.Read<double>(MemoryMarshal.AsBytes(MemoryMarshal.CreateSpan(ref data1, 1)));

        public int lVal =>
            MemoryMarshal.Read<int>(MemoryMarshal.AsBytes(MemoryMarshal.CreateSpan(ref data1, 1)));
    }
    */
    [StructLayout(LayoutKind.Explicit)]
    [DebuggerDisplay("{VariantTypeEnum}:")]
#pragma warning disable CA1815 // Override equals and operator equals on value types
    public struct VARIANT
#pragma warning restore CA1815 // Override equals and operator equals on value types
    {
        // The padding is inserted because IntPtr is 8 bytes on x64, and the struct must align the pointer fields to 8‑byte boundaries
        // => (2+2+2+2)+4+(8+8)=24 bytes : (vt + reserved1 + 2 + 3) + alignment padding + (data1 + data2)
        // => in c++/native : the native definition uses a union, not two pointers => 16 bytes
        public static readonly int MarshalSizeOf = 24; // Marshal.SizeOf<VARIANT>(); // 24 and not 16 because of the 8-byte alignment of the largest field (long/double) in the struct, which causes padding in the native layout.

        internal VarEnum VariantTypeEnum => (VarEnum)vt;

        // The VARTYPE (2 bytes)
        [FieldOffset(0)]
        public ushort vt;

        [FieldOffset(2)]
        public ushort reserved1;

        [FieldOffset(4)]
        public ushort reserved2;

        [FieldOffset(6)]
        public ushort reserved3;

        // The value union begins at offset 8
        [FieldOffset(8)]
        public long llVal;                 // VT_I8, VT_UI8, VT_INT, VT_UINT

        [FieldOffset(8)]
        public int lVal;                   // VT_I4, VT_INT

        [FieldOffset(8)]
        public byte bVal;                  // VT_UI1

        [FieldOffset(8)]
        public short iVal;                 // VT_I2

        [FieldOffset(8)]
        public float fltVal;               // VT_R4

        [FieldOffset(8)]
        public double dblVal;              // VT_R8

        [FieldOffset(8)]
        public short boolVal;              // VARIANT_BOOL (-1=true, 0=false)

#pragma warning disable CA1720 // Identifier contains type name
        [FieldOffset(8)]
        public IntPtr ptr;                 // BSTR, IDispatch*, IUnknown*, arrays, etc.
#pragma warning restore CA1720 // Identifier contains type name
    }
}
