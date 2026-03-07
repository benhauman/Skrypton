using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Xml;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    internal abstract class IDispatchBase : IDispatchAccess.IDispatch
    {
        protected IDispatchBase()
        {

        }

        //protected abstract IReflect ReflectTarget();
        private IReflectOnClrType ReflectTarget()
        {
            return IReflectOnClrType.ForType(GetType());
        }
        private object TargetInstanceForInvoke()
        {
            return this;
        }

        int IDispatchAccess.IDispatch.GetTypeInfoCount(out int Count)
        {
            Count = 0;
            return 0; // S_OK
        }

        int IDispatchAccess.IDispatch.GetTypeInfo(int iTInfo, int lcid, out ITypeInfo typeInfo)
        {
            typeInfo = null;
            return 0;// S_OK
        }

        int IDispatchAccess.IDispatch.GetIDsOfNames(ref Guid riid, string[] rgsNames, int cNames, int lcid, int[] rgDispId)
        {

            // typical mapping: map names to DispatchBaseReflect
            var reflect = ReflectTarget();
            if (reflect == null)
                throw new InvalidCastException();
            for (var idx = 0; idx < rgsNames.Length; idx++)
            {
                var rgsName = rgsNames[idx];

                var mbs = reflect.GetMember(rgsName, IReflectOnClrType.BindingFlagsVBScript);
                if (mbs == null)
                    throw new MissingMethodException(rgsName);

                DispIdAttribute attrDispid = null;
                foreach (var mb in mbs)
                {
                    var attr = mb.GetCustomAttribute<DispIdAttribute>();
                    if (attr != null)
                    {
                        attrDispid = attr;// mb.GetCustomAttribute<DispIdAttribute>() ?? throw new MissingMethodException($"Method {rgsName} is missing DispId attribute");
                        break;
                    }
                }
                if (attrDispid == null)
                    throw new MissingMethodException("DISP id not found: " + rgsName);
                rgDispId[idx] = attrDispid.Value;
            }

            return 0; // S_OK
        }

        int IDispatchAccess.IDispatch.Invoke(int dispIdMember, ref Guid riid, uint lcid, ushort wFlags, ref DISPPARAMS pDispParams, out object pVarResult, ref EXCEPINFO pExcepInfo, out uint pArgErr)
        {
            var reflect = ReflectTarget();
            if (reflect == null)
                throw new InvalidCastException();

            IDispatchAccess.InvokeFlags invFlags = (IDispatchAccess.InvokeFlags)wFlags;

            if (invFlags.HasFlag(IDispatchAccess.InvokeFlags.DISPATCH_PROPERTYGET) || invFlags.HasFlag(IDispatchAccess.InvokeFlags.DISPATCH_PROPERTYPUT) || invFlags.HasFlag(IDispatchAccess.InvokeFlags.DISPATCH_PROPERTYPUTREF))
            {
                var pis = reflect.GetProperties(IReflectOnClrType.BindingFlagsVBScript);

                foreach (var pi in pis)
                {
                    var attrDispid = pi.GetCustomAttribute<DispIdAttribute>() ?? throw new MissingMethodException($"Method {pi.Name} is missing DispId attribute");
                    if (attrDispid.Value == dispIdMember)
                    {
                        object[] parametersValues = ExtractArgs(ref pDispParams);
                        if (invFlags == IDispatchAccess.InvokeFlags.DISPATCH_PROPERTYPUT || invFlags == IDispatchAccess.InvokeFlags.DISPATCH_PROPERTYPUTREF)
                        {
                            if (parametersValues.Length != 1)
                                throw new ArgumentException($"Property set should have exactly one argument, but {parametersValues.Length} were passed"); // indexers not supported for now
                            pi.SetValue(TargetInstanceForInvoke(), parametersValues[0]);
                            pVarResult = null;
                            pArgErr = 0; // S_OK
                            return 0;// S_OK
                        }
                        else
                        {
                            if (parametersValues.Length != 0)
                                throw new ArgumentException($"Property get should not have any arguments, but {parametersValues.Length} were passed"); // indexers not supported for now
                            pVarResult = pi.GetValue(TargetInstanceForInvoke());
                            pArgErr = 0; // S_OK
                            return 0;// S_OK
                        }
                    }
                }

                // property not found, but maybe it's a method with the same DISPID?
            }

            // find method by metadata token
            if (invFlags.HasFlag(IDispatchAccess.InvokeFlags.DISPATCH_METHOD))
            {
                var mis = reflect.GetMethods(IReflectOnClrType.BindingFlagsVBScript);
                foreach (var mi in mis)
                {
                    var attrDispid = mi.GetCustomAttribute<DispIdAttribute>() ?? throw new MissingMethodException($"Method {mi.Name} is missing DispId attribute");
                    if (attrDispid.Value == dispIdMember)
                    {
                        object[] parametersValues = ExtractArgs(ref pDispParams);
                        pVarResult = mi.Invoke(TargetInstanceForInvoke(), parametersValues);
                        pArgErr = 0; // S_OK
                        return 0;// S_OK
                    }
                }

                throw new MissingMethodException($"DISPID[{dispIdMember}] not found for method");

            }

            throw new MissingMethodException($"DISPID[{dispIdMember}] not found for property");
        }
        private static object[] ExtractArgs(ref DISPPARAMS disp) // DISPPARAMS.rgvarg is a pointer to a COM VARIANT array, stored in reverse order, and must be manually unmarshalled before calling reflection.
        {
            int count = disp.cArgs;
            var result = new object[count];

            IntPtr current = disp.rgvarg;

            for (int i = 0; i < count; i++)
            {
                // COM stores parameters RIGHT‑TO‑LEFT → reverse at the end
                IntPtr variantPtr = IntPtr.Add(current, i * VARIANT.MarshalSizeOf);
                VARIANT v = Marshal.PtrToStructure<VARIANT>(variantPtr);

                result[i] = ConvertVariant(count - i, ref v);
            }

            Array.Reverse(result); // now LEFT‑TO‑RIGHT for reflection

            return result;
        }
        private static object ConvertVariant(int argumentPositionOneBased, ref VARIANT v)
        {
            VarEnum vt = (VarEnum)v.vt;

            // BYREF should be dereferenced first
            if ((vt & VarEnum.VT_BYREF) != 0)
            {
                //    return ConvertByRefVariant(ref v, vt & ~VarEnum.VT_BYREF);
                throw new NotSupportedException($"vt:{v.vt} ({vt}) pos: {argumentPositionOneBased}");
            }


            switch ((VarEnum)v.vt)
            {
                // ----------------------------
                //  INTEGER TYPES
                // ----------------------------

                case VarEnum.VT_I1: return (sbyte)v.bVal;
                case VarEnum.VT_UI1: return (byte)v.bVal;

                case VarEnum.VT_I2: return v.iVal;
                case VarEnum.VT_UI2: return (ushort)v.iVal;

                case VarEnum.VT_I4: return v.lVal;
                case VarEnum.VT_UI4: return (uint)v.lVal;

                case VarEnum.VT_I8: return v.llVal;
                case VarEnum.VT_UI8: return (ulong)v.llVal;

                // ----------------------------
                //  FLOATING POINT
                // ----------------------------

                case VarEnum.VT_R4: return v.fltVal;
                case VarEnum.VT_R8: return v.dblVal;

                // ----------------------------
                //  BOOLEAN
                // ----------------------------

                case VarEnum.VT_BOOL:   // VARIANT_BOOL: -1 = TRUE, 0 = FALSE
                    return v.boolVal;




                //case VarEnum.VT_I4:
                //    return (int)v.llVal;

                //case VarEnum.VT_I8:
                //    return v.llVal;

                //case VarEnum.VT_R8:
                //    return v.dblVal;

                case VarEnum.VT_BSTR:
                    if (v.ptr == IntPtr.Zero) return null;
                    return Marshal.PtrToStringBSTR(v.ptr);

                //                case VarEnum.VT_DISPATCH:
                //
                //#pragma warning disable CA1416
                //                    var obj = Marshal.GetObjectForIUnknown(v.ptr);
                //#pragma warning restore CA1416
                //                    return obj;

                case VarEnum.VT_EMPTY:
                case VarEnum.VT_NULL:
                    return null;

                // ----------------------------
                //  DATE (OLE Automation Date)
                // ----------------------------

                case VarEnum.VT_DATE:
                    return DateTime.FromOADate(v.dblVal);

                // ----------------------------
                //  SAFEARRAY
                // ----------------------------

                //case VarEnum.VT_ARRAY | VarEnum.VT_VARIANT:
                //    return ConvertSafeArray(v.ptr);
                //
                //case VarEnum.VT_ARRAY | VarEnum.VT_BSTR:
                //    return ConvertSafeArray(v.ptr);

                default:
                    throw new NotSupportedException($"VT type not supported: {v.vt} ({(VarEnum)v.vt}) pos: {argumentPositionOneBased}");
            }
        }
        /*
        private static object ConvertByRefVariant(ref VARIANT v, VarEnum baseType)
        {
            IntPtr p = v.ptr;
            if (p == IntPtr.Zero) return null;

            switch (baseType)
            {
                case VarEnum.VT_I2: return Marshal.ReadInt16(p);
                case VarEnum.VT_I4: return Marshal.ReadInt32(p);
                case VarEnum.VT_I8: return Marshal.ReadInt64(p);

                case VarEnum.VT_UI2: return (ushort)Marshal.ReadInt16(p);
                case VarEnum.VT_UI4: return (uint)Marshal.ReadInt32(p);
                case VarEnum.VT_UI8: return (ulong)Marshal.ReadInt64(p);

                case VarEnum.VT_R4:
                    unsafe { return *(float*)p.ToPointer(); }

                case VarEnum.VT_R8:
                    unsafe { return *(double*)p.ToPointer(); }

                case VarEnum.VT_BOOL:
                    return Marshal.ReadInt16(p) != 0;

                case VarEnum.VT_BSTR:
#pragma warning disable CA1416
                    return Marshal.PtrToStringBSTR(Marshal.ReadIntPtr(p));
#pragma warning restore CA1416

                case VarEnum.VT_DISPATCH:
#pragma warning disable CA1416
                    return Marshal.GetObjectForIUnknown(Marshal.ReadIntPtr(p));
#pragma warning restore CA1416

                default:
                    throw new NotSupportedException($"BYREF for {baseType} not supported");
            }
        }

        private static object[] ConvertSafeArray(IntPtr psa)
        {
            if (psa == IntPtr.Zero)
                return null;

#pragma warning disable CA1416
            Array arr = (Array)Marshal.GetObjectForIUnknown(psa);
#pragma warning restore CA1416

            object[] result = new object[arr.Length];
            Array.Copy(arr, result, arr.Length);
            return result;
        }
*/


    }

    public class IReflectOnClrType : IReflect
    {
        private readonly Type _type;

        public IReflectOnClrType(Type type)
        {
            _type = type;
        }
        protected IReflectOnClrType()
        {
            _type = GetType();
        }

        internal static IReflectOnClrType ForType(Type type)
        {
            return new IReflectOnClrType(type);
        }

        // --- IReflect default forwarding to the real Type ---
        public FieldInfo GetField(string name, BindingFlags bindingAttr)
        {
            var fi = _type.GetField(name, bindingAttr);
            return fi?.GetCustomAttribute<DispIdAttribute>() != null ? fi : null;
        }

        public FieldInfo[] GetFields(BindingFlags bindingAttr)
            => _type.GetFields(bindingAttr).Where(mb => mb.GetCustomAttribute<DispIdAttribute>() != null).ToArray();

        public MemberInfo[] GetMember(string name, BindingFlags bindingAttr)
        {
            return _type.GetMember(name, bindingAttr).Where(mb => mb.GetCustomAttribute<DispIdAttribute>() != null).ToArray();
        }

        public MemberInfo[] GetMembers(BindingFlags bindingAttr)
            => _type.GetMembers(bindingAttr).Where(mb => mb.GetCustomAttribute<DispIdAttribute>() != null).ToArray();

        public MethodInfo GetMethod(string name, BindingFlags bindingAttr)
        {
            var mis = _type.GetMethods(bindingAttr);
            foreach (var mi in mis)
            {
                if (string.Equals(mi.Name, name, StringComparison.Ordinal))
                {
                    if (mi.GetCustomAttribute<DispIdAttribute>() != null)
                    {
                        return mi;
                    }
                }
            }
            return null;
        }


        public MethodInfo GetMethod(
            string name,
            BindingFlags bindingAttr,
            Binder binder,
            Type[] types,
            ParameterModifier[] modifiers)
        {
            var mi = _type.GetMethod(name, bindingAttr, binder, types, modifiers);
            if (mi != null && mi.GetCustomAttribute<DispIdAttribute>() != null)
                return mi;
            return null;
        }

        public MethodInfo[] GetMethods(BindingFlags bindingAttr)
        {
            var mis = _type.GetMethods(bindingAttr);
            var misDISPID = mis.Where(mi => mi.GetCustomAttribute<DispIdAttribute>() != null).ToArray();
            return misDISPID;
        }

        public PropertyInfo[] GetProperties(BindingFlags bindingAttr)
        {
            return _type.GetProperties(bindingAttr).Where(mb => mb.GetCustomAttribute<DispIdAttribute>() != null).ToArray();
        }

        public PropertyInfo GetProperty(string name, BindingFlags bindingAttr)
        {
            var pi = _type.GetProperty(name, bindingAttr);
            return pi?.GetCustomAttribute<DispIdAttribute>() != null ? pi : null;
        }

        public PropertyInfo GetProperty(
            string name,
            BindingFlags bindingAttr,
            Binder binder,
            Type returnType,
            Type[] types,
            ParameterModifier[] modifiers)
        {
            var pi = _type.GetProperty(name, bindingAttr, binder, returnType, types, modifiers);
            return pi?.GetCustomAttribute<DispIdAttribute>() != null ? pi : null;
        }

        public Type UnderlyingSystemType => _type;

        // Default behavior: instance + public, case-insensitive (VB/VBA style),
        // and allow property get/put/method.
        internal const BindingFlags BindingFlagsVBScript = BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase;

        /// <summary>
        /// The key hook: COM late-binding calls come here via the CCW.
        /// You can customize name resolution, case sensitivity, default members, etc.
        /// </summary>
        public object InvokeMember(
            string name,
            BindingFlags invokeAttr,
            Binder binder,
            object target,
            object[] args,
            ParameterModifier[] modifiers,
            CultureInfo culture,
            string[] namedParameters)
        {
            try
            {
                return _type.InvokeMember(
                    name,
                    invokeAttr | BindingFlagsVBScript,
                    binder,
                    target ?? this,
                    args, //args_allmembers, use [DefaultParameterValue(...)] for optional parameters!
                    modifiers,
                    culture,
                    namedParameters);
            }
            catch (System.MissingMemberException ex)
            {
                if (ex.Message.Length > 0)
                    DebugInspectMember(name, invokeAttr, args);
                throw;
            }
        }
        private void DebugInspectMember(string name, BindingFlags invokeAttr, object[] args)
        {

            MemberInfo[] members = _type.GetMembers(invokeAttr | BindingFlagsVBScript);
            MethodInfo candidate = null;
            foreach (var member in members)
            {
                if (string.Equals(member.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    if (member is MethodInfo mi)
                    {
                        var prms = mi.GetParameters();
                        if (candidate == null)
                        {
                            candidate = mi;
                        }
                        else
                        {
                            if (prms.Length > candidate.GetParameters().Length)
                            {
                                candidate = mi;
                            }
                        }
                    }
                    else if (member is PropertyInfo pi)
                    {
                        if (pi.CanRead)
                        {

                        }
                        if (pi.CanWrite)
                        {
                            if (invokeAttr.HasFlag(BindingFlags.SetProperty))
                            {

                            }
                        }
                    }
                }
            }

            object[] args_allmembers;
            if (candidate == null)
            {
                args_allmembers = args;
            }
            else
            {
                var prms = candidate.GetParameters();
                if (prms.Length == args.Length)
                {
                    args_allmembers = args;
                }
                else
                {
                    args_allmembers = new object[prms.Length];

                    if (prms.Length > args.Length)
                    {
                        int ix = 0;
                        for (; ix < args.Length; ix++)
                        {
                            var arg = args[ix];
                            args_allmembers[ix] = arg;
                        }
                        for (; ix < prms.Length; ix++)
                        {
                            var prm = prms[ix];
                            if (prm.HasDefaultValue)
                            {
                                args_allmembers[ix] = prm.DefaultValue;
                            }
                            else
                            {
                                Type parameterType = prm.ParameterType;
                                if (!parameterType.IsValueType || Nullable.GetUnderlyingType(parameterType) != null)
                                {
                                    args_allmembers[ix] = null;
                                }
                                else
                                {
                                    //object defvalue = Activator.CreateInstance(parameterType);
                                    object defvalue = parameterType switch
                                    {
                                        Type t when t == typeof(int) => default(int),
                                        Type t when t == typeof(short) => default(short),
                                        Type t when t == typeof(bool) => default(bool),
                                        Type t when t == typeof(long) => default(long),
                                        Type t when t == typeof(DateTime) => default(DateTime),
                                        Type t when t == typeof(Guid) => default(Guid),
                                        _ => Activator.CreateInstance(parameterType)
                                    };

                                    args_allmembers[ix] = defvalue; // or the default value for the type
                                }
                            }
                        }
                    }
                }
            }

            if (candidate == null)
            {
                // NOT FOUND: put a breakpoint here
            }
            else
            {
                // FOUND: put a breakpoint here
            }
        }
    }
}