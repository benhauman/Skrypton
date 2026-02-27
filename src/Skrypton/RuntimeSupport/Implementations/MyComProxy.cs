using System;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using static Skrypton.RuntimeSupport.UnsafeNativeMethods;

namespace Skrypton.RuntimeSupport.Implementations
{
    [DebuggerDisplay("{_progId}")]
    internal sealed class MyComProxy : IDisposable, IReflect
    {
        private readonly string _progId;
        private readonly Type _comType;

        internal readonly object _comInstance;
        //private readonly dynamic _dyn; // dynamic boundary

        private MyComProxy(string progId, Type comType, object comInstanceInstance)
        {
            _progId = progId;
            _comType = comType;
            _comInstance = comInstanceInstance;
        }
        internal static MyComProxy CreateComProxy(string progId, Type comType)
        {
            object comInstance = Activator.CreateInstance(comType);
            if (IDispatchAccess.ImplementsIDispatch(comInstance, out IDispatchAccess.IDispatch dispatch))
            {
                // ok! => can be used 'IDispatchAccess'
            }
            else
            {
                // If a COM object does not implement IDispatch, it cannot be used from VBScript.
                if (comInstance is IReflect rfl)
                {
                    //return rfl;
                }
                else
                {
                    throw new InvalidOperationException($"Class '{progId}' does not implement IDispatch.");
                }
            }

            return new MyComProxy(progId, comType, comInstance);
        }
        public void Dispose()
        {
            if (_comInstance != null)
            {
                /*
.NET creates a Runtime Callable Wrapper (RCW). One RCW can call AddRef multiple times internally.
Examples that add hidden references:

* Passing the COM object as a method argument
* Calling InvokeMember
* Using dynamic
* Casting to object
* Crossing AppDomains
* COM property access

What ReleaseComObject actually does:
int count = Marshal.ReleaseComObject(_comInstance);

Calls IUnknown::Release
Decrements the COM ref count by 1
Returns the remaining ref count
                 */
                //while (Marshal.ReleaseComObject(_comInstance) > 0) { }
            }
        }
        #region Explicit IReflect Implementation
        private IDispatchAccess.IDispatch _dispatch => (IDispatchAccess.IDispatch)_comInstance;
        FieldInfo IReflect.GetField(string name, BindingFlags bindingAttr) => null;
        FieldInfo[] IReflect.GetFields(BindingFlags bindingAttr) => Array.Empty<FieldInfo>();

        MethodInfo IReflect.GetMethod(string name, BindingFlags bindingAttr) => new DispatchMethodInfo(_dispatch, name);
        MethodInfo IReflect.GetMethod(string name, BindingFlags bindingAttr, Binder binder, Type[] types, ParameterModifier[] modifiers)
            => new DispatchMethodInfo(_dispatch, name);
        MethodInfo[] IReflect.GetMethods(BindingFlags bindingAttr) => Array.Empty<MethodInfo>();

        PropertyInfo IReflect.GetProperty(string name, BindingFlags bindingAttr) => new DispatchPropertyInfo(_dispatch, name);
        PropertyInfo IReflect.GetProperty(string name, BindingFlags bindingAttr, Binder binder, Type returnType, Type[] types, ParameterModifier[] modifiers)
            => new DispatchPropertyInfo(_dispatch, name);
        PropertyInfo[] IReflect.GetProperties(BindingFlags bindingAttr) => Array.Empty<PropertyInfo>();

        MemberInfo[] IReflect.GetMember(string name, BindingFlags bindingAttr)
            => new MemberInfo[] { ((IReflect)this).GetMethod(name, bindingAttr), ((IReflect)this).GetProperty(name, bindingAttr) };
        MemberInfo[] IReflect.GetMembers(BindingFlags bindingAttr) => Array.Empty<MemberInfo>();

        object IReflect.InvokeMember(
            string name,
            BindingFlags invokeAttr,
            Binder binder,
            object target,
            object[] args,
            ParameterModifier[] modifiers,
            CultureInfo culture,
            string[] namedParameters)
        {
            if (target != this)
                throw new TargetException("Target object is not the wrapped IDispatch.");

            int dispid = IDispatchAccess.GetDispId(_dispatch, name);
            if (dispid == -1)
                throw new MissingMethodException($"Member '{name}' not found on IDispatch object.");

            if ((invokeAttr & BindingFlags.SetProperty) != 0)
            {
                //IDispatchAccess.prop
            }

            return IDispatchAccess.CallMethodU(_dispatch, name, args);

            //int dispid = GetDispId(name);
            //if (dispid == -1)
            //    throw new MissingMethodException($"Member '{name}' not found on IDispatch object.");
            //
            //// Determine invoke flags
            //short wFlags = 0;
            //if ((invokeAttr & BindingFlags.GetProperty) != 0) wFlags |= (short)INVOKEKIND.INVOKE_PROPERTYGET;
            //if ((invokeAttr & BindingFlags.SetProperty) != 0) wFlags |= (short)INVOKEKIND.INVOKE_PROPERTYPUT;
            //if ((invokeAttr & BindingFlags.InvokeMethod) != 0) wFlags |= (short)INVOKEKIND.INVOKE_FUNC;
            //
            //var dispParams = new DISPPARAMS
            //{
            //    cArgs = args?.Length ?? 0,
            //    cNamedArgs = namedParameters?.Length ?? 0,
            //    rgvarg = Marshal.AllocCoTaskMem(IntPtr.Size * (args?.Length ?? 0)),
            //    rgdispidNamedArgs = Marshal.AllocCoTaskMem(IntPtr.Size * (namedParameters?.Length ?? 0))
            //};
            //
            //try
            //{
            //    if (args != null)
            //    {
            //        for (int i = 0; i < args.Length; i++)
            //        {
            //            IntPtr argPtr = Marshal.GetIUnknownForObject(args[args.Length - 1 - i]);
            //            Marshal.WriteIntPtr(dispParams.rgvarg, i * IntPtr.Size, argPtr);
            //        }
            //    }
            //
            //    EXCEPINFO excepInfo = new EXCEPINFO();
            //    uint argErr;
            //    object result = null;
            //    _dispatch.Invoke(dispid, ref Guid.Empty, 0, wFlags, ref dispParams, out result, ref excepInfo, out argErr);
            //    return result;
            //}
            //finally
            //{
            //    Marshal.FreeCoTaskMem(dispParams.rgvarg);
            //    Marshal.FreeCoTaskMem(dispParams.rgdispidNamedArgs);
            //}
        }

        Type IReflect.UnderlyingSystemType => typeof(IDispatchAccess.IDispatch);

        #endregion
    }
    internal class DispatchMethodInfo : MethodInfo
    {
        private readonly IDispatchAccess.IDispatch _dispatch;
        private readonly string _name;

        public DispatchMethodInfo(IDispatchAccess.IDispatch dispatch, string name)
        {
            _dispatch = dispatch;
            _name = name;
        }

        public override object Invoke(object obj, BindingFlags invokeAttr, Binder binder, object[] parameters, CultureInfo culture)
        {
            throw new NotImplementedException();
            //return ((IReflect)new DispatchReflector(_dispatch))
            //    .InvokeMember(_name, invokeAttr | BindingFlags.InvokeMethod, binder, _dispatch, parameters, null, culture, null);
        }

        public override string Name => _name;
        public override Type DeclaringType => typeof(IDispatchAccess.IDispatch);
        public override Type ReflectedType => typeof(IDispatchAccess.IDispatch);
        public override MethodAttributes Attributes => MethodAttributes.Public;
        public override RuntimeMethodHandle MethodHandle { get; }
        public override ParameterInfo[] GetParameters() => Array.Empty<ParameterInfo>();
        public override MethodImplAttributes GetMethodImplementationFlags() => MethodImplAttributes.IL;
        public override object[] GetCustomAttributes(bool inherit) => Array.Empty<object>();
        public override object[] GetCustomAttributes(Type attributeType, bool inherit) => Array.Empty<object>();
        public override bool IsDefined(Type attributeType, bool inherit) => false;
        public override MethodInfo GetBaseDefinition()
        {
            throw new NotImplementedException();
        }

        public override Type ReturnType => typeof(object);
        public override ICustomAttributeProvider ReturnTypeCustomAttributes { get; }
    }
    internal class DispatchPropertyInfo : PropertyInfo
    {
        private readonly IDispatchAccess.IDispatch _dispatch;
        private readonly string _name;

        public DispatchPropertyInfo(IDispatchAccess.IDispatch dispatch, string name)
        {
            _dispatch = dispatch;
            _name = name;
        }

        public override object GetValue(object obj, BindingFlags invokeAttr, Binder binder, object[] index, CultureInfo culture)
        {
            //return ((IReflect)new DispatchReflector(_dispatch))
            //    .InvokeMember(_name, BindingFlags.GetProperty, binder, _dispatch, index, null, culture, null);
            throw new NotImplementedException();
        }

        public override void SetValue(object obj, object value, BindingFlags invokeAttr, Binder binder, object[] index, CultureInfo culture)
        {
            //((IReflect)new DispatchReflector(_dispatch))
            //    .InvokeMember(_name, BindingFlags.SetProperty, binder, _dispatch, index, null, culture, null);
            throw new NotImplementedException();
        }

        public override string Name => _name;
        public override Type DeclaringType => typeof(IDispatchAccess.IDispatch);
        public override Type ReflectedType => typeof(IDispatchAccess.IDispatch);
        public override PropertyAttributes Attributes => PropertyAttributes.None;
        public override bool CanRead => true;
        public override bool CanWrite => true;
        public override Type PropertyType { get; }

        public override MethodInfo[] GetAccessors(bool nonPublic)
        {
            throw new NotImplementedException();
        }

        public override MethodInfo GetGetMethod(bool nonPublic) => new DispatchMethodInfo(_dispatch, _name);
        public override MethodInfo GetSetMethod(bool nonPublic) => new DispatchMethodInfo(_dispatch, _name);
        public override ParameterInfo[] GetIndexParameters() => Array.Empty<ParameterInfo>();
        public override object[] GetCustomAttributes(bool inherit) => Array.Empty<object>();
        public override object[] GetCustomAttributes(Type attributeType, bool inherit) => Array.Empty<object>();
        public override bool IsDefined(Type attributeType, bool inherit) => false;
    }
}