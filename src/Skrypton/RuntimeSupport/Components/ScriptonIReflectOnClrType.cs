// Moved from the test project as-is; it was written without nullable reference type annotations
#nullable disable
// The members are bound by name from scripts (so they must stay instance members) and the names mirror the COM type library
#pragma warning disable CA1822, CA1008, CA1711, CA1716, CA1309, CA1859
using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// IReflect over a CLR type that exposes only the members marked with [DispId] (like a COM dispatch interface) - the base of the built-in
    /// COM object replacements (moved from the test project, where it is IReflectOnClrType)
    /// </summary>
    internal class ScriptonIReflectOnClrType : IReflect
    {
        private readonly Type _type;

        public ScriptonIReflectOnClrType(Type type)
        {
            _type = type;
        }
        protected ScriptonIReflectOnClrType()
        {
            _type = GetType();
        }

        internal static ScriptonIReflectOnClrType ForType(Type type)
        {
            return new ScriptonIReflectOnClrType(type);
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
        private static bool TryReplaceObjectsWithDefaultValues(object[] args, out object[] valueArgs)
        {
            valueArgs = (object[])args.Clone();
            bool replaced = false;
            for (int i = 0; i < valueArgs.Length; i++)
            {
                object arg = valueArgs[i];
                if ((arg == null) || (arg is string) || arg.GetType().IsPrimitive || (arg is DateTime) || (arg is decimal) || (arg is Array))
                    continue;
                PropertyInfo defaultProperty = arg.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .FirstOrDefault(p => p.CanRead && (p.GetIndexParameters().Length == 0)
                        && ((p.GetCustomAttribute<DispIdAttribute>()?.Value == 0) || (p.GetCustomAttribute<Skrypton.RuntimeSupport.Attributes.IsDefaultAttribute>() != null)));
                if (defaultProperty == null)
                    continue;
                valueArgs[i] = defaultProperty.GetValue(arg);
                replaced = true;
            }
            return replaced;
        }

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
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                // The script must see the member's own error (eg. a VBScript error number such as 53 File not found), not the reflection wrapper
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                throw;
            }
            catch (ArgumentException) when ((args != null) && TryReplaceObjectsWithDefaultValues(args, out object[] valueArgs))
            {
                // As with COM variant conversion: an object passed where a value is expected is converted through its default property (eg. a
                // Folder passed to fso.GetFolder is its Path, verified with cscript)
                return InvokeMember(name, invokeAttr, binder, target, valueArgs, modifiers, culture, namedParameters);
            }
            catch (System.MissingMemberException ex)
            {
                // obj.Prop(i) where Prop has no parameters is obj.Prop.Item(i): the default member of the value of Prop (eg. doc.Tables(1), as with COM)
                if ((args != null) && (args.Length > 0) && ((invokeAttr & (BindingFlags.GetProperty | BindingFlags.InvokeMethod)) != 0))
                {
                    PropertyInfo parameterLess = GetProperty(name, BindingFlagsVBScript);
                    if ((parameterLess != null) && (parameterLess.GetIndexParameters().Length == 0) && parameterLess.CanRead)
                    {
                        object value = parameterLess.GetValue(target ?? this, null);
                        if (value is IReflect reflect)
                            return reflect.InvokeMember("", BindingFlags.GetProperty | BindingFlags.InvokeMethod, binder, value, args, modifiers, culture, namedParameters);
                    }
                }

                // obj.Member(i, j) = v where Member(i, j) returns an object: v is assigned to the default member of that object (eg. ws.Cells(1, 2) = v
                // sets the Value of the cell, as with COM)
                if ((args != null) && (args.Length > 1) && ((invokeAttr & BindingFlags.SetProperty) != 0))
                {
                    object[] memberArgs = args.Take(args.Length - 1).ToArray();
                    object value = InvokeMember(name, BindingFlags.GetProperty | BindingFlags.InvokeMethod, binder, target, memberArgs, null, culture, null);
                    if (value is IReflect reflect)
                        return reflect.InvokeMember("", BindingFlags.SetProperty, binder, value, new[] { args[args.Length - 1] }, null, culture, null);
                }

                // No overload takes the arguments as they are: convert object arguments through their default property and try again (see above)
                if ((args != null) && TryReplaceObjectsWithDefaultValues(args, out object[] valueArgs))
                    return InvokeMember(name, invokeAttr, binder, target, valueArgs, modifiers, culture, namedParameters);
                if (ex.Message.Length > 0)
                {
                    MemberInfo candidate = DebugInspectMember(name, invokeAttr, args);
                    if (candidate != null && candidate is MethodInfo mtd)
                    {
                        //if (invokeAttr.HasFlag(BindingFlags.InvokeMethod)) // rescue named properties setter
                        {
                            return mtd.Invoke(target ?? this, args);
                        }
                    }
                    else if (candidate != null && candidate is PropertyInfo prp)
                    {
                        if (invokeAttr.HasFlag(BindingFlags.SetProperty))
                        {
                            /*
                               Dim serv: Set serv = model.QueryService("{24280463-E982-4c84-A777-89FD30227B9A}")
                               If serv.enabled ( nSUID ) = True Then
                                 serv.Enabled ( nSUID ) = False
                               Else
                                 serv.Enabled ( nSUID) = True
                               End If
                             */
                            prp.SetValue(target ?? this, args); // named properties like
                            return null;
                        }
                        else
                        {
                            if (invokeAttr.HasFlag(BindingFlags.GetProperty))
                            {
                                return prp.GetValue(target ?? this, args);
                            }

                        }
                    }
                }
                throw;
            }
        }
        private MemberInfo DebugInspectMember(string name, BindingFlags invokeAttr, object[] args)
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
                return null;
            }
            else
            {
                // FOUND: put a breakpoint here
                return candidate;
            }
        }
    }
}
