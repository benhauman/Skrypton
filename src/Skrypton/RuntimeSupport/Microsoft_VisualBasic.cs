using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace Skrypton.RuntimeSupport
{
    static class Information // @lubo: see class 'Information' in 'Microsoft.VisualBasic'
    {
        // Microsoft.VisualBasic.Information
        /// <summary>Returns a String value containing data-type information about a variable.</summary>
        /// <returns>Returns a String value containing data-type information about a variable.</returns>
        /// <param name="VarName">Required. Object variable. If Option Strict is Off, you can pass a variable of any data type except a structure.</param>
        /// <filterpriority>1</filterpriority>
        /// <PermissionSet>
        ///   <IPermission class="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" Flags="UnmanagedCode" />
        /// </PermissionSet>
        public static string TypeName(object VarName)
        {
            checked
            {
                string result;
                if (VarName == null)
                {
                    result = "Nothing";
                }
                else
                {
                    Type type = VarName.GetType();
                    bool flag = false;
                    if (type.IsArray)
                    {
                        flag = true;
                        type = type.GetElementType();
                    }
                    string text;
                    if (type.IsEnum)
                    {
                        text = type.Name;
                    }
                    else
                    {
                        switch (Type.GetTypeCode(type))
                        {
                            case TypeCode.DBNull:
                                text = "DBNull";
                                goto IL_138;
                            case TypeCode.Boolean:
                                text = "Boolean";
                                goto IL_138;
                            case TypeCode.Char:
                                text = "Char";
                                goto IL_138;
                            case TypeCode.Byte:
                                text = "Byte";
                                goto IL_138;
                            case TypeCode.Int16:
                                text = "Short";
                                goto IL_138;
                            case TypeCode.Int32:
                                text = "Integer";
                                goto IL_138;
                            case TypeCode.Int64:
                                text = "Long";
                                goto IL_138;
                            case TypeCode.Single:
                                text = "Single";
                                goto IL_138;
                            case TypeCode.Double:
                                text = "Double";
                                goto IL_138;
                            case TypeCode.Decimal:
                                text = "Decimal";
                                goto IL_138;
                            case TypeCode.DateTime:
                                text = "Date";
                                goto IL_138;
                            case TypeCode.String:
                                text = "String";
                                goto IL_138;
                        }
                        text = type.Name;
                        if (type.IsCOMObject && string.CompareOrdinal(text, "__ComObject") == 0)
                        {
                            text = Information.LegacyTypeNameOfCOMObject(VarName, true);
                        }
                    }
                    int num = text.IndexOf('+');
                    if (num >= 0)
                    {
                        text = text.Substring(num + 1);
                    }
                IL_138:
                    if (flag)
                    {
                        Array array = (Array)VarName;
                        if (array.Rank == 1)
                        {
                            text += "[]";
                        }
                        else
                        {
                            text = text + "[" + new string(',', array.Rank - 1) + "]";
                        }
                        text = Information.OldVBFriendlyNameOfTypeName(text);
                    }
                    result = text;
                }
                return result;
            }
        }

        // Microsoft.VisualBasic.Information
        internal static string OldVBFriendlyNameOfTypeName(string typename)
        {
            string text = null;
            checked
            {
                int num = typename.Length - 1;
                if (typename[num] == ']')
                {
                    int num2 = typename.IndexOf('[');
                    if (num2 + 1 == num)
                    {
                        text = "()";
                    }
                    else
                    {
                        text = typename.Substring(num2, num - num2 + 1).Replace('[', '(').Replace(']', ')');
                    }
                    typename = typename.Substring(0, num2);
                }
                string text2 = Information.OldVbTypeName(typename);
                if (text2 == null)
                {
                    text2 = typename;
                }
                string result;
                if (text == null)
                {
                    result = text2;
                }
                else
                {
                    result = text2 + Utils.AdjustArraySuffix(text);
                }
                return result;
            }
        }

        class Utils
        {
            internal static char[] m_achIntlSpace = new char[]
            {
    ' ',
    '\u3000'
            };

            // Microsoft.VisualBasic.CompilerServices.Utils
            internal static string AdjustArraySuffix(string sRank)
            {
                string text = null;
                int i = sRank.Length;
                checked
                {
                    while (i > 0)
                    {
                        char value = sRank[i - 1];
                        switch (value)
                        {
                            case '(':
                                text += ")";
                                break;
                            case ')':
                                text += "(";
                                break;
                            case '*':
                            case '+':
                                goto IL_5F;
                            case ',':
                                text += Conversions.ToString(value);
                                break;
                            default:
                                goto IL_5F;
                        }
                    IL_6C:
                        i--;
                        continue;
                    IL_5F:
                        text = Conversions.ToString(value) + text;
                        goto IL_6C;
                    }
                    return text;
                }
            }

        }

        class Conversions
        {
            public static string ToString(char Value)
            {
                return Value.ToString();
            }
        }

        internal static uint ComputeStringHash(string text)
        {
            uint num = 0u;
            if (text != null)
            {
                num = 2166136261u;
                for (int i = 0; i < text.Length; i++)
                {
                    num = ((uint)text[i] ^ num) * 16777619u;
                }
            }
            return num;
        }

        class Strings
        {
            public static string Mid(string str, int Start, int Length)
            {
                if (Start <= 0)
                {
                    throw new ArgumentException("Argument_GTZero1", nameof(Start));
                }
                if (Length < 0)
                {
                    throw new ArgumentException("Argument_GEZero1", nameof(Length));
                }
                //checked
                {
                    string result;
                    if (Length == 0 || str == null)
                    {
                        result = "";
                    }
                    else
                    {
                        int length = str.Length;
                        if (Start > length)
                        {
                            result = "";
                        }
                        else if (Start + Length > length)
                        {
                            result = str.Substring(Start - 1);
                        }
                        else
                        {
                            result = str.Substring(Start - 1, Length);
                        }
                    }
                    return result;
                }
            }
            public static string Mid(string str, int Start)
            {
                string result;
                try
                {
                    if (str == null)
                    {
                        result = null;
                    }
                    else
                    {
                        result = Strings.Mid(str, Start, str.Length);
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                return result;
            }

            public static string Left(string str, int Length)
            {
                if (Length < 0)
                {
                    throw new ArgumentException("Argument_GEZero1", nameof(Length));
                }
                string result;
                if (Length == 0 || str == null)
                {
                    result = "";
                }
                else if (Length >= str.Length)
                {
                    result = str;
                }
                else
                {
                    result = str.Substring(0, Length);
                }
                return result;
            }
            public static string Trim(string str)
            {
                string result;
                try
                {
                    if (str == null || str.Length == 0)
                    {
                        result = "";
                    }
                    else
                    {
                        char c = str[0];
                        if (c == ' ' || c == '\u3000')
                        {
                            result = str.Trim(Utils.m_achIntlSpace);
                        }
                        else
                        {
                            c = str[checked(str.Length - 1)];
                            if (c == ' ' || c == '\u3000')
                            {
                                result = str.Trim(Utils.m_achIntlSpace);
                            }
                            else
                            {
                                result = str;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                return result;
            }
        }

        class Operators
        {
            public static int CompareString(string Left, string Right, bool TextCompare)
            {
                int result;
                if (Left == Right)
                {
                    result = 0;
                }
                else if (Left == null)
                {
                    if (Right.Length == 0)
                    {
                        result = 0;
                    }
                    else
                    {
                        result = -1;
                    }
                }
                else if (Right == null)
                {
                    if (Left.Length == 0)
                    {
                        result = 0;
                    }
                    else
                    {
                        result = 1;
                    }
                }
                else
                {
                    int num;
                    if (TextCompare)
                    {
                        CultureInfo ci = System.Threading.Thread.CurrentThread.CurrentCulture;
                        num = ci.CompareInfo.Compare(Left, Right, CompareOptions.IgnoreCase | CompareOptions.IgnoreKanaType | CompareOptions.IgnoreWidth);
                    }
                    else
                    {
                        num = string.CompareOrdinal(Left, Right);
                    }
                    if (num == 0)
                    {
                        result = 0;
                    }
                    else if (num > 0)
                    {
                        result = 1;
                    }
                    else
                    {
                        result = -1;
                    }
                }
                return result;
            }
        }
        internal static string OldVbTypeName(string UrtName)
        {
            UrtName = Strings.Trim(UrtName).ToUpperInvariant();
            if (Operators.CompareString(Strings.Left(UrtName, 7), "SYSTEM.", false) == 0)
            {
                UrtName = Strings.Mid(UrtName, 8);
            }
            string text = UrtName;
            uint num = ComputeStringHash(text);
            string result;
            if (num <= 1219467820u)
            {
                if (num <= 268302705u)
                {
                    if (num != 200059396u)
                    {
                        if (num != 225828767u)
                        {
                            if (num == 268302705u)
                            {
                                if (Operators.CompareString(text, "INT16", false) == 0)
                                {
                                    result = "Short";
                                    return result;
                                }
                            }
                        }
                        else if (Operators.CompareString(text, "BYTE", false) == 0)
                        {
                            result = "Byte";
                            return result;
                        }
                    }
                    else if (Operators.CompareString(text, "INT64", false) == 0)
                    {
                        result = "Long";
                        return result;
                    }
                }
                else if (num != 435822185u)
                {
                    if (num != 456003450u)
                    {
                        if (num == 1219467820u)
                        {
                            if (Operators.CompareString(text, "DECIMAL", false) == 0)
                            {
                                result = "Decimal";
                                return result;
                            }
                        }
                    }
                    else if (Operators.CompareString(text, "OBJECT", false) == 0)
                    {
                        result = "Object";
                        return result;
                    }
                }
                else if (Operators.CompareString(text, "SINGLE", false) == 0)
                {
                    result = "Single";
                    return result;
                }
            }
            else if (num <= 2472002000u)
            {
                if (num != 2214109151u)
                {
                    if (num != 2282454687u)
                    {
                        if (num == 2472002000u)
                        {
                            if (Operators.CompareString(text, "DATETIME", false) == 0)
                            {
                                result = "Date";
                                return result;
                            }
                        }
                    }
                    else if (Operators.CompareString(text, "BOOLEAN", false) == 0)
                    {
                        result = "Boolean";
                        return result;
                    }
                }
                else if (Operators.CompareString(text, "INT32", false) == 0)
                {
                    result = "Integer";
                    return result;
                }
            }
            else if (num != 2778687069u)
            {
                if (num != 3751281736u)
                {
                    if (num == 4127814520u)
                    {
                        if (Operators.CompareString(text, "STRING", false) == 0)
                        {
                            result = "String";
                            return result;
                        }
                    }
                }
                else if (Operators.CompareString(text, "DOUBLE", false) == 0)
                {
                    result = "Double";
                    return result;
                }
            }
            else if (Operators.CompareString(text, "CHAR", false) == 0)
            {
                result = "Char";
                return result;
            }
            result = null;
            return result;
        }



        // Microsoft.VisualBasic.Information
        //[SecuritySafeCritical]
        internal static string LegacyTypeNameOfCOMObject(object VarName, bool bThrowException)
        {
            string text = "__ComObject";
            try
            {
                /// new SecurityPermission(SecurityPermissionFlag.UnmanagedCode).Demand();
            }
            catch (StackOverflowException ex)
            {
                throw ex;
            }
            catch (OutOfMemoryException ex2)
            {
                throw ex2;
            }
            catch (System.Threading.ThreadAbortException ex3)
            {
                throw ex3;
            }
            catch (Exception ex4)
            {
                if (bThrowException)
                {
                    throw ex4;
                }
                goto IL_67;
            }
            UnsafeNativeMethods.ITypeInfo typeInfo = null;
            string text2 = null;
            string text3 = null;
            string text4 = null;
            UnsafeNativeMethods.IDispatch dispatch = VarName as UnsafeNativeMethods.IDispatch;
            int num;
            if (dispatch != null && dispatch.GetTypeInfo(0, 1033, out typeInfo) >= 0 && typeInfo.GetDocumentation(-1, out text2, out text3, out num, out text4) >= 0)
            {
                text = text2;
            }
        IL_67:
            if (text[0] == '_')
            {
                text = text.Substring(1);
            }
            return text;
        }

    }

    [ComVisible(false)]/// , SuppressUnmanagedCodeSecurity]
    internal static class UnsafeNativeMethods
    {

        [EditorBrowsable(EditorBrowsableState.Never), Guid("00020403-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [ComImport]
        public interface ITypeComp
        {
            [Obsolete("Bad signature. Fix and verify signature before use.", true)]
            [System.Security.SecurityCritical]
            void RemoteBind([MarshalAs(UnmanagedType.LPWStr)][In] string szName, [MarshalAs(UnmanagedType.U4)][In] int lHashVal, [MarshalAs(UnmanagedType.U2)][In] short wFlags, [MarshalAs(UnmanagedType.LPArray)][Out] UnsafeNativeMethods.ITypeInfo[] ppTInfo, [MarshalAs(UnmanagedType.LPArray)][Out] System.Runtime.InteropServices.ComTypes.DESCKIND[] pDescKind, [MarshalAs(UnmanagedType.LPArray)][Out] System.Runtime.InteropServices.ComTypes.FUNCDESC[] ppFuncDesc, [MarshalAs(UnmanagedType.LPArray)][Out] System.Runtime.InteropServices.ComTypes.VARDESC[] ppVarDesc, [MarshalAs(UnmanagedType.LPArray)][Out] UnsafeNativeMethods.ITypeComp[] ppTypeComp, [MarshalAs(UnmanagedType.LPArray)][Out] int[] pDummy);

            [System.Security.SecurityCritical]
            void RemoteBindType([MarshalAs(UnmanagedType.LPWStr)][In] string szName, [MarshalAs(UnmanagedType.U4)][In] int lHashVal, [MarshalAs(UnmanagedType.LPArray)][Out] UnsafeNativeMethods.ITypeInfo[] ppTInfo);
        }

        [EditorBrowsable(EditorBrowsableState.Never), Guid("00020401-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [ComImport]
        public interface ITypeInfo
        {
            //[System.Security.SecurityCritical]
            [PreserveSig]
            int GetTypeAttr(out IntPtr pTypeAttr);

            //[System.Security.SecurityCritical]
            [PreserveSig]
            int GetTypeComp(out UnsafeNativeMethods.ITypeComp pTComp);

            //[System.Security.SecurityCritical]
            [PreserveSig]
            int GetFuncDesc([MarshalAs(UnmanagedType.U4)][In] int index, out IntPtr pFuncDesc);

            //[System.Security.SecurityCritical]
            [PreserveSig]
            int GetVarDesc([MarshalAs(UnmanagedType.U4)][In] int index, out IntPtr pVarDesc);

            //[System.Security.SecurityCritical]
            [PreserveSig]
            int GetNames([In] int memid, [MarshalAs(UnmanagedType.LPArray)][Out] string[] rgBstrNames, [MarshalAs(UnmanagedType.U4)][In] int cMaxNames, [MarshalAs(UnmanagedType.U4)] out int cNames);

            [Obsolete("Bad signature, second param type should be Byref. Fix and verify signature before use.", true), System.Security.SecurityCritical]
            [PreserveSig]
            int GetRefTypeOfImplType([MarshalAs(UnmanagedType.U4)][In] int index, out int pRefType);

            [Obsolete("Bad signature, second param type should be Byref. Fix and verify signature before use.", true), System.Security.SecurityCritical]
            [PreserveSig]
            int GetImplTypeFlags([MarshalAs(UnmanagedType.U4)][In] int index, [Out] int pImplTypeFlags);

            [System.Security.SecurityCritical]
            [PreserveSig]
            int GetIDsOfNames([In] IntPtr rgszNames, [MarshalAs(UnmanagedType.U4)][In] int cNames, out IntPtr pMemId);

            [Obsolete("Bad signature. Fix and verify signature before use.", true), System.Security.SecurityCritical]
            [PreserveSig]
            int Invoke();

            [System.Security.SecurityCritical]
            [PreserveSig]
            int GetDocumentation([In] int memid, [MarshalAs(UnmanagedType.BStr)] out string pBstrName, [MarshalAs(UnmanagedType.BStr)] out string pBstrDocString, [MarshalAs(UnmanagedType.U4)] out int pdwHelpContext, [MarshalAs(UnmanagedType.BStr)] out string pBstrHelpFile);

            [Obsolete("Bad signature. Fix and verify signature before use.", true), System.Security.SecurityCritical]
            [PreserveSig]
            int GetDllEntry([In] int memid, [In] System.Runtime.InteropServices.ComTypes.INVOKEKIND invkind, [MarshalAs(UnmanagedType.BStr)][Out] string pBstrDllName, [MarshalAs(UnmanagedType.BStr)][Out] string pBstrName, [MarshalAs(UnmanagedType.U2)][Out] short pwOrdinal);

            [System.Security.SecurityCritical]
            [PreserveSig]
            int GetRefTypeInfo([In] IntPtr hreftype, out UnsafeNativeMethods.ITypeInfo pTypeInfo);

            [Obsolete("Bad signature. Fix and verify signature before use.", true), System.Security.SecurityCritical]
            [PreserveSig]
            int AddressOfMember();

            [Obsolete("Bad signature. Fix and verify signature before use.", true), System.Security.SecurityCritical]
            [PreserveSig]
            int CreateInstance([In] ref IntPtr pUnkOuter, [In] ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)][Out] object ppvObj);

            [Obsolete("Bad signature. Fix and verify signature before use.", true), System.Security.SecurityCritical]
            [PreserveSig]
            int GetMops([In] int memid, [MarshalAs(UnmanagedType.BStr)][Out] string pBstrMops);

            [System.Security.SecurityCritical]
            [PreserveSig]
            int GetContainingTypeLib([MarshalAs(UnmanagedType.LPArray)][Out] UnsafeNativeMethods.ITypeLib[] ppTLib, [MarshalAs(UnmanagedType.LPArray)][Out] int[] pIndex);

            [System.Security.SecurityCritical]
            [PreserveSig]
            void ReleaseTypeAttr(IntPtr typeAttr);

            [System.Security.SecurityCritical]
            [PreserveSig]
            void ReleaseFuncDesc(IntPtr funcDesc);

            [System.Security.SecurityCritical]
            [PreserveSig]
            void ReleaseVarDesc(IntPtr varDesc);
        }

        [EditorBrowsable(EditorBrowsableState.Never), Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [ComImport]
        public interface IDispatch
        {
            [Obsolete("Bad signature. Fix and verify signature before use.", true)]//[System.Security.SecurityCritical]
            [PreserveSig]
            int GetTypeInfoCount();

            //[System.Security.SecurityCritical]
            [PreserveSig]
            int GetTypeInfo([In] int index, [In] int lcid, [MarshalAs(UnmanagedType.Interface)] out UnsafeNativeMethods.ITypeInfo pTypeInfo);

            /// [System.Security.SecurityCritical]
            [PreserveSig]
            int GetIDsOfNames();

            /// [System.Security.SecurityCritical]
            [PreserveSig]
            int Invoke();
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        public enum tagSYSKIND
        {
            SYS_WIN16,
            SYS_MAC = 2
        }


        [EditorBrowsable(EditorBrowsableState.Never)]
        public struct tagTLIBATTR
        {
            /// public Guid guid;
            ///
            /// public int lcid;
            ///
            /// public UnsafeNativeMethods.tagSYSKIND syskind;
            ///
            /// [MarshalAs(UnmanagedType.U2)]
            /// public short wMajorVerNum;
            ///
            /// [MarshalAs(UnmanagedType.U2)]
            /// public short wMinorVerNum;
            ///
            /// [MarshalAs(UnmanagedType.U2)]
            /// public short wLibFlags;
        }

        [EditorBrowsable(EditorBrowsableState.Never), Guid("00020402-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [ComImport]
        public interface ITypeLib
        {
            [Obsolete("Bad signature. Fix and verify signature before use.", true)]
            [System.Security.SecurityCritical]
            void RemoteGetTypeInfoCount([MarshalAs(UnmanagedType.LPArray)][Out] int[] pcTInfo);

            [System.Security.SecurityCritical]
            void GetTypeInfo([MarshalAs(UnmanagedType.U4)][In] int index, [MarshalAs(UnmanagedType.LPArray)][Out] UnsafeNativeMethods.ITypeInfo[] ppTInfo);

            [System.Security.SecurityCritical]
            void GetTypeInfoType([MarshalAs(UnmanagedType.U4)][In] int index, [MarshalAs(UnmanagedType.LPArray)][Out] System.Runtime.InteropServices.ComTypes.TYPEKIND[] pTKind);

            [System.Security.SecurityCritical]
            void GetTypeInfoOfGuid([In] ref Guid guid, [MarshalAs(UnmanagedType.LPArray)][Out] UnsafeNativeMethods.ITypeInfo[] ppTInfo);

            [Obsolete("Bad signature. Fix and verify signature before use.", true)]
            [System.Security.SecurityCritical]
            void RemoteGetLibAttr([MarshalAs(UnmanagedType.LPArray)][Out] UnsafeNativeMethods.tagTLIBATTR[] ppTLibAttr, [MarshalAs(UnmanagedType.LPArray)][Out] int[] pDummy);

            [System.Security.SecurityCritical]
            void GetTypeComp([MarshalAs(UnmanagedType.LPArray)][Out] UnsafeNativeMethods.ITypeComp[] ppTComp);

            [Obsolete("Bad signature. Fix and verify signature before use.", true)]
            [System.Security.SecurityCritical]
            void RemoteGetDocumentation(int index, [MarshalAs(UnmanagedType.U4)][In] int refPtrFlags, [MarshalAs(UnmanagedType.LPArray)][Out] string[] pBstrName, [MarshalAs(UnmanagedType.LPArray)][Out] string[] pBstrDocString, [MarshalAs(UnmanagedType.LPArray)][Out] int[] pdwHelpContext, [MarshalAs(UnmanagedType.LPArray)][Out] string[] pBstrHelpFile);

            [Obsolete("Bad signature. Fix and verify signature before use.", true)]
            [System.Security.SecurityCritical]
            void RemoteIsName([MarshalAs(UnmanagedType.LPWStr)][In] string szNameBuf, [MarshalAs(UnmanagedType.U4)][In] int lHashVal, [MarshalAs(UnmanagedType.LPArray)][Out] IntPtr[] pfName, [MarshalAs(UnmanagedType.LPArray)][Out] string[] pBstrLibName);

            [Obsolete("Bad signature. Fix and verify signature before use.", true)]
            [System.Security.SecurityCritical]
            void RemoteFindName([MarshalAs(UnmanagedType.LPWStr)][In] string szNameBuf, [MarshalAs(UnmanagedType.U4)][In] int lHashVal, [MarshalAs(UnmanagedType.LPArray)][Out] UnsafeNativeMethods.ITypeInfo[] ppTInfo, [MarshalAs(UnmanagedType.LPArray)][Out] int[] rgMemId, [MarshalAs(UnmanagedType.LPArray)][In][Out] short[] pcFound, [MarshalAs(UnmanagedType.LPArray)][Out] string[] pBstrLibName);

            [Obsolete("Bad signature. Fix and verify signature before use.", true)]
            [System.Security.SecurityCritical]
            void LocalReleaseTLibAttr();
        }
    }
}
