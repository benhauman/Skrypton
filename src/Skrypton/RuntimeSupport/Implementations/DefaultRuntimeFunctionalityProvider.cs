using Microsoft.CodeAnalysis.CSharp.Syntax;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;
using System;
using System.Collections;
using System.Collections.Frozen;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace Skrypton.RuntimeSupport.Implementations
{
    /// <summary>
    /// Instances of this class should be used only by a single request and so is not written to be thread safe. This is partly because the SETERROR and
    /// CLEARANYERROR methods have no explicit way to be associated with a specific request (which is not a problem if each instance is associated with
    /// one specific request) but also so that it can be explicitly disposed after each request completes to ensure that any unmanaged resources are
    /// cleaned up. VBScript's deterministic garbage collector can tidy up more aggressively, relying upon reference counting, the best that we can do
    /// with the C# code is for this to implement IDiposable and to ensure that everything is tidy when the request completes and Dispose is called.
    /// </summary>
    public sealed class DefaultRuntimeFunctionalityProvider : IProvideVBScriptCompatFunctionalityToIndividualRequests
    {
        /// <summary>
        /// VBScript has a string length limited by its data storage mechanism; each character is represented by two bytes and the index into that
        /// array of data must be an signed int, since it is capped at half of int.MaxValue.. minus one. I'm not sure if the minus one is to do with
        /// a requirement for there to be a null terminator at the end or some VBScript one-based-index weirdness.. or something else.
        /// </summary>
        private const int MAX_VBSCRIPT_STRING_LENGTH = (int.MaxValue / 2) - 1;

        private readonly IRuntimeHost _runtimeHost;
        private readonly IRuntimeLogger _runtimeLogger;
        private readonly IAccessValuesUsingVBScriptRules _valueRetriever;
        private readonly CultureInfo _culture;
        private readonly List<IDisposable> _disposableReferencesToClearAfterTheRequest;
        private readonly Queue<int> _availableErrorTokens;
        private readonly Dictionary<int, ErrorTokenState> _activeErrorTokens;
        private readonly DefaultArithmeticFunctionalityProvider _arithmeticHandler;
        private int _randomSeed;
        private Exception? _trappedErrorIfAny;

        public DefaultRuntimeFunctionalityProvider(IRuntimeHost hostServices, IRuntimeLogger runtimeLogger, IAccessValuesUsingVBScriptRules valueRetriever, CultureInfo culture)
        {
            _runtimeHost = hostServices ?? throw new ArgumentNullException(nameof(hostServices));
            _runtimeLogger = runtimeLogger ?? throw new ArgumentNullException(nameof(runtimeLogger));
            _valueRetriever = valueRetriever ?? throw new ArgumentNullException(nameof(valueRetriever));
            _culture = culture ?? throw new ArgumentNullException(nameof(culture));
            _disposableReferencesToClearAfterTheRequest = new List<IDisposable>();
            _availableErrorTokens = new Queue<int>();
            _activeErrorTokens = new Dictionary<int, ErrorTokenState>();
            _arithmeticHandler = new DefaultArithmeticFunctionalityProvider(valueRetriever);
            DateLiteralParser = DateParser.ForCulture(culture);
            _randomSeed = 0; // Doesn't really matter what this is initially, just that it's always the same
            _trappedErrorIfAny = null;
        }

        void IProvideVBScriptCompatFunctionalityToIndividualRequests.ValidateDateTimeLiteralAgainstCurrentCulture(params Tuple<string, int>[] literalsToValidate)
        {
            foreach (var dateLiteralValueAndLineNumbers in literalsToValidate)
            {
                try { DateLiteralParser.Parse(dateLiteralValueAndLineNumbers.Item1, _culture); }
                catch
                {
                    throw new SyntaxError($@"Invalid date literal '{dateLiteralValueAndLineNumbers.Item1}' on line:{dateLiteralValueAndLineNumbers.Item2}");
                }
            }
        }

        private readonly Dictionary<string, Func<string?, object>> _objectCreateFactories = new Dictionary<string, Func<string?, object>>(StringComparer.OrdinalIgnoreCase);

        private static object HandlePostInitializationHandler(string progId, object objectInstance)
        {
            // The behavior comes from three different layers that were never fully documented together:
            // * 1/3) VBScript’s automatic type coercion rules. Key point: VBScript will coerce strings to booleans, numbers, objects, etc. when calling COM methods. https://learn.microsoft.com/en-us/previous-versions//d1wf56tt(v=vs.85)
            // * 2/3) MSXML’s COM overloading rules. VBScript uses IDispatch::Invoke with very permissive rules. https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-invoke
            //  -> unwraps COM objects automatically
            //  -> chooses the correct overloaded COM method
            //  -> retries calls with different type coercions
            //  -> suppresses many COM errors
            // * 3/3) The IDispatch binder inside Windows Script Host

            if (string.Equals(progId, "Msxml2.DOMDocument", StringComparison.OrdinalIgnoreCase))
            {
                // !!!! VBScript silently sets:             xmlDoc.preserveWhiteSpace = False
                //https://zetcode.com/vbscript/dom-msxml-domdocument/
                // https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms757828(v=vs.85)
                //IDispatchAccess.CallPropertySet(objectInstance, "preserveWhiteSpace", new object[] { false });
                //objectInstance.InvokeMember("preserveWhiteSpace", BindingFlags.SetProperty, null, xmlDoc, new object[] { false });
                //domType.InvokeMember("preserveWhiteSpace", BindingFlags.SetProperty, null, xmlDoc, new object[] { false });

                // Late‑bind using reflection
#pragma warning disable CA1304 // Specify CultureInfo
                object elem = objectInstance.GetType().InvokeMember("preserveWhiteSpace", BindingFlags.SetProperty, null, objectInstance, new object[] { false });
#pragma warning restore CA1304 // Specify CultureInfo
            }
            //IntPtr pDispatch = Marshal.GetIDispatchForObject(objectInstance);
            //var disp = Marshal.GetComInterfaceForObject<IDispatchAccess.IDispatch>();
            //IDispatch disp = (IDispatch)Marshal.GetObjectForIUnknown(pDispatch);
            return objectInstance;
        }

        public void RegisterObjectCreateFactory(string progId, Func<string?, object> factory)
        {
            if (string.IsNullOrEmpty(progId))
                throw new ArgumentException("Value can not be null or empty", nameof(progId));
            _objectCreateFactories[progId] = factory ?? throw new ArgumentNullException(nameof(factory));
        }

        private enum ErrorTokenState
        {
            OnErrorResumeNext,
            OnErrorGoto0
        }

        ~DefaultRuntimeFunctionalityProvider()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                foreach (IDisposable disposableResource in _disposableReferencesToClearAfterTheRequest)
                {
#pragma warning disable CA1031 // Do not catch general exception types
                    try { disposableResource.Dispose(); }
                    catch { }
#pragma warning restore CA1031 // Do not catch general exception types
                }
            }
        }

        // Arithmetic operators
        public object ADD(object l, object r) { return _arithmeticHandler.ADD(l, r); }
        public object SUBT(object o) { return _arithmeticHandler.SUBT(o); }
        public object SUBT(object l, object r) { return _arithmeticHandler.SUBT(l, r); }
        public object MULT(object l, object r) { return _arithmeticHandler.MULT(l, r); }
        public object DIV(object l, object r) { return _arithmeticHandler.DIV(l, r); }
        public int INTDIV(object l, object r) { return _arithmeticHandler.INTDIV(l, r); }
        public double POW(object l, object r) { return _arithmeticHandler.POW(l, r); }
        public object MOD(object l, object r) { return _arithmeticHandler.MOD(l, r); }

        // String concatenation
        public object CONCAT(object? l, object? r)
        {
            // Try to get both values as value types - if either is Nothing or an object without a default parameterless member, then it's an ObjectVariableNotSetException
            // or ObjectDoesNotSupportPropertyOrMemberException, resp. If one is an object WITH a default parameterless member, but the value of that member is not a value
            // type, then it's a TypeMismatchException. (If the values are value types to begin with, or are objects with default parameterless member that has a value
            // type, then there's nothing to worry about).
            bool parameterLessDefaultMemberWasAvailable;
            if (!TryVAL(l, out parameterLessDefaultMemberWasAvailable, out l))
            {
                if (parameterLessDefaultMemberWasAvailable)
                    throw new TypeMismatchException("left parameterLessDefaultMemberWasAvailable");
                if (IsVBScriptNothing(l))
                    throw new ObjectVariableNotSetException("left is nothing");
                throw new ObjectDoesNotSupportPropertyOrMemberException($"l:{l}");
            }
            if (!TryVAL(r, out parameterLessDefaultMemberWasAvailable, out r))
            {
                if (parameterLessDefaultMemberWasAvailable)
                    throw new TypeMismatchException("right parameterLessDefaultMemberWasAvailable");
                if (IsVBScriptNothing(r))
                    throw new ObjectVariableNotSetException("right is nothing");
                throw new ObjectDoesNotSupportPropertyOrMemberException($"r:{r}");
            }
            if ((l == DBNull.Value) && (r == DBNull.Value))
                return DBNull.Value;
            string lString = (l == DBNull.Value) ? "" : _valueRetriever.STR(l);
            string rString = (r == DBNull.Value) ? "" : _valueRetriever.STR(r);
            if ((lString.Length + rString.Length) > MAX_VBSCRIPT_STRING_LENGTH)
                throw new OutOfStringSpaceException($"l:{lString.Length}, r:{rString.Length}");
            return lString + rString;
        }

        /// <summary>
        /// This may never be called with less than two values (otherwise an exception will be thrown)
        /// </summary>
        public object CONCAT(params object[] values)
        {
            if (values == null)
                throw new ArgumentNullException(nameof(values));

            if (values.Length < 2)
                throw new ArgumentException("There must be at least two values specified for the CONCAT operation");

            // Concatenate the first two values (using the standard two-value version of the method) and then concatenate each further values on to
            // this accumulator. This could very likely be done in a more efficient manner by recursively splitting the array of values but this will
            // do for now.
            object combinedValue = CONCAT(values[0], values[1]);
            foreach (object additionalValue in values.Skip(2))
                combinedValue = CONCAT(combinedValue, additionalValue);
            return combinedValue;
        }

        // Logical operators (these return VBScript Null if one or both sides of the comparison are VBScript Null)
        // - Read http://blogs.msdn.com/b/ericlippert/archive/2004/07/15/184431.aspx
        public object NOT(object o)
        {
            Tuple<IEnumerable<int?>, Func<int, object>, Type> bitwiseOperationValues = GetForBitwiseOperations("'Not'", o);
            int? valueToNot = bitwiseOperationValues.Item1.Single();
            if (valueToNot == null)
            {
                // GetForBitwiseOperations returns nullable int values - since VBScript's Empty (ie. C#'s null) will be interpreted as zero then any
                // null values here mean VBScript's null (ie. DBNull.Value), and so that is what must be returned from this function
                return DBNull.Value;
            }
            return bitwiseOperationValues.Item2(~valueToNot.Value); // Note: VBScript's Not operation is bitwise, not logical (so the ~ operator is used)
        }
        public object AND(object l, object r)
        {
            return ANDCore(l, () => r);
        }
        public bool ANDe2(bool evaluationResult) => evaluationResult;
        public object ANDe2z(object l, Func<object> rp)
        {
            if (rp == null) throw new ArgumentNullException(nameof(rp));
            return ANDCore(l, rp);
        }
        private object ANDCore(object l, Func<object> rp)
        {
            Tuple<IEnumerable<int?>, Func<int, object>, Type> bitwiseOperationValuesL = GetForBitwiseOperations("'And'", l);
            int? left = bitwiseOperationValuesL.Item1.First();
            if (left == null)
            {
                // If GetForBitwiseOperations returns null values then it means there were VBScript Null values provided. When AND'ing, if either (or both)
                // values are Null then Null is returned.
                return DBNull.Value;
            }

            //if (left.Value ==)


            object r = rp();
            Tuple<IEnumerable<int?>, Func<int, object>, Type> bitwiseOperationValuesR = GetForBitwiseOperations("'And'", r);
            int? right = bitwiseOperationValuesR.Item1.First();
            if (right == null)
            {
                // If GetForBitwiseOperations returns null values then it means there were VBScript Null values provided. When AND'ing, if either (or both)
                // values are Null then Null is returned.
                return DBNull.Value;
            }

            return bitwiseOperationValuesR.Item2(left.Value & right.Value);
        }
        private object ANDCoreOld(object l, object r)
        {
            Tuple<IEnumerable<int?>, Func<int, object>, Type> bitwiseOperationValues = GetForBitwiseOperations("'And'", l, r);
            int? left = bitwiseOperationValues.Item1.First();
            int? right = bitwiseOperationValues.Item1.Skip(1).Single();
            if ((left == null) || (right == null))
            {
                // If GetForBitwiseOperations returns null values then it means there were VBScript Null values provided. When AND'ing, if either (or both)
                // values are Null then Null is returned.
                return DBNull.Value;
            }
            return bitwiseOperationValues.Item2(left.Value & right.Value);
        }
        public object OR(object? l, object? r)
        {
            Tuple<IEnumerable<int?>, Func<int, object>, Type> bitwiseOperationValues = GetForBitwiseOperations("'Or'", l, r);
            int? left = bitwiseOperationValues.Item1.First();
            int? right = bitwiseOperationValues.Item1.Skip(1).Single();
            if ((left == null) && (right == null))
            {
                // If GetForBitwiseOperations returns null values then it means there were VBScript Null values provided. When OR'ing, if one value is Null
                // but the other isn't, then the non-Null value is returned - only if both values are Null is Null returned.
                return DBNull.Value;
            }
            else if (left == null)
                return right!;
            else if (right == null)
                return left!;
            return bitwiseOperationValues.Item2(left.Value | right.Value);
        }
        public object XOR(object l, object r)
        {
            Tuple<IEnumerable<int?>, Func<int, object>, Type> bitwiseOperationValues = GetForBitwiseOperations("'Xor'", l, r);
            int? left = bitwiseOperationValues.Item1.First();
            int? right = bitwiseOperationValues.Item1.Skip(1).Single();
            if ((left == null) || (right == null))
            {
                // If GetForBitwiseOperations returns null values then it means there were VBScript Null values provided. When XOR'ing, if either (or both)
                // values are Null then Null is returned.
                return DBNull.Value;
            }
            return bitwiseOperationValues.Item2(left.Value ^ right.Value);
        }

        // Comparison operators (these return VBScript Null if one or both sides of the comparison are VBScript Null)
        /// <summary>
        /// This will return DBNull.Value or boolean value. VBScript has rules about comparisons between "hard-typed" values (aka literals), such
        /// that a comparison between (a = 1) requires that the value "a" be parsed into a numeric value (resulting in a Type Mismatch if this is
        /// not possible). However, this logic must be handled by the translation process before the EQ method is called. Both comparison values
        /// must be treated as non-object-references, so if they are not when passed in then the method will try to retrieve non-object values
        /// from them - if this fails then a Type Mismatch error will be raised. If there are no issues in preparing both comparison values,
        /// this will return DBNull.Value if either value is DBNull.Value and a boolean otherwise.
        /// </summary>
        public object EQ(object l, object r) { return ToVBScriptNullable(EQ_Internal(l, r)); }
        private bool? EQ_Internal(object? l, object? r)
        {
            // Both sides of the comparison must be simple VBScript values (ie. not object references) - pushing both values through VAL will handle
            // that (an exception will be raised if this operation fails and the value will not be affect if it was already an acceptable type)
            l = _valueRetriever.VAL(l);
            r = _valueRetriever.VAL(r);

            // Let's get the outliers out of the way; VBScript Null and Empty..
            if ((l == DBNull.Value) || (r == DBNull.Value))
                return null; // If one or both sides of the comparison are "Null" then this is what is returned
            if ((l == null) && (r == null))
                return true; // If both sides are Empty then they are considered to match
            else if ((l == null) || (r == null))
            {
                // The default values of VBScript primitives (number, strings and booleans) are considered to match Empty
                object nonNullValue = l ?? r!;
#pragma warning disable CA1820 // Test for empty strings using string length
                if ((IsDotNetNumericType(nonNullValue) && (Convert.ToDouble(nonNullValue, CultureInfo.InvariantCulture)) == 0)
                || ((nonNullValue as string) == "")
                || ((nonNullValue is bool) && !(bool)nonNullValue))
                    return true;
#pragma warning restore CA1820 // Test for empty strings using string length
                return false;
            }

            // Booleans have some funny behaviour in that they will match values of other types (numbers, but not strings unless string literals
            // are in the comparison, which is not logic that this method has to deal with). If one of the values is a boolean and the other isn't,
            // and none of the special cases are met, then there must not be a match.
            if ((l is bool) && (r is bool))
                return (bool)l == (bool)r;
            else if ((l is bool) || (r is bool))
            {
                bool boolValue = (bool)((l is bool) ? l : r);
                object nonBoolValue = (l is bool) ? r : l;
                if (!IsDotNetNumericType(nonBoolValue))
                    return false;
                return (boolValue && (Convert.ToDouble(nonBoolValue, CultureInfo.InvariantCulture) == -1)) || (!boolValue && (Convert.ToDouble(nonBoolValue, CultureInfo.InvariantCulture) == 0));
            }

            // Now consider numbers on one or both sides - all special cases are out of the way now so they're either equal or they're not (both
            // sides must be numbers, otherwise it's a non-match)
            if (IsDotNetNumericType(l) && IsDotNetNumericType(r))
                return Convert.ToDouble(l, CultureInfo.InvariantCulture) == Convert.ToDouble(r, CultureInfo.InvariantCulture);
            else if (IsDotNetNumericType(l) || IsDotNetNumericType(r))
            {
                // lubo:
                if (IsDotNetNumericType(l))
                {
                    // left is numeric
                    if (r is string)
                    {
                        double numR;
                        if (Double.TryParse((string)r, out numR))
                        {
                            if (Convert.ToDouble(l, CultureInfo.InvariantCulture) == numR)
                                return true;
                            return false;
                        }
                    }
                }
                else
                {
                    // right is numeric
                    if (l is string)
                    {
                        double numL;
                        if (Double.TryParse((string)l, out numL))
                        {
                            if (numL == Convert.ToDouble(r, CultureInfo.InvariantCulture))
                                return true;
                            return false;
                        }
                    }
                }
                return false;
            }

            // Now do the same for strings and then dates - same deal; they must have consistent types AND values
            if ((l is string) && (r is string))
                return (string)l == (string)r;
            else if ((l is string) || (r is string))
                return false;
            if ((l is DateTime) && (r is DateTime))
                return (DateTime)l == (DateTime)r;

            // Frankly, if we get here then I have no idea what's happened. It will be much easier to identify issues (if any are encountered) if an
            // exception is raised rather than a false response return
            throw new NotSupportedException("Don't know how to compare values of type " + TYPENAME(l) + " and " + TYPENAME(r));
        }

        public object? NOTEQ(object l, object r)
        {
            // We can just reverse EQ_Internal's result here, unless it returns null - if it returns null then it means that comparison was not
            // meaningful (one or both sides were DBNull.Value) and so DBNull.Value should be returned.
            bool? opposingEqualityResult = EQ_Internal(l, r);
            if (opposingEqualityResult == null)
                return null;
            return !opposingEqualityResult.Value;
        }

        public object LT(object l, object r) { return ToVBScriptNullable(LT_Internal(l, r, allowEquals: false)); }
        public object LTE(object l, object r) { return ToVBScriptNullable(LT_Internal(l, r, allowEquals: true)); }
        /// <summary>
        /// This takes the logic from LT but throws an exception if a DBNull.Value is taken as part of the comparison (which is how it is able to
        /// return a boolean, rather than an object - which LT has to since it may return a boolean OR DBNull.Value)
        /// </summary>
        public bool StrictLT(object l, object r)
        {
            bool? result = LT_Internal(l, r, allowEquals: false);
            if (result == null)
                throw new InvalidUseOfNullException("result ");
            return result.Value;
        }
        /// <summary>
        /// This takes the logic from LTE but throws an exception if a DBNull.Value is taken as part of the comparison (which is how it is able to
        /// return a boolean, rather than an object - which LTE has to since it may return a boolean OR DBNull.Value)
        /// </summary>
        public bool StrictLTE(object l, object r)
        {
            bool? result = LT_Internal(l, r, allowEquals: true);
            if (result == null)
                throw new InvalidUseOfNullException("result is null");
            return result.Value;
        }
        private bool? LT_Internal(object? l, object? r, bool allowEquals)
        {
            // Both sides of the comparison must be simple VBScript values (ie. not object references) - pushing both values through VAL will handle
            // that (an exception will be raised if this operation fails and the value will not be affect if it was already an acceptable type)
            l = _valueRetriever.VAL(l);
            r = _valueRetriever.VAL(r);

            // If one or both sides of the comparison as VBScript Null then that is what is returned
            if ((l == DBNull.Value) || (r == DBNull.Value))
                return null;

            // Check the equality case first, since there may be an early exit we can make (this should return a true or false since the "Null" cases
            // have been handled) - if the values ARE equal then either return true (if allowEquals is true) or false (if allowEquals is false). If
            // not then we'll have to do more work.
            bool? eq = EQ_Internal(l, r);
            if (eq == null)
                throw new NotSupportedException("Don't know how to compare values of type " + TYPENAME(l) + " and " + TYPENAME(r));
            if (eq.Value)
                return allowEquals;

            // Deal with string special cases next - if both are strings then perform a string comparison. If only one is a string, and it is not blank,
            // then that value is bigger (so if it's on the left then return false and if it's on the right then return true). Blank strings get special
            // handling and are effectively treated as zero (see further down).
            string? lString = l as string;
            string? rString = r as string;
            if ((lString != null) && (rString != null))
            {
                int? stringComparisonResult = STRCOMP_Internal(lString, rString, 0);
                if ((stringComparisonResult == null) || (stringComparisonResult.Value == 0))
                    throw new NotSupportedException("Don't know how to compare values of type " + TYPENAME(l) + " and " + TYPENAME(r));
                return stringComparisonResult.Value < 0;
            }
#pragma warning disable CA1820 // Test for empty strings using string length
            if ((lString != null) && (lString != ""))
                return false;
            if ((rString != null) && (rString != ""))
                return true;

            // Now we should only have values which can treated as numeric
            // - Actual numbers
            // - Booleans (which return zero or minus one when passed through CDBL)
            // - Null aka VBScript Empty (which returns zero when passed through CDBL)
            // - Blank strings (which can not be passed through CDBL without causing an error, but which we can treat as zero)
            double lNumeric = (lString == "") ? 0 : CDBL_Precise(l);
            double rNumeric = (rString == "") ? 0 : CDBL_Precise(r);
#pragma warning restore CA1820 // Test for empty strings using string length
            return lNumeric < rNumeric;
        }

        public object GT(object l, object r) { return ToVBScriptNullable(GT_Internal(l, r, allowEquals: false)); }
        public object GTE(object l, object r) { return ToVBScriptNullable(GT_Internal(l, r, allowEquals: true)); }
        /// <summary>
        /// This takes the logic from GT but throws an exception if a DBNull.Value is taken as part of the comparison (which is how it is able to
        /// return a boolean, rather than an object - which GT has to since it may return a boolean OR DBNull.Value)
        /// </summary>
        public bool StrictGT(object l, object r)
        {
            bool? result = GT_Internal(l, r, allowEquals: false);
            if (result == null)
                throw new InvalidUseOfNullException("result is null");
            return result.Value;
        }
        /// <summary>
        /// This takes the logic from GTE but throws an exception if a DBNull.Value is taken as part of the comparison (which is how it is able to
        /// return a boolean, rather than an object - which GTE has to since it may return a boolean OR DBNull.Value)
        /// </summary>
        public bool StrictGTE(object l, object r)
        {
            bool? result = GT_Internal(l, r, allowEquals: true);
            if (result == null)
                throw new InvalidUseOfNullException("result is null");
            return result.Value;
        }
        private bool? GT_Internal(object l, object r, bool allowEquals)
        {
            // This can just LT_Internal, rather than trying to deal with too much logic itself. When calling LT_Internal, the "allowEquals" value must be
            // the opposite of what we have here - if we are considering GTE then we want !LT (since the equality case should be a match here and not a
            // result which is inverted), if we are considering GT here then we want !LTE (since then equality case would not be a match and LTE would
            // return true for equal l and r values and we would want to invert that result). If LT_Internal returns null, then it means that the
            // comparison is not meaningful (in other words, DBNull.Value was on one or both sides and so DBNull.Value should be returned for
            // any comparison - whether EQ, NOTEQ, LT, GT, etc..)
            bool? opposingLessThanResult = LT_Internal(l, r, !allowEquals);
            if (opposingLessThanResult == null)
                return null;
            return !opposingLessThanResult.Value;
        }

        public bool IS(object l, object r)
        {
            if (IsVBScriptNothing(l) && IsVBScriptNothing(r))
                return true;
            return _valueRetriever.OBJ(l, "'Is'") == _valueRetriever.OBJ(r, "'Is'");
        }
        public object EQV(object l, object r) { throw new NotImplementedException(); }
        public object IMP(object l, object r) { throw new NotImplementedException(); }

        // Builtin functions - TODO: These are not fully specified yet (eg. LEFT requires more than one parameter and INSTR requires multiple parameters and
        // overloads to deal with optional parameters)
        // - Type conversions
        public byte CBYTE(object value) { return CBYTECore(value, "'CByte'"); }

        private static readonly Encoding Windows1252 = InitializeEncodingWindows1252();

        private static Encoding InitializeEncodingWindows1252()
        {
            Encoding.RegisterProvider(provider: CodePagesEncodingProvider.Instance); // from nuget package 'System.Text.Encoding.CodePages'
            Encoding windows1252 = Encoding.GetEncoding("Windows-1252");
            return windows1252;
        }

        private byte CBYTECore(object? value, string exceptionMessageForInvalidContent)
        {
            if (value == null)
            {
                return 0;
            }
            if (value == DBNull.Value)
            {
                return 0;
            }
            if (value is string and "")
            {
                return 0;
            }

            object? valueX = _valueRetriever.VAL(value, exceptionMessageForInvalidContent);
            //return GetAsNumber<byte>(valueX, exceptionMessageForInvalidContent, Convert.ToByte);
            return GetAsNumber<byte>(valueX, exceptionMessageForInvalidContent, ConvertToByte, true);
        }
#pragma warning disable SA1204 // Static elements should appear before instance elements
        private static byte ConvertToByte(object value)
#pragma warning restore SA1204 // Static elements should appear before instance elements
        {
            // Convert the int code point to char
            byte b;
            if (value is int valueInt32)
            {
                char c = (char)(int)valueInt32; // 8364 -> '€'
                b = Windows1252.GetBytes(new[] { c })[0];
                return b;
            }
            else
            {
                try
                {
                    b = Convert.ToByte(value, CultureInfo.InvariantCulture); // double, null, empty
                    return b;
                }
                catch (System.OverflowException e) // for -0.6
                {
                    throw new InvalidProcedureCallOrArgumentException("'CHR'", e);
                }
            }
        }
        public bool CBOOL(object value) { return BOOL(value, "'CBool'"); }
        private bool CBOOL(object value, string exceptionMessageForInvalidContent) { return _valueRetriever.BOOL(value, exceptionMessageForInvalidContent); }
        public decimal CCUR(object value) { return CCUR(value, "'CCur'"); }
        private decimal CCUR(object value, string exceptionMessageForInvalidContent)
        {
            decimal currencyValue = GetAsNumber<decimal>(value, exceptionMessageForInvalidContent, Convert.ToDecimal);
            if ((currencyValue < VBScriptConstants.MinCurrencyValue) || (currencyValue > VBScriptConstants.MaxCurrencyValue))
                throw new VBScriptOverflowException("'CCur' (" + currencyValue.ToString(CultureInfo.InvariantCulture) + ")");
            return currencyValue;
        }
        public double CDBL(object value)
        {
            // When working with CDBL / CDATE, it seemed like some precision was getting lost when values are passed back and forth through them - eg. if 40000.01
            // is passed into CDATE and then back through CDBL then 40000.01 should come back out. This can be emulate with a double-decimal-double conversion so
            // that these sort of translations seem consistent. However, when trying to parse numbers for other purposes internally, this shouldn't be done since
            // the precision may be important (there are some edge cases in DATEADD where this applies - eg. adding 1.999999999999999 seconds (15x 9s) to a date
            // results in 1 second being added, while adding 1.9999999999999999 seconds (16x 9s) results in 2 seconds being added. I don't think there's a way
            // to perfectly recreate all of VBScript's precision oddities in all cases, so I'm just trying to stick to it being consistent in as many places
            // as possible (which unfortunately means that there's a discrepancy between the internal and public CDBL implementations here).
            return (double)((decimal)CDBL_Precise(value, "'CDbl'"));
        }
        private double CDBL_Precise(object? value) { return CDBL_Precise(value, null); }
        private double CDBL_Precise(object? value, string? optionalExceptionMessageForInvalidContent)
        {
            return GetAsNumber<double>(value, optionalExceptionMessageForInvalidContent, Convert.ToDouble);
        }
        public DateTime CDATE(object value) { return CDATECore(value, "'CDate'"); }
        private DateTime CDATECore(object? value, string exceptionMessageForInvalidContent)
        {
            if (string.IsNullOrWhiteSpace(exceptionMessageForInvalidContent))
                throw new ArgumentException("Null/blank exceptionMessageForInvalidContent specified");

            // Hand off all parsing here to the base valueRetriever.DATE to avoid code duplication
            return _valueRetriever.DATE(value, exceptionMessageForInvalidContent);
        }
        public Int16 CINT(object value) { return CINT(value, "'CInt'"); }
        private Int16 CINT(object value, string exceptionMessageForInvalidContent) { return GetAsNumber<Int16>(value, exceptionMessageForInvalidContent, Convert.ToInt16); }
        public int CLNG(object? value) { return CLNG(value, "'CLng'"); }
        public int CLNG(object? value, string exceptionMessageForInvalidContent) { return GetAsNumber<int>(value, exceptionMessageForInvalidContent, Convert.ToInt32); }
        public float CSNG(object? value) { return CSNG(value, "'CSng'"); }
        private float CSNG(object? value, string exceptionMessageForInvalidContent) { return GetAsNumber<float>(value, exceptionMessageForInvalidContent, Convert.ToSingle); }
        public string CSTR(object value) { return CSTR(value, "'CStr'"); }
        private string CSTR(object value, string exceptionMessageForInvalidContent)
        {
            if (string.IsNullOrWhiteSpace(exceptionMessageForInvalidContent))
                throw new ArgumentException("Null/blank exceptionMessageForInvalidContent specified");

            // Hand off all parsing here to the base valueRetriever.STR to avoid code duplication
            return _valueRetriever.STR(value, exceptionMessageForInvalidContent);
        }
#pragma warning disable CA1720 // Identifier contains type name
        public object INT(object? value)
#pragma warning restore CA1720 // Identifier contains type name
        {
            value = _valueRetriever.VAL(value);

            // Deal with null-like cases
            if (value == DBNull.Value)
                return value;
            if (value == null)
                return (Int16)0;

            // Deal with value type that don't need changing
            if ((value is byte) || (value is Int16) || (value is Int32))
                return value;

            // Deal with a couple of simple case; boolean -> Int16 and Date -> Date (though without any time component)
            if (value is bool)
                return (Int16)((bool)value ? -1 : 0);
            if (value is DateTime)
                return ((DateTime)value).Date;
            bool valueWasSingle = value is Single;
            bool valueWasDecimal = value is Decimal;
            double valueDouble = GetAsNumber<double>(value, "'Int' (" + value.ToString() + ")", Convert.ToDouble);
            valueDouble = Math.Floor(valueDouble);
            if (valueWasSingle)
                return (Single)valueDouble;
            else if (valueWasDecimal)
                return (Decimal)valueDouble;
            return valueDouble;
        }
#pragma warning disable CA1720 // Identifier contains type name
        public string STRING(object? numberOfTimesToRepeat, object? character)
#pragma warning restore CA1720 // Identifier contains type name
        {
#pragma warning disable CA1820 // Test for empty strings using string length

            character = _valueRetriever.VAL(character, "'String'");
            numberOfTimesToRepeat = _valueRetriever.VAL(numberOfTimesToRepeat, "'String'");
            if ((numberOfTimesToRepeat == DBNull.Value) || (character == DBNull.Value))
                throw new InvalidUseOfNullException("'String'");
            int numberOfTimesToRepeatNumber;
            if (numberOfTimesToRepeat == null)
                numberOfTimesToRepeatNumber = 0;
            else
            {
                numberOfTimesToRepeatNumber = CLNG(numberOfTimesToRepeat, "'String'");
                if (numberOfTimesToRepeatNumber < 0)
                    throw new InvalidProcedureCallOrArgumentException("'String'");
            }
            char characterChar;
            if (character == null)
                characterChar = '\0';
            else
            {
                string? characterString = character as string;
                if (characterString != null)
                {
                    if (characterString == "")
                        throw new InvalidProcedureCallOrArgumentException("'String'");
                    characterChar = characterString[0];
                }
                else
                {
                    short characterCode = CINT(character, "'String'");
                    if (characterCode > 256)
                        characterCode = (short)(characterCode % 256);
                    else if (characterCode < 0)
                    {
                        double numberOf256sToAdd = Math.Ceiling(Math.Abs((double)characterCode / 256));
                        characterCode += (short)(numberOf256sToAdd * 256);
                    }
                    characterChar = (char)characterCode;
                }
            }
#pragma warning restore CA1820 // Test for empty strings using string length

            if (numberOfTimesToRepeatNumber > MAX_VBSCRIPT_STRING_LENGTH)
                throw new OutOfStringSpaceException("'String'");
            if (numberOfTimesToRepeatNumber == 0)
                return "";
            return new string(characterChar, numberOfTimesToRepeatNumber);
        }
        // - Randomisation functions
        public void RANDOMIZE() { RANDOMIZE(DateTime.Now.TimeOfDay.TotalSeconds); }
        public void RANDOMIZE(object seed)
        {
            // The very first time that RANDOMIZE is called with a particular value, the following sequence of random numbers that is produced should be the same. However, if
            // RANDOMIZE is called later with the same seed number then there is no guarantee that the same sequence will be generated. This is why the new seed value that is
            // calculated here takes into account the RANDOMIZE value *and* the current seed. See the note "Repeatedly passing the same number to Randomize doesn’t cause Rnd
            // to repeat the same sequence of random numbers." from https://www.safaribooksonline.com/library/view/vbscript-in-a/1565927206/re148.htm
            // Note: The seed should only have the precision of the Single type (in VBScript, though it's the same in .NET) ad so precision after a certain point will have no
            // effect. For example, the following two seeds will result in the same sequence being generated:
            //   Randomize 1.111111
            //   Randomize 1.1111111
            int valueFromSeed = CSNG(seed).GetHashCode();
            double randomValueFromCurrentSeed = GenerateRandomDouble();
            _randomSeed = (valueFromSeed * randomValueFromCurrentSeed).GetHashCode();
        }
        private double GenerateRandomDouble()
        {
#pragma warning disable CA5394 // Do not use insecure randomness
            return new Random(_randomSeed).NextDouble();
#pragma warning restore CA5394 // Do not use insecure randomness
        }
        // - Number functions
        public object ABS(object? value)
        {
            value = _valueRetriever.VAL(value, "'Abs'");
            if (value is bool)
                return (bool)value ? (Int16)1 : (Int16)0;
            if (value is byte)
                return value;
            if (value is Int16)
                return (Int16)Math.Abs((Int16)value);
            if (value is Int32)
                return (Int16)Math.Abs((Int16)value);
            if (value is decimal)
                return (decimal)Math.Abs((decimal)value);
            return Math.Abs(CDBL_Precise(value, "'Abs'"));
        }
        public object ATN(object value)
        {
            // TODO: Tests need to confirm that double precision is used (eg. COS returns different values for 1.111111 and 1.1111111)
            double radians = CDBL_Precise(value, "'Atn'");
            return Math.Atan(radians);
        }
        public object COS(object value)
        {
            // TODO: Tests need to confirm that double precision is used (eg. COS returns different values for 1.111111 and 1.1111111)
            double radians = CDBL_Precise(value, "'Cos'");
            return Math.Cos(radians);
        }
        public object EXP(object value) { throw new NotImplementedException(); }
        public object FIX(object value) { throw new NotImplementedException(); }
        public object LOG(object value) { throw new NotImplementedException(); }
        public object OCT(object value) { throw new NotImplementedException(); }
        public float RND()
        {
            return RND(1); // Any value greater than zero passed to RND will just get the next random number, so calling RND(1) should be the same as VBScript calling just "RND()"
        }
        public float RND(object? value)
        {
            value = _valueRetriever.VAL(value, "'Rnd'");
            if (value == DBNull.Value)
                throw new InvalidUseOfNullException("RND argument may not be null");

            // See https://msdn.microsoft.com/en-us/library/e566zd96(v=vs.84).aspx
            float valueAsSingle = CSNG(value);
            if (valueAsSingle == 0)
            {
                // Return the most recently generated number (if called repeatedly, the same number should be returned - so no changes to the seed should be mde
                return (float)GenerateRandomDouble();
            }
            else if (valueAsSingle < 0)
            {
                // Use the provided value as the seed (always return the same number and change the sequence for any subsequent numbers - ie. don't just use the
                // value as the seed here but update the global seed)
                _randomSeed = valueAsSingle.GetHashCode();
                return (float)GenerateRandomDouble();
            }

            // Greater than zero => next random number in the sequence (this should move the sequence along, so we need to change the global seed before getting
            // the next number - if RND(0) is called next then the same number will be returned, as required for compatibility)
#pragma warning disable CA5394 // Do not use insecure randomness
            _randomSeed = new Random(_randomSeed).Next();
#pragma warning restore CA5394 // Do not use insecure randomness
            return (float)GenerateRandomDouble();
        }
        public object ROUND(object value) => ROUNDCore(value, 0);
        public object ROUND(object value, object decimals)
        {
            int nDecimals = GetAsNumber<int>(decimals, "'ROUND'", Convert.ToInt32);
            return ROUNDCore(value, nDecimals);
        }
        private decimal ROUNDCore(object value, int nDecimals)
        {
            decimal decimalValue = GetAsNumber<decimal>(value, "'ROUND'", Convert.ToDecimal);
            if (nDecimals < 0)
                throw new ArgumentOutOfRangeException(nameof(nDecimals), "Must be >= 0 to match VBScript.");

            return Math.Round(decimalValue, nDecimals, MidpointRounding.ToEven);
        }
        public object SGN(object value) { throw new NotImplementedException(); }
        public object SIN(object value)
        {
            // TODO: Tests need to confirm that double precision is used (eg. COS returns different values for 1.111111 and 1.1111111)
            double radians = CDBL_Precise(value, "'Sin'");
            return Math.Sin(radians);
        }
        public object SQR(object value)
        {
            // TODO: Require tests
            // - Always returns double
            // - Accepts double precision input (eg. 1.111111 vs 1.1111111)
            // - Negative values => InvalidProcedureCallOrArgumentException (though zero is, of course, an acceptable input)
            double numericValue = CDBL_Precise(value, "'Sqr'");
            if (numericValue < 0)
                throw new InvalidProcedureCallOrArgumentException($"numericValue must be positive. value:{value}");
            return Math.Sqrt(numericValue);
        }
        public object TAN(object value) { throw new NotImplementedException(); }
        /// <summary>
        /// Returns the number of seconds that have elapsed since midnight
        /// </summary>
        public float TIMER()
        {
            // VBScript returns it as a "Single" (which is equivalent to a .NET float aka Single) and only appears to return up to two decimal place
            return (float)Math.Round((decimal)DateTime.Now.TimeOfDay.TotalSeconds, decimals: 2);
        }
        // - String functions
        public short ASC(object value)
        {
            var valueSafe = VAL(value);
            if (valueSafe == null)
                throw new InvalidProcedureCallOrArgumentException($"value is null. value:{value}");
            if (valueSafe == DBNull.Value)
                throw new InvalidUseOfNullException($"result is null. value:{valueSafe}");

            string s = CSTR(valueSafe);
#pragma warning disable CA1820 // Test for empty strings using string length
            if (s == "")
                throw new InvalidProcedureCallOrArgumentException($"empty text is not supported. s:{s}");
#pragma warning restore CA1820 // Test for empty strings using string length

            char characterValue = s[0];
            return (short)Encoding.Default.GetBytes(new[] { characterValue })[0];
        }
        public object ASCB(object value) { throw new NotImplementedException(); }
        public short ASCW(object value)
        {
            var valueSafe = VAL(value);
            if (valueSafe == null)
                throw new InvalidProcedureCallOrArgumentException($"value is null. value:{value}");
            if (valueSafe == DBNull.Value)
                throw new InvalidUseOfNullException($"valueSafe is null");

            string s = CSTR(valueSafe);
#pragma warning disable CA1820 // Test for empty strings using string length
            if (s == "")
                throw new InvalidProcedureCallOrArgumentException($"empty text is not supported. s:{s}");
#pragma warning restore CA1820 // Test for empty strings using string length

            return (short)s[0];
        }
        public string CHR(object value) // lubo:return the character associated with a specific ASCII (or ANSI) code. test with 125 : vbscript(Windows-1252) is different than (char)155 in .net (Unicode)
        {
            //lubo:Encoding.Default on Windows corresponds to ANSI code page (e.g. Windows-1252 for en-US).
            // Windows-1252 maps '€' (Unicode U+20AC = 8364) to a single byte 0x80. => cannot just cast 8364 to byte! => encode the character to bytes using Encoding!
            try
            {
                //if (value != null)
                //{
                byte b = CBYTECore(value, "'CHR'");
                char result = Windows1252.GetChars([b])[0]; // not Encoding.UTF8 for VBScript
                return new string(result, 1);
                //}
                /*
                if (value != null && value is not IConvertible && value is DispatchWrapper dwY)
                {
                    value = dwY.WrappedObject;
                }

                if (value == null)
                {
                    //char c = (char)0;// Encoding.Default.GetChars(new[] { CBYTECore(value, "'CHR'") })[0];
                    return new string((char)0, 1);
                }
                if (value == DBNull.Value)
                {
                    var c = Encoding.Default.GetChars(new[] { CBYTECore(value, "'CHR'") })[0];
                    return new string(c, 1);
                }
                if (value is string valueString && valueString == string.Empty)
                {
                    return new string((char)0, 1);
                }
                //if (value != null)
                {
                    // Convert the int code point to char
                    byte b;
                    if (value is int valueInt32)
                    {
                        char c = (char)(int)valueInt32; // 8364 -> '€'
                        b = Windows1252.GetBytes(new[] { c })[0];
                    }
                    else
                    {
                        if (value is not IConvertible)
                        {
                            // try get default property's value and convert it to byte
                            var pis = value.GetType().GetProperties();
                            if (pis.Length == 1)
                            {
                                var pi = pis[0];
                                var propertyValue = pi.GetValue(value);
                                if (propertyValue != null && propertyValue is DispatchWrapper dwX)
                                    propertyValue = dwX.WrappedObject;
                                if (propertyValue == null)
                                {
                                    return new string((char)0, 1);
                                }
                                if (propertyValue == DBNull.Value)
                                {
                                    var c = Encoding.Default.GetChars(new[] { CBYTE(propertyValue) })[0];
                                    return new string(c, 1);
                                }
                                b = Convert.ToByte(propertyValue); // double, null, empty
                            }
                            else
                            {
                                b = Convert.ToByte(value); // double, null, empty
                            }
                        }
                        try
                        {
                            b = Convert.ToByte(value); // double, null, empty
                        }
                        catch (System.OverflowException e) // for -0.6
                        {
                            throw new InvalidProcedureCallOrArgumentException("'CHR'", e);
                        }
                    }
                    //byte b = Convert.ToByte(value);
                    char result = Windows1252.GetChars(new[] { b })[0]; // not Encoding.UTF8 for VBSCript
                    return new string(result, 1);
                }
                //else
                //{
                //
                //    // Need to use Encoding.Default.GetChars so that we can reliably get the information back out using ASC (if used something that seems simple like
                //    // "return new string((char)CBYTE(value), 1);" then the correct value won't always be returned from ASC - eg. 155)
                //    var c = Encoding.Default.GetChars(new[] { CBYTE(value) })[0];
                //    return new string(c, 1);
                //}*/
            }
            catch (VBScriptOverflowException e)
            {
                throw new InvalidProcedureCallOrArgumentException("'CHR'", e);
            }
        }
        public object CHRB(object value) { throw new NotImplementedException(); }
        public object CHRW(object value) { throw new NotImplementedException(); }
        public object FILTER(object value) { throw new NotImplementedException(); }
        public object FORMATCURRENCY(object value) { throw new NotImplementedException(); }

        public object FORMATDATETIME(object value)
        {
            return FORMATDATETIMECore(value, VbDateTimeFormat.vbGeneralDate);
        }
        public object FORMATDATETIME(object value, int format)
        {
            return FORMATDATETIMECore(value, Enum.IsDefined(typeof(VbDateTimeFormat), format) ? (VbDateTimeFormat)format : VbDateTimeFormat.vbGeneralDate);
        }

        private enum VbDateTimeFormat
        {
            vbGeneralDate = 0,
            vbLongDate = 1,
            vbShortDate = 2,
            vbLongTime = 3,
            vbShortTime = 4
        }
        private string FORMATDATETIMECore(object value, VbDateTimeFormat format)
        {
            DateTime dt = CDATECore(value, "FORMATDATETIME");
            //d = "2026-01-05 15:42"
            //
            //WScript.Echo FormatDateTime(d)        ' => 0
            //WScript.Echo FormatDateTime(d, 0)     ' 0:vbGeneralDate : General date/time : Short date + long time
            //WScript.Echo FormatDateTime(d, 1)     ' 1:vbLongDate  : Long date
            //WScript.Echo FormatDateTime(d, 2)     ' 2:vbShortDate : Short date
            //WScript.Echo FormatDateTime(d, 3)     ' 3:vbLongTime  : Long time
            //WScript.Echo FormatDateTime(d, 4)     ' 4:vbShortTime :  Short time

            switch (format)
            {
                case VbDateTimeFormat.vbLongDate:
                    return dt.ToString("D", _culture);

                case VbDateTimeFormat.vbShortDate:
                    return dt.ToString("d", _culture);

                case VbDateTimeFormat.vbLongTime:
                    return dt.ToString("T", _culture);

                case VbDateTimeFormat.vbShortTime:
                    return dt.ToString("t", _culture);

                case VbDateTimeFormat.vbGeneralDate:
                default:
                    // VBScript behavior:
                    // - If time present → short date + long time
                    // - If only date → short date
                    // - If only time → long time
                    if (dt.Date == DateTime.MinValue.Date)
                        return dt.ToString("T", _culture); // time only

                    if (dt.TimeOfDay == TimeSpan.Zero)
                        return dt.ToString("d", _culture); // date only

                    return dt.ToString("g", _culture); // short date + short time (closest)
            }
        }
        public object FORMATNUMBER(object expression, object numDigitsAfterDecimal) => FORMATNUMBERCore(expression, numDigitsAfterDecimal);
        public object FORMATNUMBER(object expression) => FORMATNUMBERCore(expression, _culture.NumberFormat.NumberDecimalDigits);
        private string FORMATNUMBERCore(object expression, object numDigitsAfterDecimal) // FormatNumber(Expression, NumDigitsAfterDecimal [, IncludeLeadingDigit [, UseParensForNegativeNumbers [, GroupDigits ]]])
        {
            var expressionV = _valueRetriever.VAL(expression, "'FORMATNUMBER'");
            double expressionNum = GetAsNumber<double>(expressionV, "'FORMATNUMBER'", Convert.ToDouble);

            int decimals = GetAsNumber<int>(numDigitsAfterDecimal, "'FORMATNUMBER'", Convert.ToInt32);

            string result = expressionNum.ToString($"F{decimals}", _culture);
            return result;
        }
        public object FORMATPERCENT(object value) { throw new NotImplementedException(); }
        public object HEX(object? value)
        {
            value = _valueRetriever.VAL(value, "'Hex'");
            if (value == DBNull.Value)
                return DBNull.Value;

            bool useShortFormatForNegativeValues = (value is bool) || (value is short);
            int numericValue = CLNG(value, "'Hex'");
            if (numericValue >= 0)
                return numericValue.ToString("X", CultureInfo.InvariantCulture);

            // For "short" values (ie. VBScript Ints and Booleans), -1 should be returned as FFFF -2 as as FFFE while for other values (Single, Long, Double, etc..)
            // -1 should be returned as FFFFFFFF and -2 as FFFFFFFE
            return ((useShortFormatForNegativeValues ? 0x10000 : 0x100000000) + numericValue).ToString("X", CultureInfo.InvariantCulture);
        }

        public object INSTR(object valueToSearch, object valueToSearchFor) { return INSTR(1, valueToSearch, valueToSearchFor); }
        public object INSTR(object startIndex, object valueToSearch, object valueToSearchFor) { return INSTR(startIndex, valueToSearch, valueToSearchFor, 0); }
        public object INSTR(object? startIndex, object? valueToSearch, object? valueToSearchFor, object? compareMode)
        {
            // Validate input
            startIndex = _valueRetriever.VAL(startIndex, "'InStr'");
            valueToSearch = _valueRetriever.VAL(valueToSearch, "'InStr'");
            valueToSearchFor = _valueRetriever.VAL(valueToSearchFor, "'InStr'");
            compareMode = _valueRetriever.VAL(compareMode, "'InStr'");
            if (startIndex == DBNull.Value)
                throw new InvalidUseOfNullException("startIndex may not be null");
            int startIndexInt = CLNG(startIndex, "'InStr'");
            if (startIndexInt <= 0)
                throw new InvalidProcedureCallOrArgumentException("'INSTR' (startIndex must be a positive integer)");
            if (compareMode == DBNull.Value)
                throw new InvalidUseOfNullException("compareMode may not be null");
            int compareModeInt = CLNG(compareMode, "'InStr'");
            if ((compareModeInt != 0) && (compareModeInt != 1))
                throw new InvalidProcedureCallOrArgumentException("'INSTR' (compareMode may only be 0 or 1)");

            // Deal with null-ish special cases
            if ((valueToSearch == DBNull.Value) || (valueToSearchFor == DBNull.Value))
                return DBNull.Value;
            if (valueToSearch == null)
                return 0;
            if (valueToSearchFor == null)
                return 1;

            // If the startIndex would go past the end of valueToSearch then return zero
            // - Since startIndex is one-based, we need to subtract one from it to perform this test
            string valueToSearchString = _valueRetriever.STR(valueToSearch);
            string valueToSearchForString = _valueRetriever.STR(valueToSearchFor);
            if (valueToSearchForString.Length + (startIndexInt - 1) > valueToSearchString.Length)
                return 0;

            bool useCaseInsensitiveTextComparisonMode = (compareModeInt == 1);
            int zeroBasedMatchIndex = valueToSearchString.IndexOf(
                valueToSearchForString,
                startIndexInt - 1, // This is one-based in VBScript but zero-based in C# (hence the minus one)
                useCaseInsensitiveTextComparisonMode ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal
            );
            return zeroBasedMatchIndex + 1;
        }

        public object INSTRREV(object? valueToSearch, object valueToSearchFor)
        {
            // Unlike INSTR, we have to do some work if no startIndex is specified since the default value should be indicate the last character in
            // valueToSearch, if that can be transformed into a non-blank string (if it can not be transformed into a non-object reference at all then
            // throw an exception, and if it is considered to be the equivalent of blank string then default to a startIndex of one, since it's not
            // valid to have a startIndex of zero)
            valueToSearch = _valueRetriever.VAL(valueToSearch, "'InStrRev'");
            int startIndex;
            if ((valueToSearch == null) || (valueToSearch == DBNull.Value))
                startIndex = 1;
            else
                startIndex = Math.Max(1, _valueRetriever.STR(valueToSearch).Length);
            return INSTRREV(valueToSearch, valueToSearchFor, startIndex);
        }
        public object INSTRREV(object? valueToSearch, object? valueToSearchFor, object? startIndex) { return INSTRREV(valueToSearch, valueToSearchFor, startIndex, 0); }
        public object INSTRREV(object? valueToSearch, object? valueToSearchFor, object? startIndex, object? compareMode)
        {
            // Validate input
            startIndex = _valueRetriever.VAL(startIndex, "'InStrRev'");
            valueToSearch = _valueRetriever.VAL(valueToSearch, "'InStrRev'");
            valueToSearchFor = _valueRetriever.VAL(valueToSearchFor, "'InStrRev'");
            compareMode = _valueRetriever.VAL(compareMode, "'InStrRev'");
            if (startIndex == DBNull.Value)
                throw new InvalidUseOfNullException("startIndex may not be null");
            int startIndexInt = CLNG(startIndex, "'InStrRev'");
            if (startIndexInt <= 0)
                throw new InvalidProcedureCallOrArgumentException("'INSTRREV' (startIndex must be a positive integer)");
            if (compareMode == DBNull.Value)
                throw new InvalidUseOfNullException("compareMode may not be null");
            int compareModeInt = CLNG(compareMode, "'InStrRev'");
            if ((compareModeInt != 0) && (compareModeInt != 1))
                throw new InvalidProcedureCallOrArgumentException("'INSTRREV' (compareMode may only be 0 or 1)");

            // Deal with null-ish special cases
            if ((valueToSearch == DBNull.Value) || (valueToSearchFor == DBNull.Value))
                return DBNull.Value;
            if (valueToSearch == null)
                return 0;
            if (valueToSearchFor == null)
                return 1;

            // For INSTRREV, the startIndex is taken from the start of the string, like INSTR. But, unlike INSTR, the content to consider is the content
            // preceding this point, rather than the content following it. As such, there is different past-the-end-of-the-content logic to consider and
            // different substring matching logic to apply.
            // - If the startIndex goes beyond the end of the valueToSearch then no match is allowed, similarly if the startIndex indicates a point in
            //   the valueToSearch where there is insufficient content to match valueToSearchFor
            string valueToSearchString = _valueRetriever.STR(valueToSearch);
            string valueToSearchForString = _valueRetriever.STR(valueToSearchFor);
            if ((startIndexInt > valueToSearchString.Length) || (valueToSearchForString.Length > startIndexInt))
                return 0;

            // When searching for a match, only consider the allowed substring of valueToSearch
            bool useCaseInsensitiveTextComparisonMode = (compareModeInt == 1);
            int zeroBasedMatchIndex = valueToSearchString.Substring(0, startIndexInt).LastIndexOf(
                valueToSearchForString,
                useCaseInsensitiveTextComparisonMode ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal
            );
            return zeroBasedMatchIndex + 1;
        }

        public object MID(object value, object start)
        {
            string valueString = CSTR(value, "'Mid'");
            return MID(value, start, valueString.Length);
        }
        public object MID(object value, object start, object length)
        {
            // TODO: This is just a thrown-together implementation, it needs proper testing relating to the order in which arguments should be evaluated, what argument values are and
            // aren't valid (is length -1 valid??) but it's just enough to make it work for my particular case that I have right at hand now.
            int lengthAsNumber = CLNG(length, "'Mid'");
            int startAsNumber = CLNG(start, "'Mid'");
            string valueString = CSTR(value, "'Mid'");
            int startIndex = startAsNumber - 1;
            if (startIndex >= valueString.Length)
                return "";
            return valueString.Substring(startIndex, Math.Min(lengthAsNumber, valueString.Length - startIndex));
        }
        public object LEN(object? value)
        {
            value = _valueRetriever.VAL(value, "'Len'");
            if (value == null)
                return 0;
            else if (value == DBNull.Value)
                return DBNull.Value;
            return _valueRetriever.STR(value).Length;
        }
        public object LENB(object? value)
        {
            // For almost all cases, returning twice the string length should work here. VBScript uses UTF-16 but I think that it's possible to construct strings that are an odd number
            // of bytes long if you try hard. As such, this shouldn't be considered a particularly robust implementation but (hopefully) it will be good enough. The places that I have
            // encountered it are where binary data has been stored in a string to pass to an ADO Command as type adBinary - a length is required for the data that is passed and that
            // (so far) seems to work fine when a value of twice the string is used (and it seems to work when the length passed to CreateParamter is larger than the actual data, so
            // it hopefully isn't a problem to over-report the LENB value by one byte for cases where there IS an odd number of bytes).
            value = _valueRetriever.VAL(value, "'LenB'");
            if (value == null)
                return 0;
            else if (value == DBNull.Value)
                return DBNull.Value;
            return _valueRetriever.STR(value).Length * 2;
        }
        public object LEFT(object value, object maxLength)
        {
            // Validate inputs first
            var valueSafe = _valueRetriever.VAL(value, "'Left'");
            var maxLengthSafe = _valueRetriever.VAL(maxLength, "'Left'");
            if (maxLengthSafe == DBNull.Value)
                throw new InvalidUseOfNullException($"maxvalue is null. value:{value}, maxvalue:{maxLength}");
            int maxLengthInt = CLNG(maxLengthSafe, "'Left'");
            if (maxLengthInt < 0)
                throw new InvalidProcedureCallOrArgumentException("'LEFT' (maxLength may not be a negative value)");

            // Deal with special cases
            if (valueSafe == null)
                return "";
            if (valueSafe == DBNull.Value)
                return DBNull.Value;

            string valueString = _valueRetriever.STR(valueSafe);
            maxLengthInt = Math.Min(valueString.Length, maxLengthInt);
            return valueString.Substring(0, maxLengthInt);
        }
        public object LEFTB(object value, object maxLength) { throw new NotImplementedException(); }
        public object RGB(object red, object green, object blue)
        {
            short redComponent = red is short redShort ? redShort : Convert.ToInt16(_valueRetriever.NUM(red), CultureInfo.InvariantCulture);
            short greenComponent = red is short greenShort ? greenShort : Convert.ToInt16(_valueRetriever.NUM(green), CultureInfo.InvariantCulture);
            short blueComponent = blue is short blueShort ? blueShort : Convert.ToInt16(_valueRetriever.NUM(blue), CultureInfo.InvariantCulture);

            // Validate input ranges
            if (redComponent < 0 || redComponent > 255)
            {
                throw new ArgumentOutOfRangeException(nameof(red), "values must be between 0 and 255.");
            }
            if (greenComponent < 0 || greenComponent > 255)
            {
                throw new ArgumentOutOfRangeException(nameof(green), "values must be between 0 and 255.");
            }
            if (blueComponent < 0 || blueComponent > 255)
            {
                throw new ArgumentOutOfRangeException(nameof(blue), "values must be between 0 and 255.");
            }
            // packs them as: low byte = red, middle = green, high = blue
            int result = (redComponent & 0xFF) | ((greenComponent & 0xFF) << 8) | ((blueComponent & 0xFF) << 16);
            return result;
        }
        public object RIGHT(object value, object maxLength)
        {
            // Validate inputs first
            var valueSafe = _valueRetriever.VAL(value, "'Right'");
            var maxLengthSafe = _valueRetriever.VAL(maxLength, "'Right'");
            if (maxLengthSafe == DBNull.Value)
                throw new InvalidUseOfNullException($"maxvalue is null. value:{value}, maxvalue:{maxLength}");
            int maxLengthInt = CLNG(maxLengthSafe, "'Right'");
            if (maxLengthInt < 0)
                throw new InvalidProcedureCallOrArgumentException("'LEFT' (maxLength may not be a negative value)");

            // Deal with special cases
            if (valueSafe == null)
                return "";
            if (valueSafe == DBNull.Value)
                return DBNull.Value;

            string valueString = _valueRetriever.STR(valueSafe);
            maxLengthInt = Math.Min(valueString.Length, maxLengthInt);
            return valueString.Substring(valueString.Length - maxLengthInt);
        }
        public object RIGHTB(object value, object maxLength) { throw new NotImplementedException(); }
        public string REPLACE(object value, object toSearchFor, object toReplaceWith) { return REPLACE(value, toSearchFor, toReplaceWith, 1); }
        public string REPLACE(object value, object toSearchFor, object toReplaceWith, object startIndex) { return REPLACE(value, toSearchFor, toReplaceWith, startIndex, -1); }
        public string REPLACE(object value, object toSearchFor, object toReplaceWith, object startIndex, object maxNumberOfReplacements) { return REPLACE(value, toSearchFor, toReplaceWith, startIndex, maxNumberOfReplacements, 0); }
        public string REPLACE(object value, object toSearchFor, object toReplaceWith, object? startIndex, object? maxNumberOfReplacements, object? compareMode)
        {
            // Input validation / type-enforcing
            compareMode = _valueRetriever.VAL(compareMode, "'Replace'");
            if (compareMode == DBNull.Value)
                throw new InvalidUseOfNullException("'Replace'");
            int compareModeNumber = CLNG(compareMode, "'Replace'");
            if ((compareModeNumber != 0) && (compareModeNumber != 1))
                throw new InvalidProcedureCallOrArgumentException("'Replace'");
            maxNumberOfReplacements = _valueRetriever.VAL(maxNumberOfReplacements, "'Replace'");
            if (maxNumberOfReplacements == DBNull.Value)
                throw new InvalidUseOfNullException("'Replace'");
            int maxNumberOfReplacementsNumber = CLNG(maxNumberOfReplacements);
            if (maxNumberOfReplacementsNumber < -1)
                throw new InvalidProcedureCallOrArgumentException("'Replace'");
            startIndex = _valueRetriever.VAL(startIndex, "'Replace'");
            if (startIndex == DBNull.Value)
                throw new InvalidUseOfNullException("'Replace'");
            int startIndexNumber = CLNG(startIndex);
            if ((startIndexNumber < 1) || (startIndexNumber > MAX_VBSCRIPT_STRING_LENGTH))
                throw new InvalidProcedureCallOrArgumentException("'Replace'");
            string toReplaceWithString = _valueRetriever.STR(toReplaceWith, "'Replace'");
            string toSearchForString = _valueRetriever.STR(toSearchFor, "'Replace'");
            string valueString = _valueRetriever.STR(value, "'Replace'");
#pragma warning disable CA1820 // Test for empty strings using string length
            if ((maxNumberOfReplacementsNumber == 0) || (valueString == "") || (toSearchForString == "") || (startIndexNumber > valueString.Length)) // Note: VBScript's startIndex is one-based while C#'s is zero-based
                return valueString;
#pragma warning restore CA1820 // Test for empty strings using string length

            // Real work (2017-08-10 DWR: This loops has been rewritten to use a string builder to try to reduce the string allocations - inspired by https://stackoverflow.com/a/244933/3813189)
            StringBuilder sb = new StringBuilder();
            if (startIndexNumber > 1)
                sb.Append(valueString.AsSpanX(0, startIndexNumber - 1));
            int indexToStartAt = startIndexNumber - 1;
            StringComparison comparison = (compareModeNumber == 0) ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            while ((maxNumberOfReplacementsNumber == -1) || (maxNumberOfReplacementsNumber > 0))
            {
                int index = valueString.IndexOf(toSearchForString, indexToStartAt, comparison);
                if (index == -1)
                    break;

                sb.Append(valueString.AsSpanX(indexToStartAt, index - indexToStartAt));
                sb.Append(toReplaceWithString);
                index += toSearchForString.Length;

                indexToStartAt = index;
                if (maxNumberOfReplacementsNumber != -1)
                    maxNumberOfReplacementsNumber--;
            }
            sb.Append(valueString.AsSpanX(indexToStartAt));
            return sb.ToString();
        }
        public object SPACE(object value)
        {
            object? numberOfSpaces = _valueRetriever.VAL(value, "'Space'");
            if (numberOfSpaces == DBNull.Value)
                throw new InvalidUseOfNullException("'Space'");
            int numberOfSpacesNumber;
            if (numberOfSpaces == null)
                numberOfSpacesNumber = 0;
            else
            {
                numberOfSpacesNumber = CLNG(numberOfSpaces, "'Space'");
                if (numberOfSpacesNumber < 0)
                    throw new InvalidProcedureCallOrArgumentException("'Space'");
            }

            return new string(' ', numberOfSpacesNumber);
        }
        public object[] SPLIT(object value) { return SPLIT(value, " "); }
        public object[] SPLIT(object? value, object? delimiter)
        {
            // Basic input validation
            delimiter = _valueRetriever.VAL(delimiter, "'Split'");
            if (delimiter == DBNull.Value)
                throw new InvalidUseOfNullException("'Split'");
            value = _valueRetriever.VAL(value, "'Split'");
            if (value == DBNull.Value)
                throw new InvalidUseOfNullException("'Split'");

            // Should be fine to translate both values into strings using the standard mechanism (no exception should arise)
            // Note that Empty and blank string are special cases; always return an empty array, NOT an array with a single element (which would seem more logical)
            // - eg. Split(" ", ",") returns an array with a single element " " while Split("", ",") returns an empty array
            string valueString = _valueRetriever.STR(value, "'Split'");
            string delimiterString = _valueRetriever.STR(delimiter, "'Split'");
            if (string.IsNullOrEmpty(valueString))
                return [];
            return valueString.Split(new[] { delimiterString }, StringSplitOptions.None).Cast<object>().ToArray();
        }
        public object STRCOMP(object string1, object string2) { return STRCOMP(string1, string2, 0); }
        public object STRCOMP(object string1, object string2, object compare) { return ToVBScriptNullable<int>(STRCOMP_Internal(string1, string2, compare)); }
        private int? STRCOMP_Internal(object? string1, object? string2, object compare)
        {
            string? text1 = string1 == null || string1 == DBNull.Value ? null : _valueRetriever.STR(string1);
            string? text2 = string2 == null || string2 == DBNull.Value ? null : _valueRetriever.STR(string2);
            if (text1 == null && text2 == null)
                return null;

            int compareModeCode = compare is bool compareBool ?
                                                  compareBool
                                                    ? 1 // Text compare => ignore case
                                                    : 0 // Binary compare => don't ignore case
                            : (int)compare;
            //_valueRetriever.(compare);
            var comparison = compareModeCode == 1
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;
            return string.Compare(text1, text2, comparison);
        }
        public object STRREVERSE(object value) { throw new NotImplementedException(); }
        public object TRIM(object? value)
        {
            value = _valueRetriever.VAL(value, "'Trim'");
            if (value == null)
                return "";
            else if (value == DBNull.Value)
                return DBNull.Value;
            return _valueRetriever.STR(value).Trim(' ');
        }
        public object LTRIM(object? value)
        {
            value = _valueRetriever.VAL(value, "'LTrim'");
            if (value == null)
                return "";
            else if (value == DBNull.Value)
                return DBNull.Value;
            return _valueRetriever.STR(value).TrimStart(' ');
        }
        public object RTRIM(object? value)
        {
            value = _valueRetriever.VAL(value, "'RTrim'");
            if (value == null)
                return "";
            else if (value == DBNull.Value)
                return DBNull.Value;
            return _valueRetriever.STR(value).TrimEnd(' ');
        }
        public object LCASE(object? value)
        {
            value = _valueRetriever.VAL(value, "'LCase'");
            if (value == null)
                return "";
            else if (value == DBNull.Value)
                return DBNull.Value;
#pragma warning disable CA1308
            return _valueRetriever.STR(value).ToLower(CultureInfo.InvariantCulture);
#pragma warning restore CA1308
        }
        public object UCASE(object? value)
        {
            value = _valueRetriever.VAL(value, "'UCase'");
            if (value == null)
                return "";
            else if (value == DBNull.Value)
                return DBNull.Value;
#pragma warning disable CA1304 // Specify CultureInfo
            return _valueRetriever.STR(value).ToUpper(CultureInfo.InvariantCulture);
#pragma warning restore CA1304 // Specify CultureInfo
        }
        private const string NonEscapedChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789@*_+-./";
        public object ESCAPE(object? value)
        {
            value = _valueRetriever.VAL(value, "'ESCAPE'");
            if (value == null)
                return "";
            else if (value == DBNull.Value)
                return DBNull.Value;

            string valueString = _valueRetriever.STR(value);
#pragma warning disable CA1820 // Test for empty strings using string length
            if (valueString == "")
                return "";
#pragma warning restore CA1820 // Test for empty strings using string length

            StringBuilder sb = new StringBuilder();
            foreach (char c in valueString)
            {
                if (NonEscapedChars.Contains(c, StringComparison.Ordinal))
                {
                    sb.Append(c);
                }
                else if (c <= 0xFF)
                {
                    sb.Append('%');
                    sb.Append(((int)c).ToString("X2", CultureInfo.InvariantCulture));
                }
                else
                {
                    sb.Append("%u");
                    sb.Append(((int)c).ToString("X4", CultureInfo.InvariantCulture));
                }
            }

            return sb.ToString();
        }
        private static int? HexDigitToInt(char digit)
        {
            if (digit >= '0' && digit <= '9')
                return digit - '0';
            if (digit >= 'A' && digit <= 'F')
                return (digit - 'A') + 0xA;
            if (digit >= 'a' && digit <= 'f')
                return (digit - 'a') + 0xA;

            return null;
        }
        public object UNESCAPE(object? value)
        {
            value = _valueRetriever.VAL(value, "'UNESCAPE'");
            if (value == null)
                return "";
            else if (value == DBNull.Value)
                return DBNull.Value;

            string valueString = _valueRetriever.STR(value);
#pragma warning disable CA1820 // Test for empty strings using string length
            if (valueString == "")
                return "";
#pragma warning restore CA1820 // Test for empty strings using string length

            int length = valueString.Length;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < length; i++)
            {
                if (valueString[i] == '%')
                {
                    // Try to parse a %uXXXX sequence
                    if (i + 5 < length && valueString[i + 1] == 'u')
                    {
                        int? digit1 = HexDigitToInt(valueString[i + 2]);
                        int? digit2 = HexDigitToInt(valueString[i + 3]);
                        int? digit3 = HexDigitToInt(valueString[i + 4]);
                        int? digit4 = HexDigitToInt(valueString[i + 5]);

                        if (digit1.HasValue && digit2.HasValue && digit3.HasValue && digit4.HasValue)
                        {
                            sb.Append((char)((digit1 << 12) + (digit2 << 8) + (digit3 << 4) + digit4));
                            i += 5;
                            continue;
                        }
                    }

                    // Try to parse a %XX sequence
                    if (i + 2 < length)
                    {
                        int? digit1 = HexDigitToInt(valueString[i + 1]);
                        int? digit2 = HexDigitToInt(valueString[i + 2]);

                        if (digit1.HasValue && digit2.HasValue)
                        {
                            sb.Append((char)((digit1 << 4) + digit2));
                            i += 2;
                            continue;
                        }
                    }
                }

                // Add the character as-is
                sb.Append(valueString[i]);
            }

            return sb.ToString();
        }
        // - Type comparisons
        public bool ISARRAY(object? value)
        {
            // Use the same approach as for ISEMPTY..
            try
            {
                bool parameterLessDefaultMemberWasAvailable;
                if (!_valueRetriever.TryVAL(value, out parameterLessDefaultMemberWasAvailable, out object? valueVal))
                    return false;
                return (valueVal != null) && valueVal.GetType().IsArray;
            }
#pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception e)
            {
                SETERROR(e);
                return false;
            }
#pragma warning restore CA1031 // Do not catch general exception types
        }
        public bool ISDATE(object value)
        {
            // Use the same basic approach as for ISEMPTY..
            bool swallowAnyError = false;
            try
            {
                bool parameterLessDefaultMemberWasAvailable;
                if (!_valueRetriever.TryVAL(value, out parameterLessDefaultMemberWasAvailable, out object? valueVal))
                    return false;

                // Any error encountered in evaluating the default member (if required to coerce value into a value type) should be recorded with
                // SETERROR, but if the value is not a valid date and an exception is thrown by the DateParser, then that should NOT be recorded
                swallowAnyError = true;
                if (valueVal == null)
                    return false;
                if (valueVal is DateTime)
                    return true;
                DateParser.ForCulture(_culture).Parse(valueVal.ToString(), _culture); // If this doesn't throw an exception then it must be a valid-for-VBScript date string
                return true;
            }
#pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception e)
            {
                if (!swallowAnyError)
                    SETERROR(e);
                return false;
            }
#pragma warning restore CA1031 // Do not catch general exception types
        }
        public bool ISEMPTY(object value)
        {
            try
            {
                // If this can not be coerced into a value type then it can't be Empty, so return false
                bool parameterLessDefaultMemberWasAvailable;
                if (!_valueRetriever.TryVAL(value, out parameterLessDefaultMemberWasAvailable, out object? valueVal))
                {
                    if (value is ScriptDispatchWrapper dw)
                    {
                        return false;//return dw.WrappedObject == null; // value is an object variable containing Nothing, not Empty, BUT lubo checks for null :-))
                    }
                    return false;
                }

                // If it IS a value type, or was manipulated into one, then check for null (aka VBScript's Empty)
                return valueVal == null;
            }
#pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception e)
            {
                // If an exception was raised while evaluating a default member (meaning "value" was not a value but it had a default member that could
                // be investigated.. but an exception was raised within the evaluation of that member) then record the error and return false (as is
                // consistent with VBScript's behaviour)
                SETERROR(e);
                return false;
            }
#pragma warning restore CA1031 // Do not catch general exception types
        }
        public bool ISNULL(object value)
        {
            // Use the same approach as for ISEMPTY..
            try
            {
                bool parameterLessDefaultMemberWasAvailable;
                if (!_valueRetriever.TryVAL(value, out parameterLessDefaultMemberWasAvailable, out object? valueVal))
                    return false;
                return valueVal == DBNull.Value;
            }
#pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception e)
            {
                SETERROR(e);
                return false;
            }
#pragma warning restore CA1031 // Do not catch general exception types
        }
        private static Regex SpaceFollowingMinusSignRemover = new Regex(@"-\s+", RegexOptions.Compiled);
        public bool ISNUMERIC(object value)
        {
            // Use the same basic approach as for ISEMPTY..
            try
            {
                bool parameterLessDefaultMemberWasAvailable;
                if (!_valueRetriever.TryVAL(value, out parameterLessDefaultMemberWasAvailable, out object? valueVal))
                    return false;
                if (valueVal == null)
                    return true; // Empty is identified as numeric in VBScript
                                 // double.TryParse seems to match VBScript's pretty well (see the test cases for more details) with one exception; VBScript will tolerate whitespace between
                                 // a negative sign and the start of the content, so we need to do consider replacements (any "-" followed by whitespace should become just "-")
                double numericValue;
                return double.TryParse(SpaceFollowingMinusSignRemover.Replace(valueVal.ToString(), "-"), out numericValue);
            }
#pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception e)
            {
                SETERROR(e);
                return false;
            }
#pragma warning restore CA1031 // Do not catch general exception types
        }
        public bool ISOBJECT(object value)
        {
            return !_valueRetriever.IsVBScriptValueType(value);
        }
        public string TYPENAME(object? value)
        {
            if (value == null)
                return "Empty";
            if (value == DBNull.Value)
                return "Null";
            if (IsVBScriptNothing(value))
                return "Nothing";

            Type type = value.GetType();
            if (type.IsArray && (type.GetElementType() == typeof(Object)))
                return "Variant()";
            if (_valueRetriever.IsVBScriptValueType(value))
            {
                if (type == typeof(bool))
                    return "Boolean";
                if (type == typeof(byte))
                    return "Byte";
                if (type == typeof(Int16))
                    return "Integer";
                if (type == typeof(Int32))
                    return "Long";
                if (type == typeof(double))
                    return "Double";
                if (type == typeof(DateTime))
                    return "Date";
                if (type == typeof(Decimal))
                    return "Currency";
                return Information.TypeName(value, _culture);
            }

            if (type.IsCOMObject)
            {
                string typeDescriptorClassName = TypeDescriptor.GetClassName(value);
                if (!string.IsNullOrWhiteSpace(typeDescriptorClassName))
                    return typeDescriptorClassName;
            }
            SourceClassNameAttribute? sourceClassName = type.GetCustomAttributes(typeof(SourceClassNameAttribute), inherit: true).OfType<SourceClassNameAttribute>().FirstOrDefault();
            if (sourceClassName != null)
                return sourceClassName.Name;

            // This will always fall through to Object if it finds nothing better along the way
            while (true)
            {
                ComVisibleAttribute? comVisibleAttributeIfAny = type.GetCustomAttributes(typeof(ComVisibleAttribute), inherit: false).OfType<ComVisibleAttribute>().FirstOrDefault();
                if ((comVisibleAttributeIfAny != null) && comVisibleAttributeIfAny.Value)
                    return type.Name;
                type = type.BaseType;
            }
        }
        public object VARTYPE(object value)
        {
            if (value == null)
                return MyVarEnum.VT_EMPTY;
            return (short)IDispatchAccess.GetVariantType(value!);
        }

        // - Array functions
        public object ARRAY(params object[] value)
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value));
            return value;
        }
        public void ERASE(object target, Action<object> targetSetter)
        {
            if (targetSetter == null) throw new ArgumentNullException(nameof(targetSetter));
            // ERASE is more like a keyword in VBScript than a function - none of the builtin VBScript functions take arguments by-ref and nearly all of them apply a lot of
            // similar handling to inputs such as raising invalid-use-of-null errors where VBScript Null is not expected and considering parameter-less default properties
            // and function when expected a value type and receiving an object reference. ERASE does not do that; if the target is not an array then it's a type mismatch,
            // doesn't matter whether it's Empty, Null, Nothing, a number, a string, a date, an object reference with a default parameterless property; it's type mismatch!
            // - Note: A "targetSetter" is required to update the array, rather than just taking the target argument as by-ref, since it would be common for translated
            //   code to call "_.ERASE(ref outer.names)", which would be invalid C# code since "ref" cannot be used with property accessors
            if ((target == null) || !target.GetType().IsArray)
                throw new TypeMismatchException("'Erase'");
            targetSetter(Array.Empty<object>());
        }
        public void ERASE(object target, params object[] arguments)
        {
            // This variation of ERASE is similarly strict to the one above (target must be an array or it's a type mismatch, no matter what!) but the arguments are then
            // evaluated and interpreted as array index values - if this fails then it's a type mismatch as well. The indices must point at an element in the array that
            // is also an array, that is what will get erased. If the argument count does not match the array rank then it's a subscript-out-of-range failure (this
            // includes the case of zero arguments, which is what "ERASE a()" is translated into - it needs to get to this point at runtime so that the type of
            // "a" can be checked, which determines whether the failure is a type-mismatch or subscript-out-of-range).
            Array? targetArray = target as Array;
            if (targetArray == null)
                throw new TypeMismatchException("'Erase'");
            if ((arguments == null) || (arguments.Length == 0))
                throw new SubscriptOutOfRangeException("'Erase'");
            int[] numericArguments = arguments.Select(a => CLNG(a, "'Erase'")).ToArray();
            if (targetArray.Rank != numericArguments.Length)
                throw new SubscriptOutOfRangeException("'Erase'");
            object elementValue;
            try
            {
                elementValue = targetArray.GetValue(numericArguments);
            }
            catch (Exception e)
            {
                throw new SubscriptOutOfRangeException("'Erase'", e);
            }
            if ((elementValue as Array) == null)
            {
                // The element in the target array must also so be an array since that is what's effectively getting erased
                throw new TypeMismatchException("'Erase'");
            }
            targetArray.SetValue(Array.Empty<object>(), numericArguments);
        }
        public string JOIN(object value) { return JOIN(value, " "); }
        public string JOIN(object? value, object? delimiter)
        {
            delimiter = _valueRetriever.VAL(delimiter, "'Join'");
            if (value == DBNull.Value)
                throw new InvalidUseOfNullException("'Join'");
            value = _valueRetriever.VAL(value, "'Join'");
            if (delimiter == DBNull.Value)
                throw new InvalidUseOfNullException("'Join'");
            if (value == null)
                throw new TypeMismatchException("'Join'");
            if (value == DBNull.Value)
                throw new InvalidUseOfNullException("'Join'");
            Type valueType = value.GetType();
            if (!valueType.IsArray)
                throw new TypeMismatchException("'Join'");
            int arrayRank = valueType.GetArrayRank();
            if (arrayRank == 0)
                return "";
            else if (arrayRank > 1)
                throw new TypeMismatchException("'Join'");
            return string.Join(
                (delimiter == null) ? "" : _valueRetriever.STR(delimiter),
                ((Array)value)
                    .Cast<object>()
                    .Select(element =>
                    {
                        var elementVal = _valueRetriever.VAL(element, "'Join'");
                        if (elementVal == DBNull.Value)
                            throw new TypeMismatchException("'Join'");
                        return (elementVal == null) ? "" : _valueRetriever.STR(elementVal);
                    })
            );
        }
        public int LBOUND(object value) { return LBOUND(value, 1); }
        public int LBOUND(object? value, object dimension)
        {
            // If both the value and dimension are invalid values, the dimension errors should be raised first (so try to process that value first)
            int dimensionInt = CLNG(dimension, "'LBound'");
            Array? array = _valueRetriever.VAL(value, "'LBound'") as Array;
            if (array == null)
                throw new TypeMismatchException("'LBound'");
            if ((dimensionInt < 1) || (dimensionInt > array.Rank))
                throw new SubscriptOutOfRangeException("'LBound'");
            return array.GetLowerBound(dimensionInt - 1); // Note: VBScript uses one-based a dimension value here while C# is zero-based, hence the -1
        }
        public int UBOUND(object value) { return UBOUND(value, 1); }
        public int UBOUND(object value, object dimension)
        {
            // If both the value and dimension are invalid values, the dimension errors should be raised first (so try to process that value first)
            int dimensionInt = CLNG(dimension, "'UBound'");
            Array? array = _valueRetriever.VAL(value, "'UBound'") as Array;
            if (array == null)
                throw new TypeMismatchException("'UBound'");
            if ((dimensionInt < 1) || (dimensionInt > array.Rank))
                throw new SubscriptOutOfRangeException("'UBound'");
            return array.GetUpperBound(dimensionInt - 1); // Note: VBScript uses one-based a dimension value here while C# is zero-based, hence the -1
        }
        // - Date functions
        public DateTime NOW() { return DateTime.Now; }
        public DateTime DATE() { return DateTime.Now.Date; }
        public DateTime TIME() { return new DateTime(DateTime.Now.TimeOfDay.Ticks); }
        public object DATEADD(object? interval, object? number, object? value)
        {
            // DateAdd seems to be an usual functions - it ignores fractions in "number" rather than rounding them (so adding 101, 101.5 or 101.9 is the same as adding 101).
            // It's also unusual in that it won't overflow for enormous numeric values, it always falls back to an invalid-procedure-call-or-argument error (if the number
            // would result in an unrepresentable date). On top of this, it doesn't validate all of its arguments before considering any work - DateAdd("x", "y", Null)
            // returns Null, despite the fact that the "interval" and "number" arguments are nonsense; DateAdd("x", "y", Now()) would result in a type-mismatch error.
            value = _valueRetriever.VAL(value, "'DateAdd'");
            if (value == DBNull.Value)
                return DBNull.Value; // Don't even check the other arguments if we got a Null value argument
            DateTime dateValue = CDATECore(value, "'DateAdd'");
            // The MSDN documentation (for VBA, but which is the closest I could find: https://msdn.microsoft.com/en-us/library/aa262710%28v=vs.60%29.aspx) says that "If
            // number isn't a Long value, it is rounded to the nearest whole number before being evaluated." However, testing with VBScript shows this not to be the case.
            // For example, adding (for any interval) 103, 103.1, 103.5 or 103.9 all have the same effect, as do adding 102, 102.1, 102.5 or 102.9, which indicates that
            // the fractional part of the number is being ignored, not rounded. Pushing the limits shows that 1.999999999999999 (15x 9s) will result in 1 being added while
            // 1.9999999999999999 (16x 9s) will result in 2 being added. With 10.9999 it's still 15 vs 16 9s where it changes (from 10 to 11), while with 100.999 it's
            // 14 vs 15 9s. This is consistent with double precision in .net and using CDBL_Precise and then truncating the value will achieve the same effect.
            // - On top of this, if the number lies outside the Int32 range ("Long" in VBScript), then it initially looks like it rolls over.. but actually it just rolls
            //   over and gets stuck at Int32.MinValue; for example any of the following number values will result in the same as if -2147483648 (Int32.MinValue) had been
            //   specified as the number argument: 2147483648 (Int32.MaxValue + 1), 21474836470 (Int32.MaxValue * 10), 1844674407370955161500 (UInt64.MaxValue * 10)
            int intNumber;
            double doubleNumber = Math.Truncate(CDBL_Precise(number, "'DateAdd'"));
            if ((doubleNumber < int.MinValue) || (doubleNumber > int.MaxValue))
                intNumber = int.MinValue;
            else
                intNumber = (int)doubleNumber;
            interval = _valueRetriever.VAL(interval, "'DateAdd'");
            if (interval == DBNull.Value)
                throw new InvalidUseOfNullException("'DateAdd'");
            string? intervalString = interval as string;
            if (intervalString == null)
                throw new InvalidProcedureCallOrArgumentException("'DateAdd'");
            Func<DateTime, int, DateTime> dateManipulator;
#pragma warning disable CA1308
            switch (intervalString.ToLower(CultureInfo.InvariantCulture)) // Interval matching is case-insensitive in VBScript (it won't allow leading or trailing whitespace, though)
#pragma warning restore CA1308
            {
                default:
                    throw new InvalidProcedureCallOrArgumentException("'DateAdd'");
                case "yyyy":
                    dateManipulator = (date, increment) => date.AddYears(increment);
                    break;
                case "q":
                    dateManipulator = (date, increment) => date.AddMonths(increment * 3); // quarter
                    break;
                case "m":
                    dateManipulator = (date, increment) => date.AddMonths(increment);
                    break;
                case "ww":
                    dateManipulator = (date, increment) => date.AddDays(increment * 7); // week
                    break;
                case "y":
                case "d":
                case "w":
                    // Any of "y" (Day of year), "d" (Day) or "w" (weekday) may be used to alter the date, apparently, according to an MSDN article (but this also says that fractional number
                    // values are rounded to the NEAREST whole number, which they aren't, so what does it know.. https://msdn.microsoft.com/en-us/library/aa262710%28v=vs.60%29.aspx). Presumably
                    // these three values are all supported for consistency with related functions such as DATEPART, where the three values will NOT act the same)
                    dateManipulator = (date, increment) => date.AddDays(increment);
                    break;
                case "h":
                    dateManipulator = (date, increment) => date.AddHours(increment);
                    break;
                case "n":
                    dateManipulator = (date, increment) => date.AddMinutes(increment); // This is minutes since "m" is used for months (and don't differentiate between "M" and "m", unlike .net)
                    break;
                case "s":
                    dateManipulator = (date, increment) => date.AddSeconds(increment);
                    break;
            }
            try
            {
                dateValue = dateManipulator(dateValue, intNumber);
            }
            catch (Exception e)
            {
                throw new InvalidProcedureCallOrArgumentException("'DateAdd'", e);
            }
            if ((dateValue < VBScriptConstants.EarliestPossibleDate) || (dateValue.Date > VBScriptConstants.LatestPossibleDate.Date))
                throw new InvalidProcedureCallOrArgumentException("'DateAdd'");
            return dateValue;
        }
        public object DATEDIFF(object interval, object date1, object date2) // TODO: Need to support optional firstDayOrWeek and firstWeekOfYear arguments
        {
            // TODO: Need to confirm that arguments are evaluated in the correct order (if date1 and date2 are invalid, which is reported?)
            // TODO: Document that it returns VBScript "Long" (aka .NET Int32)
            string i = CSTR(interval, "'DateDiff'");
            DateTime d1 = CDATECore(date1, "'DateDiff'");
            DateTime d2 = CDATECore(date2, "'DateDiff'");

            TimeSpan difference = d2.Subtract(d1);
            switch (i)
            {
                default:
                    throw new NotSupportedException($"Unsupported interval: '{interval}'"); // This will be a different exception type once all VBScript-support interval strings are supported
                case "n": // minutes (Ignores seconds => Truncates both dates down to the minute)
                    DateTime dtx1 = new DateTime(d1.Year, d1.Month, d1.Day, d1.Hour, d1.Minute, 0); // Truncate to minute precision
                    DateTime dtx2 = new DateTime(d2.Year, d2.Month, d2.Day, d2.Hour, d2.Minute, 0); // Truncate to minute precision
                    return (int)(dtx2 - dtx1).TotalMinutes;
                case "d":
                    /*
    VBScript: DateDiff("d", "2024-01-01 12:00", "2024-01-02 11:00") == 0
    Your current code → 1 (because of Ceiling)
                    BUT!!!
                    VBScript DateDiff("d", ...) counts day boundaries crossed, not elapsed 24-hour periods. => "d" should always be based on .Date, not TotalDays. (but wrong in other cases)
                     */
                    //return (int)(d2 - d1).TotalDays; //return (int)Math.Ceiling(difference.TotalDays);
                    return (int)Math.Ceiling(difference.TotalDays);
                case "m": // months
                    int yearDifference = d2.Year - d1.Year;
                    int monthDifference = d2.Month - d1.Month;
                    return (yearDifference * 12) + monthDifference;
                case "s":
                    return (int)(d2 - d1).TotalSeconds;
            }
        }
        public object DATEPART(object interval, object valueDate) { throw new NotImplementedException(); }
        public object DATEPART(object interval, object valueDate, object firstDayOfWeek) { throw new NotImplementedException(); } // , object firstweekofyear
        public object DATESERIAL(object year, object month, object day)
        {
            // TODO: This is not a complete implementation, it's just enough to get moving

            // TODO: Implement (and write tests) for this more thoroughly - eg. (99,2,10) => 1999-2-10, (99,14,10) => 100-2-10, (2017,13,1) => 2018-1-1

            int numericYear = CLNG(year);
            int numericMonth = CLNG(month);
            int numericDate = CLNG(day);

            if ((numericMonth < 0) || (numericMonth > 12))
            {
                int numberOfYearsToAdd = (int)Math.Floor((double)numericMonth / 12);
                numericYear += numberOfYearsToAdd;
                numericMonth = numericMonth % 12;
                if (numericMonth < 0)
                    numericMonth += 12; // For negative values (eg. -1 % 12 is -1 so need to add 12 to get to 11, -13 % 12 is also -1 so never need to add more or less than 12)
            }

            // TODO: Check days <= 0 or days > days-in-month/year

            // TODO: Check small year values
            // TODO: What about negative year values??

            return new DateTime(numericYear, numericMonth, numericDate);
        }
        public DateTime DATEVALUE(object? value)
        {
            // In summary, this will do a subset of the processing of CDATE (it will accept a DateTime or a parse-able string, but not a numeric value such as 123.45) and return only the date:
            //   "The reasons for using DateValue and TimeValue to convert a string instead of CDate may not be immediately obvious. Consider the example above. CDate is creating a Date value for the entire supplied
            //    string.  DateValue and TimeValue will allow you to create Date values containing only the specified portion of the string while ignoring the rest."
            // - http://www.aspfree.com/c/a/windows-scripting/working-with-dates-and-times-in-vbscript/
            value = _valueRetriever.VAL(value, "'DateValue'");
            if (value == null)
                throw new TypeMismatchException("'DateValue'");
            if (value == DBNull.Value)
                throw new InvalidUseOfNullException("'DateValue'");
            DateTime dateValue;
            if (value is DateTime)
                dateValue = (DateTime)value;
            else
            {
                try
                {
                    dateValue = DateParser.ForCulture(_culture).Parse(value.ToString(), _culture);
                }
                catch (Exception e)
                {
                    throw new TypeMismatchException("'DateValue'", e);
                }
            }
            return dateValue.Date;
        }
        public DateTime TIMESERIAL(object hours, object minutes, object seconds)
        {
            short secondsAsNumber = CINT(seconds, "'TimeSerial'");
            short minutesAsNumber = CINT(minutes, "'TimeSerial'");
            short hoursAsNumber = CINT(hours, "'TimeSerial'");

            minutesAsNumber += GetQuantityAtNextLargestUnit(ref secondsAsNumber, 60);
            hoursAsNumber += GetQuantityAtNextLargestUnit(ref minutesAsNumber, 60);
            short days = GetQuantityAtNextLargestUnit(ref hoursAsNumber, 24);

            // I have no idea what the original VBScript library authors must have been thinking when they wrote their code, I've just tried to work out an algorithm
            // that matches their results. The first oddity is that (2, 0, 0) and (-2, 0, 0) both return the same value, as if the "-" from -2 is ignored. However,
            // (2, 1, 0) and (-2, 1, 0) do NOT return the same, so the "-" clearly isn't ignore; the first returns is interpreted as "02:01:00" while the second as
            // "01:59:00" as if it flipped the signs and decided to treat it as (2, -1, 0). This sign-flipping-for-negative-values seems to work for every set of
            // values that I've put at it (including large positive and negative minutes and seconds values, such as +/-8000). Note that it seems to be the first
            // non-zero term that triggers the sign-flipping if it is negative, not just the hours values; for example (0, 13, 20) and (0, -13, -20) both return
            // the same result (1899-12-30 00:13:20). To make it even crazier, it is the first non-zero term AFTER values have been shifted around in order to
            // ensure that the seconds and minutes values are smaller than sixty - for example, (2, 0, -8000) is adjusted by realising that -8000s is the same
            // as -2h, -13m and -20s and so the hours value is cancelled out (2 - 2), leaving the three time values as (0, -13, -20) and so the final result
            // from VBScript is 1899-12-30 00:13:20.
            short[] nonZeroTermsInDescendingMagnitude = new[] { days, hoursAsNumber, minutesAsNumber, secondsAsNumber }.Where(value => value != 0).ToArray();
            int multiplier = (nonZeroTermsInDescendingMagnitude.Length != 0 && (nonZeroTermsInDescendingMagnitude.First() < 0)) ? -1 : 1;
            return VBScriptConstants.ZeroDate
                .AddDays(days)
                .AddHours(hoursAsNumber * multiplier)
                .AddMinutes(minutesAsNumber * multiplier)
                .AddSeconds(secondsAsNumber * multiplier);
        }
        private static short GetQuantityAtNextLargestUnit(ref short value, short numberInNextUnit)
        {
            short valueOfNextUnit = (value > 0) ? (short)Math.Floor(value / (double)numberInNextUnit) : (short)Math.Ceiling(value / (double)numberInNextUnit);
            value = (short)(value - (valueOfNextUnit * numberInNextUnit));
            return valueOfNextUnit;
        }
        public DateTime TIMEVALUE(object? value)
        {
            // In summary, this will do a subset of the processing of CDATE (it will accept a DateTime or a parse-able string, but not a numeric value such as 123.45) and return only the time component:
            //   "The reasons for using DateValue and TimeValue to convert a string instead of CDate may not be immediately obvious. Consider the example above. CDate is creating a Date value for the entire supplied
            //    string.  DateValue and TimeValue will allow you to create Date values containing only the specified portion of the string while ignoring the rest."
            // - http://www.aspfree.com/c/a/windows-scripting/working-with-dates-and-times-in-vbscript/
            value = _valueRetriever.VAL(value, "'TimeValue'");
            if (value == null)
                throw new TypeMismatchException("'TimeValue'");
            if (value == DBNull.Value)
                throw new InvalidUseOfNullException("'TimeValue'");
            DateTime dateValue;
            if (value is DateTime)
                dateValue = (DateTime)value;
            else
            {
                try
                {
                    dateValue = DateParser.ForCulture(_culture).Parse(value.ToString(), _culture);
                }
                catch (Exception e)
                {
                    throw new TypeMismatchException("'TimeValue'", e);
                }
            }
            // VBScript represents times by taking its "zero date" and adding hours / minutes / seconds to it
            return VBScriptConstants.ZeroDate.Add(dateValue.TimeOfDay);
        }
        public object DAY(object? value)
        {
            value = _valueRetriever.VAL(value, "'Day'");
            if (value == DBNull.Value)
                return DBNull.Value; // This is special case is the only real difference between the logic here and in CDATE
            return ToClosestSecond(CDATECore(value, "'Day'")).Day;
        }
        public object MONTH(object? value)
        {
            value = _valueRetriever.VAL(value, "'Month'");
            if (value == DBNull.Value)
                return DBNull.Value; // This is special case is the only real difference between the logic here and in CDATE
            return ToClosestSecond(CDATECore(value, "'Month'")).Month;
        }
        public object MONTHNAME(object value) { throw new NotImplementedException(); }
        public object YEAR(object? value)
        {
            value = _valueRetriever.VAL(value, "'Year'");
            if (value == DBNull.Value)
                return DBNull.Value; // This is special case is the only real difference between the logic here and in CDATE
            return ToClosestSecond(CDATECore(value, "'Year'")).Year;
        }
        public object WEEKDAY(object value)
        {
            return WEEKDAY(value, VBScriptConstants.vbSunday);
        }
        public object WEEKDAY(object? value, object firstDayOfWeek)
        {
            value = _valueRetriever.VAL(value, "'Weekday'");
            if (value == DBNull.Value)
                return DBNull.Value; // This is special case is the only real difference between the logic here and in CDATE
            DateTime date = ToClosestSecond(CDATECore(value, "'Weekday'"));

            // NOTE: VBScript weekdays go from Sunday (1) to Saturday (7) (unless overriden by firstDayOfWeek), while .NET DayOfWeek goes from Sunday (0) to Saturday (6)
            int vbsFirstDayOfWeek = CLNG(firstDayOfWeek, "'Weekday'");
            if (vbsFirstDayOfWeek < 0 || vbsFirstDayOfWeek > 7)
                throw new InvalidProcedureCallOrArgumentException("'Weekday'");
            if (vbsFirstDayOfWeek == VBScriptConstants.vbUseSystemDayOfWeek)
                vbsFirstDayOfWeek = (int)_culture.DateTimeFormat.FirstDayOfWeek + 1;

            return (((int)date.DayOfWeek + (8 - vbsFirstDayOfWeek)) % 7) + 1;
        }
        public object WEEKDAYNAME(object value) { return WEEKDAYNAME(value, abbreviate: false); }
        public object WEEKDAYNAME(object value, object abbreviate) { return WEEKDAYNAME(value, abbreviate, firstDayOfWeek: VBScriptConstants.vbSunday); }
        public object WEEKDAYNAME(object value, object abbreviate, object firstDayOfWeek)
        {
            int numericValue = CLNG(value, "'WeekdayName'");
            if ((numericValue < 1) || (numericValue > 7))
                throw new InvalidProcedureCallOrArgumentException($"numericValue not in supported range. numericValue:{numericValue}");

            bool booleanAbbreviate = CBOOL(abbreviate, "'WeekdayName'"); // TODO: Ensure that this behaviour is correct (including errors) and ensure evaluate arguments in correct order

            int numericFirstDayOfWeek = CLNG(firstDayOfWeek, "'WeekdayName'");
            if (numericFirstDayOfWeek != VBScriptConstants.vbSunday)
                throw new NotSupportedException(); // TODO: Deal with firstDayOfWeek properly (and ensure evaluate arguments in correct order)

            // The first day in January 2017 is Sunday and VBScript treats day 1 as Sunday, so we can just take our 1-7 range and use that as the day number in Jan 2017
            // (then we take the name of the day for the generated date and we're all done)
            return new DateTime(2017, 1, numericValue).ToString(booleanAbbreviate ? "ddd" : "dddd", CultureInfo.InvariantCulture);
        }
        public object HOUR(object? value)
        {
            value = _valueRetriever.VAL(value, "'Hour'");
            if (value == DBNull.Value)
                return DBNull.Value; // This is special case is the only real difference between the logic here and in CDATE
            return ToClosestSecond(CDATECore(value, "'Hour'")).Hour;
        }
        public object MINUTE(object? value)
        {
            value = _valueRetriever.VAL(value, "'Minute'");
            if (value == DBNull.Value)
                return DBNull.Value; // This is special case is the only real difference between the logic here and in CDATE
            return ToClosestSecond(CDATECore(value, "'Minute'")).Minute;
        }
        public object SECOND(object? value)
        {
            value = _valueRetriever.VAL(value, "'Second'");
            if (value == DBNull.Value)
                return DBNull.Value; // This is special case is the only real difference between the logic here and in CDATE
            return ToClosestSecond(CDATECore(value, "'Second'")).Second;
        }
        private static DateTime ToClosestSecond(DateTime value)
        {
            DateTime approximateValue = new DateTime(value.Year, value.Month, value.Day, value.Hour, value.Minute, value.Second);
            if (value.Millisecond >= 500) // TODO: Check whether this rounding is correct, should it be banker's rounding?
            {
                if ((DateTime.MaxValue - approximateValue).TotalSeconds > 1.0)
                {
                    approximateValue = approximateValue.AddSeconds(1);
                }
            }
            return approximateValue;
        }
        // - Object creation
        public object CREATEOBJECT(object value)
        {
            string classProgId = _valueRetriever.STR(value);
            if (string.IsNullOrEmpty(classProgId))
                throw new InvalidOperationException("object id:" + value);
            return CREATEOBJECTCore(classProgId, optionalMonikerValues: null);
        }

        private object CREATEOBJECTCore(string classProgId, string? optionalMonikerValues)
        {
            // Creates a new instance of the specified COM object
            // => Set obj = CreateObject("Excel.Application") → always starts a new Excel process
            // => for files: Cannot open persisted objects from file
            // => Works as long as the ProgID is valid
            // Automating applications by starting fresh

            if (string.IsNullOrEmpty(classProgId))
                throw new ArgumentException("object prog-id cannot be null or empty.", nameof(classProgId));
            if (_objectCreateFactories.TryGetValue(classProgId, out Func<string?, object> objectFactory))
                return HandlePostInitializationHandler(classProgId, objectFactory(optionalMonikerValues));

            try
            {
                IHostObjectFactoryHostService? objectFactories = _runtimeHost.TryGetRuntimeHostService<IHostObjectFactoryHostService>();
                if (objectFactories != null)
                {
                    Func<IRuntimeHost, object>? creator = objectFactories.TryGetObjectFactoryRegistration(classProgId);
                    if (creator != null)
                    {
                        return (IReflect)creator(_runtimeHost);
                    }
                }

                //if (classProgId.Length > 0)
                throw new NotSupportedException($"classProgId:{classProgId}");
                //Type comType = Type.GetTypeFromProgID(classProgId, true);
                //
                //return HandlePostInitializationHandler(classProgId, CreateComObject(classProgId, comType));
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create com object for '{classProgId}'", ex);
            }
        }
        public static IReflect TestCreateComObjectTest(string classProgId, Type comType)
        {
            MyComProxy prx = MyComProxy.CreateComProxy(classProgId, comType);
            return prx._comInstance is IReflect rfl ? rfl : prx;
        }

        //"Scripting.Dictionary"
        //throw new InvalidOperationException($"object factory for '{classProgId}' not registered.");
        private static object CreateComObject(string classProgId, Type comType)
        {
            MyComProxy prx = MyComProxy.CreateComProxy(classProgId, comType);
            return prx._comInstance;
        }
        public object GETOBJECT(object path, object value)
        {
            return GETOBJECTCore(path, value);
        }
        public object GETOBJECT(object value)
        {
            return GETOBJECTCore(null, value);
        }
        private object GETOBJECTCore(object? unused, object value)
        {
            // Retrieves an existing instance of a running object, or loads from a file
            // => Set obj = GetObject(, "Excel.Application") → attaches to an already running Excel
            // => Can open objects stored in files (e.g., GetObject("C:\MyDoc.doc"))
            // => Fails if no instance exists and no file is specified (Run-time error 429)
            // Controlling or reusing an already running application instance
            // Designed to bind to system services via monikers not ProgId (progid is for CreateObject)
            // Samples:
            //   "winmgmts:"
            //   "winMgmts::Win32_scheduledJob"
            //   "winMgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\Root\CIMv2:Win32_ScheduledJob""
            //   "LDAP:"
            //   "IIS:"

            /*
            A moniker is a COM naming string that tells Windows how to bind to an object or service.
            It:
            - Often ends with a colon :
            - Can include a path or server info

            Examples of Monikers
              GetObject("winmgmts:")
              GetObject("LDAP://CN=Users,DC=domain,DC=com")
              GetObject("IIS://localhost/W3SVC")

            File moniker:
                Set obj = GetObject("C:\Windows\System32\notepad.exe")
                Set obj = GetObject(".\MyComObject.dll")
             */
            string valueText = _valueRetriever.STR(value, $"'{nameof(GETOBJECT)}'");
            string[] tokens = valueText.Split([':'], StringSplitOptions.RemoveEmptyEntries);

            if (tokens[0].Length == 1) // C:\\
                throw new InvalidOperationException("Local file moniker not supported:" + valueText);

            if (valueText.StartsWith('\\', StringComparison.Ordinal)) // \\192.01.01.01\sharedfiles\
                throw new InvalidOperationException("Shared file moniker not supported:" + valueText);

            string? progid = TryResolveMonikerName(tokens[0]);
            if (progid == null)
            {
                return CREATEOBJECTCore(valueText, null);
                //throw new InvalidOperationException($"Unsupported moniker name: '{tokens[0]}'");
            }

            string optionalMonikerValues = valueText.Substring(tokens.Length + 1);
            IMyMoniker moniker = (IMyMoniker)CREATEOBJECTCore(progid, optionalMonikerValues);
            Guid riid = Guid.Empty;
            object bindedInstance = moniker.BindToObject(null, null, ref riid);
            return bindedInstance;

            //if (tokens.Length == 1)
            //{
            //}
            //else
            //{
            //    //IMyMoniker svc = (IMyMoniker)unbindedInstance;
            //    //return unbindedInstance;
            //    throw new NotImplementedException($"{valueText} > {optionalMonikerValues}");
            //}
        }

        private static FrozenDictionary<string, string> _monikerToProgIdMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            { "winmgmts", "WbemScripting.SWbemLocator"}
        }.ToFrozenDictionary();
        private static string? TryResolveMonikerName(string monikerName)
        {
            return _monikerToProgIdMap.TryGetValue(monikerName, out string? progid) ? progid : null;
        }


        public object EVAL(object value) { throw new NotImplementedException("Dynamic script evaluation. Code:" + value); }
        public object EXECUTE(object value) { throw new NotImplementedException("Dynamic script execution. Code:" + value); } // "script in script"
        public object EXECUTEGLOBAL(object value) { throw new NotImplementedException(); }
        // - Misc
        public object GETLOCALE(object value) { throw new NotImplementedException(); }
        public object GETREF(object value) { throw new NotImplementedException(); }
        public object INPUTBOX(object prompt, object? title = null, object? defaultValue = null)
        {
            // If the user clicks OK or presses ENTER, the InputBox function returns whatever is in the text box.
            // If the user clicks Cancel, it returns a zero-length string ("").
            // he maximum length of prompt is approximately 1024 characters.
            string promptText = _valueRetriever.TryRetrieveStringOrEmpty(prompt);
            string? titleText = _valueRetriever.TryRetrieveStringOrEmpty(title);
            string? defaultText = _valueRetriever.TryRetrieveStringOrEmpty(title);

            IHostInputBoxHostService svc = _runtimeHost.TryGetRuntimeHostService<IHostInputBoxHostService>() ?? throw new InvalidOperationException($"Host service '{nameof(IHostMessageBoxHostService)}' not registered.");
            string result = svc.ShowInputBox(promptText, titleText, defaultText);
            return result;
        }
        public object LOADPICTURE(object value) { throw new NotImplementedException(); }
        public object MSGBOX(object value)
        {
            string prompt = _valueRetriever.STR(value);
            return MSGBOXCore(prompt, null, null);
        }
        public object MSGBOX(object value, object buttons)
        {
            string prompt = _valueRetriever.STR(value);
            short buttonsNum = Convert.ToInt16(_valueRetriever.NUM(buttons), CultureInfo.InvariantCulture);
            return MSGBOXCore(prompt, buttonsNum, null);
        }
        public object MSGBOX(object value, object buttons, object title)
        {
            string prompt = _valueRetriever.STR(value);
            short buttonsNum = Convert.ToInt16(_valueRetriever.NUM(buttons), CultureInfo.InvariantCulture);
            string titleString = _valueRetriever.STR(title);
            return MSGBOXCore(prompt, buttonsNum, titleString);
        }
        private object MSGBOXCore(string prompt, short? buttons = null, string? title = null)
        {
            IHostMessageBoxHostService svc = _runtimeHost.TryGetRuntimeHostService<IHostMessageBoxHostService>() ?? throw new InvalidOperationException($"Host service '{nameof(IHostMessageBoxHostService)}' not registered.");
            MessageBoxResult result = svc.ShowMessageBox(prompt, (MessageBoxButtons)buttons.GetValueOrDefault(0), title ?? "Application");
            return (int)result;
        }
        public string SCRIPTENGINE(object value) { throw new NotImplementedException(); }
        public int SCRIPTENGINEBUILDVERSION(object value) { throw new NotImplementedException(); }
        public int SCRIPTENGINEMAJORVERSION(object value) { throw new NotImplementedException(); }
        public int SCRIPTENGINEMINORVERSION(object value) { throw new NotImplementedException(); }
        public object SETLOCALE(object value) { throw new NotImplementedException(); }

        /// <summary>
        /// This returns the value without any immediate processing, but may keep a reference to it and dispose of it (where applicable) after
        /// the request completes (to try to avoid resources from not being cleaned up in the absence of the VBScript deterministic garbage
        /// collection - classes with a Class_Terminate function are translated into IDisposable types and, while IDisposable.Dispose will not
        /// be called by the translated code, it may be called after the request ends if the requests are tracked here. This will throw an
        /// exception for a null value.
        /// </summary>
        public object NEW(object value)
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value));

            IDisposable? disposableResource = value as IDisposable;
            if (disposableResource != null)
                _disposableReferencesToClearAfterTheRequest.Add(disposableResource);
            return value;
        }

        // Array definitions
        public object NEWARRAY(IEnumerable<object> dimensions)
        {
            if (dimensions == null)
                throw new ArgumentNullException(nameof(dimensions));

            // Note that VBScript specifies upper bounds for arrays, rather than the size - so ReDim a(2) means that the array "a" needs three
            // elements (0, 1 and 2) and so must be declared in C# as object[3]. In VBScript, if negative ranges are specified below -1 (since
            // -1 means zero in C#, which is not an unreasonable request - eg. object[0]) then an out-of-memory error is raised. It shouldn't
            // be possible for this to be called without any dimensions from translated code since that would be a syntax error (and so may
            // be an ArgumentException rather than a specialise VBScript exception).
            int[] dimensionSizes = dimensions.Select(d => CLNG(d, "'NewArray'") + 1).ToArray();
            if (dimensionSizes.Length == 0)
                throw new ArgumentException("No dimensions specified for NEWARRAY");
            if (dimensionSizes.Any(d => d < 0))
                throw new InvalidOperationException("Invalid negative dimensions used for NEWARRAY call");
            return Array.CreateInstance(typeof(object), dimensionSizes);
        }

        public object RESIZEARRAY(object array, IEnumerable<object> dimensions)
        {
            // Note: Don't even check "array" for null until the dimensions have been evaluated
            if (dimensions == null)
                throw new ArgumentNullException(nameof(dimensions));

            // Note that VBScript specifies upper bounds for arrays, rather than the size - so ReDim a(2) means that the array "a" needs three
            // elements (0, 1 and 2) and so must be declared in C# as object[3]. In VBScript, if negative ranges are specified below -1 (since
            // -1 means zero in C#, which is not an unreasonable request - eg. object[0]) then an out-of-memory error is raised. It shouldn't
            // be possible for this to be called without any dimensions from translated code since that would be a syntax error (and so may
            // be an ArgumentException rather than a specialise VBScript exception).
            // - The dimensions are evaulated before the target array is validated (before it is even checked for null, even) in order to
            //   be consistent with VBScript's runtime behaviour
            int[] dimensionSizes = dimensions.Select(d => CLNG(d, "'ResizeArray'") + 1).ToArray();
            if (dimensionSizes.Length == 0)
                throw new ArgumentException("No dimensions specified for RESIZEARRAY");
            Array? arrayTyped = array as Array;
            if (arrayTyped == null)
                throw new TypeMismatchException("'ResizeArray' target not an array");
            if (dimensionSizes.Length != arrayTyped.Rank)
                throw new SubscriptOutOfRangeException("Inconsistent number of dimensions specified for RESIZEARRAY");
            if (dimensionSizes.Any(d => d < 0))
                throw new InvalidOperationException("Invalid negative dimensions used for RESIZEARRAY call");

            for (int dimension = 0; dimension < arrayTyped.Rank - 1; dimension++)
            {
                if (arrayTyped.GetLength(dimension) != dimensionSizes[dimension])
                    throw new SubscriptOutOfRangeException("Invalid dimensions specified for RESIZEARRAY - only the last dimension may vary in size");
            }

            if (dimensionSizes.Length == 1)
            {
                // Copying a 1D array is easy..
                object[] newArray = new object[dimensionSizes[0]];
                Array.Copy(arrayTyped, newArray, Math.Min(arrayTyped.Length, dimensionSizes[0]));
                return newArray;
            }
            else if (dimensionSizes.Length == 2)
            {
                // Copying a 2D array can be done column-by-column, so there's only one loop and an Array.Copy per iteration..
#pragma warning disable CA1814 // Prefer jagged arrays over multidimensional
                object[,] newArray = new object[dimensionSizes[0], dimensionSizes[1]];
#pragma warning restore CA1814 // Prefer jagged arrays over multidimensional
                int numberOfElementsToCopyEachTime = Math.Min(arrayTyped.GetLength(1), dimensionSizes[1]);
                if (numberOfElementsToCopyEachTime > 0)
                {
                    for (int i = 0; i < dimensionSizes[0]; i++)
                    {
                        Array.Copy(
                            arrayTyped,
                            i * arrayTyped.GetLength(1),
                            newArray,
                            i * dimensionSizes[1],
                            numberOfElementsToCopyEachTime
                        );
                    }
                }
                return newArray;
            }
            else
            {
                // Copying an array with more dimensions is more awkward.. the only way I can think of is to go through every element of the
                // new array and copy each value from the old array, so long as the element exists in the old array. This is MUCH less
                // efficient than the process for the 1D or 2D arrays.
                Array newArray = Array.CreateInstance(typeof(object), dimensionSizes);
                int totalNumberOfElements = dimensionSizes.Aggregate(1, (acc, value) => acc * value);
                int[] indicesOfElementToCopy = new int[dimensionSizes.Length];
                for (int i = 0; i < totalNumberOfElements; i++)
                {
                    if (i > 0)
                    {
                        int indexToIncrementNext = 0;
                        while (true)
                        {
                            if (indicesOfElementToCopy[indexToIncrementNext] < (dimensionSizes[indexToIncrementNext] - 1))
                            {
                                indicesOfElementToCopy[indexToIncrementNext]++;
                                break;
                            }
                            indicesOfElementToCopy[indexToIncrementNext] = 0;
                            indexToIncrementNext++;
                        }
                    }
                    bool elementDoesNotExistInSource = false;
                    for (int j = 0; j < indicesOfElementToCopy.Length; j++)
                    {
                        if (arrayTyped.GetLength(j) <= indicesOfElementToCopy[j])
                        {
                            elementDoesNotExistInSource = true;
                            break;
                        }
                    }
                    if (!elementDoesNotExistInSource)
                        newArray.SetValue(arrayTyped.GetValue(indicesOfElementToCopy), indicesOfElementToCopy);
                }
                return newArray;
            }
        }

        public object NEWREGEXP()
        {
            // TODO: Ideally, the object returned here would be a managed implementation of "VBScript.RegExp" (which has a fairly simple interface), to reduce the
            // number of dependencies. But this works and so will do for the time being.
            return CREATEOBJECTCore("VBScript.RegExp", optionalMonikerValues: null);
        }

        /// <summary>
        /// This will never be null (if there is no error then an ErrorDetails with Number zero will be returned)
        /// </summary>
        public ErrorDetails ERR
        {
            get
            {
                Exception? currentError = _trappedErrorIfAny;
                if (currentError == null)
                    return ErrorDetails.NoError;
                SpecificVBScriptException? currentErrorAsVBScriptSpecificError = currentError as SpecificVBScriptException;
                return new ErrorDetails(
                    number: (currentErrorAsVBScriptSpecificError != null) ? currentErrorAsVBScriptSpecificError.ErrorNumber : currentError.HResult, // TODO: Is HResult appropriate?
                    source: currentError.Source,
                    text: currentError.Message,
                    description: "",
                    originalExceptionIfKnown: currentError
                );
            }
        }

        /// <summary>
        /// There are some occassions when the translated code needs to throw a runtime exception based on the content of the source code - eg.
        ///   WScript.Echo 1()
        /// It is clear that "1" is a numeric constant and not a function, and so may not be called as one. However, this is not invalid VBScript and so is
        /// not a compile time error, it is something that must result in an exception at runtime. In these cases, where it is known at the time of translation
        /// that an exception must be thrown, this method may be used to do so at runtime. This is different to SETERROR, since that records an exception that
        /// has already been thrown - this throws the specified exception (it returns an object, rather than void, for the same reason as the below signatures).
        /// </summary>
        public object RAISEERROR(Exception e)
        {
            if (e == null)
                throw new ArgumentNullException(nameof(e));

            throw e;
        }

        // These method signatures have to return a value since these are what are called when the source code includes "Err.Raise 123", which VBScript allows
        // to exist in the form "If (Err.Raise(123)) Then" - if these didn't return values then there could be compile errors in the generated C# that were
        // valid VBScript.
        public object RAISEERROR(object number) { return RAISEERROR(number, ""); }
        public object RAISEERROR(object number, object source) { return RAISEERROR(number, source, ""); }
        public object RAISEERROR(object number, object source, object description)
        {
            // This is another function (like ERASE) that doesn't give many clues - almost every failure is a "Type mismatch" (Null values do not result in
            // "Invalid use of null" and Nothing does not result in "Object variable not set"). However, if "number" is zero then the other two arguments
            // are not evaluated - this only happens if the value for number is ok. And if number is zero then it DOES get a different error :S
            int numericNumber;
            try
            {
                numericNumber = CLNG(number);
            }
            catch (Exception e)
            {
                throw new TypeMismatchException("Err.Raise", e);
            }
            if (numericNumber == 0)
                throw new InvalidProcedureCallOrArgumentException("Err.Raise");
            string sourceString, descriptionString;
            try
            {
                sourceString = _valueRetriever.STR(source);
                descriptionString = _valueRetriever.STR(description);
            }
            catch (Exception e)
            {
                throw new TypeMismatchException("Err.Raise", e);
            }
            throw new CustomException(numericNumber, sourceString, descriptionString);
        }

        public void SETERROR(Exception e)
        {
            // Note that there is (at most) only a single error associated with an executing request. If the error-trapping is enabled and a function F1()
            // executes code that raises an error but then goes and calls F2() which also raises an error, the error recorded from the code in F1 that
            // occured before calling F2 is lost, it is overwritten by F2. So there is no need to try to map trapped errors onto error tokens, there is
            // only one per request (or zero - if there has been no error trapped, or if there HAS been an error trapped that has then been cleared).
            SetErrorCore(e);
        }
        private void SetErrorCore(Exception e)
        {
            if (e == null)
                throw new ArgumentNullException(nameof(e));
            _runtimeLogger.LogException(e);
            _trappedErrorIfAny = e;
        }

        public void CLEARANYERROR()
        {
            // This should be called by translated code that originates from an ON ERROR GOTO 0 with no corresponding ON ERROR RESUME NEXT - the translation
            // process will not emit code to call GETERRORTRAPPINGTOKEN since the source is not trying to trap any errors. However, any error information
            // must be cleared nonetheless, since there was an ON ERROR GOTO 0 in the source. It will also be required when Err.Clear is called.
            _trappedErrorIfAny = null;
        }

        public int GETERRORTRAPPINGTOKEN()
        {
            // Every time error-trapping is enabled within a function (or the outermost scope, where code doesn't run within a function in VBScript), the
            // translated code must request an "error trapping token". This is used to keep track of where error-trapping is and isn't enabled. If, for
            // example, a function F1 includes an ON ERROR RESUME NEXT and then calls F2 which includes its own ON ERROR RESUME NEXT and then later an
            // ON ERROR GOTO 0, this must only disable error-trapping within F2, the error-trapping that was enabled in F1 must continue to be enabled.
            // It isn't known at translation time how many error tokens may be required since this depends upon how the code executes - if F2 calls
            // itself then within its ON ERROR RESUME NEXT .. ON ERROR GOTO 0 region, an ON ERROR GOTO 0 call from that second call to F2 must not
            // disable error-trapping in the context of the first F2 call. So error tokens need to be handled dynamically. To try to only maintain as
            // many as strictly necessary, there is a queue of available tokens that is used to service GETERRORTRAPPINGTOKEN calls - after an error
            // token is returned (through RELEASEERRORTRAPPINGTOKEN), it goes back into the queue to potentially be used again. If the queue is empty
            // when this method is called then a new token is created. The token values are incremented each time this happens to ensure that they are
            // unique. This is why it's important that tokens are properly released - either when error-trapping is disabled (through an explicit ON
            // ERROR GOTO 0 or through an error being trapped or through a function scope ending where ON ERROR RESUME NEXT was set).
            // Note: When tokens are first requested, they default to the "OnErrorGoto0" state - meaning that error-trapping is not enabled currently
            // for that token. Error-trapping is enabled through a subsequent call to STARTERRORTRAPPINGANDCLEARANYERROR.
            int token;
            if (_availableErrorTokens.Count != 0)
                token = _availableErrorTokens.Dequeue();
            else
                token = _availableErrorTokens.Count + _activeErrorTokens.Count + 1;
            _activeErrorTokens.Add(token, ErrorTokenState.OnErrorGoto0);
            return token;
        }

        public void RELEASEERRORTRAPPINGTOKEN(int errorToken)
        {
            if (!_activeErrorTokens.ContainsKey(errorToken))
                throw new InvalidOperationException("This error token is not active - this indicates mismatched error token (de)registrations in the translated code");
            _activeErrorTokens.Remove(errorToken);
            _availableErrorTokens.Enqueue(errorToken);
        }

        public void STARTERRORTRAPPINGANDCLEARANYERROR(int errorToken)
        {
            // Note: Whenever error trapping is explicitly enabled or disabled, any error is cleared. If two methods are called within an OERN..
            //   ON ERROR RESUME Next
            //   F1()
            //   F2()
            // .. and F1() raises an error, that error's information will be maintained while F2 is called (if it is called without an error being
            // raised) unless F2 or any code it calls contains On Error Resume Next or On Error Goto - if this is the case then the error from F1
            // is lost forever. This is why _trappedErrorIfAny is set to null here and in STOPERRORTRAPPINGANDCLEARANYERROR.
            if (!_activeErrorTokens.ContainsKey(errorToken))
                throw new InvalidOperationException("This error token is not active - this indicates mismatched error token (de)registrations in the translated code");
            _activeErrorTokens[errorToken] = ErrorTokenState.OnErrorResumeNext;
            _trappedErrorIfAny = null;
        }

        public void STOPERRORTRAPPINGANDCLEARANYERROR(int errorToken)
        {
            if (!_activeErrorTokens.ContainsKey(errorToken))
                throw new InvalidOperationException("This error token is not active - this indicates mismatched error token (de)registrations in the translated code");
            _activeErrorTokens[errorToken] = ErrorTokenState.OnErrorGoto0;
            _trappedErrorIfAny = null;
        }

        // TODO: Disable debugger attribute? Does it help??
        public void HANDLEERROR(int errorToken, Action action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            if (!_activeErrorTokens.ContainsKey(errorToken))
                throw new InvalidOperationException("This error token is not active - this indicates mismatched error token (de)registrations in the translated code");

            try
            {
                action();
            }
            catch (Exception e)
            {
                // Translated programs shouldn't provide any actions that register or unregister error tokens, but since we've just gone off and
                // attempted to do some unknown work, it's best to check
                if (!_activeErrorTokens.TryGetValue(errorToken, out ErrorTokenState errorState))
                    throw new InvalidOperationException("This error token is not active - this indicates mismatched error token (de)registrations in the translated code", e);

                if (errorState == ErrorTokenState.OnErrorResumeNext)
                {
                    SETERROR(e);
                }
                else
                {
                    RELEASEERRORTRAPPINGTOKEN(errorToken);
                    throw;
                }
            }
        }

        /// <summary>
        /// This layers error-handling on top of the IAccessValuesUsingVBScriptRules.IF method, if error-handling is enabled for the specified
        /// token then evaluation of the value will be attempted - if an error occurs then it will be recorded and the condition will be treated
        /// as true, since this is VBScript's behaviour. It will throw an exception for a null valueEvaluator or an invalid errorToken.
        /// </summary>
        public bool IF(Func<object> valueEvaluator, int errorToken)
        {
            if (valueEvaluator == null)
                throw new ArgumentNullException(nameof(valueEvaluator));

            // VBScript's behaviour is quite mad here; if error-trapping is enabled when an IF condition must be evaluated, and if that evaluation results in
            // and error being raised, then act as if the condition was met. So we default to true and then try to perform the evalaluation with HANDLEERROR.
            // If an error is thrown and error-trapping is enabled, then true will be returned. If an error is throw an error-trapping is NOT enabled, then
            // that error will be allowed to propagate up. If there is no error raised then the result of the IF evaluation is returned.
            // - Note: In http://blogs.msdn.com/b/ericlippert/archive/2004/08/19/error-handling-in-vbscript-part-one.aspx, Eric Lippert does sort of
            //   describe this in passing (see the note that reads "If Blah raises an error then it resumes on the Print "Hello" in either case")
            bool result = true;
            HANDLEERROR(
                errorToken,
                () => { result = _valueRetriever.IF(valueEvaluator()); }
            );
            return result;
        }

        /// <summary>
        /// This is used by implementation of CINT, CSNG, CDBL and the like - it handles special cases of types such as Empty or booleans (and with error cases
        /// such as blanks string or VBScript Null) to try to extract a number. This number will be passed through the specified converter to ensure that it is
        /// translated into the desired type. If there are no applicable special cases then the value will be passed through the VAL function and then through
        /// the processor (if this fails then a TypeMismatchException will be raised).
        /// </summary>
        private T GetAsNumber<T>(object? value, string? optionalExceptionMessageForInvalidContent, Func<object, T> converter, bool rethrowUn = false) where T : struct
        {
            if (converter == null)
                throw new ArgumentNullException(nameof(converter));

            value = _valueRetriever.VAL(value, optionalExceptionMessageForInvalidContent);
            value = _valueRetriever.NUM(value);
            if (value is DateTime)
                value = DateToDouble((DateTime)value);
            if (value is T)
                return (T)value;
            try
            {
                return converter(value);
            }
            catch (OverflowException e)
            {
                throw new VBScriptOverflowException(Convert.ToDouble(value, CultureInfo.InvariantCulture), e);
            }
            catch (Exception e)
            {
                if (rethrowUn)
                    throw;
                else
                    throw new TypeMismatchException(optionalExceptionMessageForInvalidContent, e);
            }
        }

        /// <summary>
        /// Given a set of values, this will return nullable ints for each of the values - null if the value was a VBScript null (ie. DBNull.Value) and an int
        /// otherwise. Along with this set, it will return a lambda which will transform an int into the largest common bitwise-applicable value type that was
        /// encountered across the values (if all of the values were booleans, then this will transform an int into a boolean, if all of the values were booleans
        /// or bytes then it will transform an int into a byte - the range of types are, in ascending order: boolean, byte, Int16, Int32). If it is is not possible
        /// to translate any of the values into an int (or if there were no values specified) then an exception will be raised. This is because the VBScript "logical"
        /// operators actually perform bitwise operations, limiting the size of those numbers to int aka Int32 aka VBScript "Long" (so any number that won't fit into
        /// the range of an Int32 will result in an overflow). After VBScript performs the operation, it will return a value that relates to the inputs - so if two
        /// booleans were operated on then a boolean will be returned, if an "Integer" (Int16) or a "Long" (Int32) were operated on then an Int32 will be returned
        /// (this is what the lambda is for).
        /// </summary>
        private Tuple<IEnumerable<int?>, Func<int, object>, Type> GetForBitwiseOperations(string exceptionMessageForInvalidContent, params object?[] values)
        {
            if (values == null)
                throw new ArgumentNullException(nameof(values));
            if (values.Length == 0)
                throw new ArgumentException("At least one value must be specified");
            if (string.IsNullOrWhiteSpace(exceptionMessageForInvalidContent))
                throw new ArgumentException("Null/blank exceptionMessageForInvalidContent specified");

            // 1. Ensure that all values are of acceptable types (note that Empty will be parsed as a number, becoming an Int32 since it has no explicit type)
            //    and DBNull.Value will remain as DBNull.Value
            values = values.Select(v => _valueRetriever.VAL(v, exceptionMessageForInvalidContent)).ToArray();

            // 2. Determine the return type based upon all of the values types and generate a lambda that will transform an Int32 into this type
            //    - It seems that VBScript does not do anything as simple as choosing the smallest data type (boolean, byte, Int16, Int32).. while in MOST cases it does
            //      that (eg. boolean and Int16 => Int16, boolean and Int32 => Int32, Int16 and Int32 => Int32, boolean and Int32 => Int32) if there is a boolean and a
            //      byte then it jumps to Int16. In fairness, I imagine this is because byte is an unsigned type and so can not represent -1, which is what boolean True
            //      is represented by as a number - so Int16 is the smallest type that can contain all boolean AND byte values.
            Func<int, object> returnTypeConverter;
            Type returnType;
            if (values.All(v => v is bool))
            {
                returnType = typeof(bool);
                returnTypeConverter = finalValue => (finalValue != 0);
            }
            else if (values.All(v => v is byte))
            {
                returnType = typeof(byte);
                returnTypeConverter = finalValue => Convert.ToByte(finalValue & byte.MaxValue);
            }
            else if (values.All(v => (v is bool) || (v is byte) || (v is Int16)))
            {
                // This is the only complicated type conversion, really. To translate from an Int32 into an Int16, we want to take the last 8 (of 16) bits. For
                // example, if this is part of a NOT operation that is given an Int16 value of one, then that will be translated into a long (see below) and then
                // manipulated by the caller - in this case, changing from binary "0000000000000001" to "1111111111111110" (-2 in decimal). If we mask out the last
                // 8 bits then we get "11111110" and need only cast that to an Int16. This is why we can't use Convert.ToInt16 - since the long value we have manipulated
                // could easily cause an overflow (as it would in the example here).
                returnTypeConverter = finalValue => (Int16)(finalValue & 0xffff);
                returnType = typeof(Int16);
            }
            else
            {
                returnType = typeof(int);
                returnTypeConverter = finalValue => finalValue;
            }

            // 3. Return the values as null (where VBScript Null values were found) or as Int32 values - using the convention of a C# null for a VBScript null
            //    (despite the fact that they're not the same elsewhere VBScript Null = DBNull.Value in C#, VBScript Empty = null in C#) allows us to take
            //    advantage of the Nullable<int> type, rather than having to return IEnumerable<object> where everything is either Int32 or DBNull.Value
            return Tuple.Create(
                values.Select(v => (v == DBNull.Value) ? (int?)null : CLNG(v, exceptionMessageForInvalidContent)),
                returnTypeConverter,
                returnType
            );
        }

        private static bool IsDotNetNumericType(object? l)
        {
            if (l == null)
                return false;
            return
                IsDotNetIntegerType(l) ||
                (l is decimal) || (l is double) || (l is float);
        }

        private static bool IsDotNetIntegerType(object? l)
        {
            if (l == null)
                return false;
            if (l.GetType().IsEnum)
                return true;
            return (l is byte) || (l is char) || (l is int) || (l is long) || (l is sbyte) || (l is short) || (l is uint) || (l is ulong) || (l is ushort);
        }

        /// <summary>
        /// The comparison (o == VBScriptConstants.Nothing) will return false even if o is VBScriptConstants.Nothing due to the implementation details of
        /// DispatchWrapper. This method delivers a reliable way to test for it.
        /// </summary>
        private static bool IsVBScriptNothing(object? o)
        {
            return o != null && ((o is ScriptDispatchWrapper) && ((ScriptDispatchWrapper)o).WrappedObject == null);
        }

        private static double DateToDouble(DateTime value)
        {
            // When VBScript describes a date as a number, it applies somewhat counter-intuitive handling to the date component; 400.2 and -400.2 both
            // represent that same time (but on different days). Which means that -400.2 comes AFTER -400.0 chronologically, where as -400.2 comes
            // BEFORE -400 on the number scale. This behaviour needs to be reflected when we translate back from a DateTime to a double, the time
            // needs to be applied differently to a positive value than a negative. -400.2 is equivalent to "25/11/1898 04:48:00". If this date was
            // naively translated back into a number (by taking the total number of days, both whole and fractional, between the value and VBScript's
            // "zero date") then it would become -399.8. Instead the date component must be used to calculate -400 and then this SUBTRACTED from the
            // value for negatives, rather than added, so -400 becomes -400.2 (a subtraction of 0.2 from -400).
            double valueDouble;
            if (value < VBScriptConstants.ZeroDate)
                valueDouble = value.Date.Subtract(VBScriptConstants.ZeroDate).Subtract(value.TimeOfDay).TotalDays;
            else
                valueDouble = value.Date.Subtract(VBScriptConstants.ZeroDate).Add(value.TimeOfDay).TotalDays;
            return valueDouble;
        }

        /// <summary>
        /// VBScript has comparisons that will return true, false or Null (meaning DBNull.Value) which is a return type that is difficult to represent
        /// without resorting to "object" (which could be anything) or an enum (which wouldn't be the end of the world). I think the best approach,
        /// though, is to return a nullable bool from methods internally and then translate this for VBScript (so null becomes DBNull.Value).
        /// The same approach works for other non-nullable types.
        /// </summary>
        private static object ToVBScriptNullable<T>(T? value) where T : struct
        {
            if (value == null)
                return DBNull.Value;
            return value.Value;
        }

        // Feed all of these straight through to the _valueRetriever we have
        public IBuildCallArgumentProviders ARGS
        {
            get { return _valueRetriever.ARGS; }
        }
        public object? CALL(object? context, object target, IReadOnlyCollection<string> members, IProvideCallArguments argumentProvider, [CallerLineNumber] int line = 0)
        {
            return _valueRetriever.CALL(context, target, members, argumentProvider, line);// ?? VBScriptConstants.Nothing;
        }
        public void SET(object? context, object target, string? optionalMemberAccessor, IProvideCallArguments argumentProvider, object? valueToSetTo)
        {
            _valueRetriever.SET(context, target, optionalMemberAccessor, argumentProvider, valueToSetTo);
        }
        public bool IsVBScriptValueType(object o)
        {
            return _valueRetriever.IsVBScriptValueType(o);
        }
        public bool TryVAL(object? o, out bool parameterLessDefaultMemberWasAvailable, out object? asValueType)
        {
            return _valueRetriever.TryVAL(o, out parameterLessDefaultMemberWasAvailable, out asValueType);
        }
        public object? VAL(object? o, string? exceptionMessageForInvalidContent = null)
        {
            return _valueRetriever.VAL(o, exceptionMessageForInvalidContent);
        }
        public object OBJ(object o, string? optionalExceptionMessageForInvalidContent = null)
        {
            return _valueRetriever.OBJ(o, optionalExceptionMessageForInvalidContent);
        }
        public bool BOOL(object o, string? optionalExceptionMessageForInvalidContent = null)
        {
            return _valueRetriever.BOOL(o, optionalExceptionMessageForInvalidContent);
        }
        public object NUM(object? o, params object[] numericValuesTheTypeMustBeAbleToContain)
        {
            return _valueRetriever.NUM(o, numericValuesTheTypeMustBeAbleToContain);
        }
        public object NullableNUM(object o)
        {
            return _valueRetriever.NullableNUM(o);
        }
        public object NullableDATE(object o)
        {
            return _valueRetriever.NullableDATE(o);
        }
        public DateTime DATE(object? o, string? optionalExceptionMessageForInvalidContent = null)
        {
            return _valueRetriever.DATE(o);
        }
        public object NullableSTR(object o) => _valueRetriever.TryRetrieveStringOrEmpty(o);
        public string TryRetrieveStringOrEmpty(object? o) => _valueRetriever.TryRetrieveStringOrEmpty(o);

        public string STR(object? o, string? optionalExceptionMessageForInvalidContent = null)
        {
            return _valueRetriever.STR(o);
        }
        public bool IF(object o)
        {
            return _valueRetriever.IF(o);
        }
        public IEnumerable ENUMERABLE(object o)
        {
            return _valueRetriever.ENUMERABLE(o);
        }

        /// <summary>
        /// Where date literals were present in the source code, in a format that does not specify a date, they must be translated into dates at
        /// runtime. They must all be expanded to have whatever year it was when the request started - if the request happens to take sufficient
        /// time that the year ticks over during processing, all date literals (without explicit years) must be associated with the year when
        /// the request started. Note that if a new request starts with a different year, then date literals without years within that request
        /// must be associatedw ith the new year (this is consistent with how the VBScript interpreter would re-process the script each time).
        /// </summary>
        public DateParser DateLiteralParser { get; private set; }

        public T NnT<T>(T? targetInstance, string targetName) where T : class
        {
            if (targetInstance == null)
            {
                throw new InvalidOperationException($"Reference not set:({targetName})");
            }
            return targetInstance;
        }
        public object NnO(object? targetInstance, string targetName)
        {
            if (targetInstance == null)
            {
                throw new InvalidOperationException($"Reference not set:({targetName})");
            }
            return targetInstance;
        }
    }

    public interface IMyMoniker
    {
        object BindToObject(object? pbc, object? pmkToLeft, ref Guid riid);
    }
}
