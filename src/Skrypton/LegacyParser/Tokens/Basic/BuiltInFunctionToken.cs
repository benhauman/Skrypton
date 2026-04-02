using System;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    [DataContract(Namespace = "http://vbs")]
    public sealed class BuiltInFunctionToken : AtomToken
    {
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the
        /// same token type while parsing the original content.
        /// </summary>
        internal BuiltInFunctionToken(StringUpper contentUpper, int lineIndex) : base(contentUpper, WhiteSpaceBehaviourOptions.Disallow, lineIndex)
        {
            // Do all this validation (again) here in case this constructor wasn't called by the AtomToken.GetNewToken method
            if (contentUpper.IsNullOrWhiteSpace())
                throw new ArgumentException("Null/blank content specified");
            if (!AtomToken.isVBScriptFunctionUpper(contentUpper, out BuiltInFunctionInfo? binFun))
                throw new ArgumentException("Invalid content specified - not a VBScript function");
            FunctionId = binFun!.FunctionId;
        }
        public static BuiltInFunctionToken TestCreateBuiltInFunctionTokenTest(string content, int lineIndex) => new BuiltInFunctionToken(content.ToUpperX(), lineIndex); // test

        internal BuiltInFunctionId FunctionId { get; }

        /// <summary>
        /// Is this a function that will always return a numeric value (or raise an error)? This will not return true for functions such as ABS
        /// which return VBScript Null in some cases, it will only apply to functions which always return a "true" number (eg. CDBL).
        /// </summary>
        public bool GuaranteedToReturnNumericContent()
        {
            return AtomToken.isVBScriptFunctionThatAlwaysReturnsNumericContentUpper(Content.ToUpperX());
        }
    }
}
