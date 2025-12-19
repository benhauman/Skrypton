using System;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This represents a member accessor, such as a function or property - eg. the period in target.Name or target.Method() - or the
    /// decimal point in a number - eg. in "1.1"
    /// </summary>
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public class MemberAccessorOrDecimalPointToken : AtomToken
    {
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the
        /// same token type while parsing the original content.
        /// </summary>
        public MemberAccessorOrDecimalPointToken(string content, int lineIndex) : this(content.ToUpperX(), lineIndex) { } // test
        public MemberAccessorOrDecimalPointToken(StringUpper contentUpper, int lineIndex) : base(contentUpper, WhiteSpaceBehaviourOptions.Disallow, lineIndex)
        {
            // Do all this validation (again) here in case this constructor wasn't called
            // by the AtomToken.GetNewToken method
            if (contentUpper.Length == 0)
                throw new ArgumentException("Blank content specified for MemberAccessorToken - invalid");
            if (!AtomToken.isMemberAccessorUpper(contentUpper))
                throw new ArgumentException("Invalid content specified - not a MemberAccessor");
        }
    }
}
