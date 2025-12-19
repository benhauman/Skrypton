using System;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public class LogicalOperatorToken : OperatorToken
    {
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the
        /// same token type while parsing the original content.
        /// </summary>
        public LogicalOperatorToken(StringUpper contentUpper, int lineIndex) : base(contentUpper, lineIndex)
        {
            // Do all this validation (again) here in case this constructor wasn't called
            // by the AtomToken.GetNewToken method
            if (contentUpper.Length == 0)
                throw new ArgumentException("Blank content specified for OperatorToken - invalid");
            /// StringUpper contentUpper = content.ToUpperX();
            if (!AtomToken.isLogicalOperatorUpper(contentUpper))
                throw new ArgumentException("Invalid content specified - not a Logical Operator");
        }
        public LogicalOperatorToken(string content, int lineIndex) : this(content.ToUpperX(), lineIndex) { } // test
    }
}
