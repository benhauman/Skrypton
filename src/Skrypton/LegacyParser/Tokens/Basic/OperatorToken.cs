using System;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public class OperatorToken : AtomToken
    {
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the same token type while parsing the original content.
        /// </summary>
        public OperatorToken(StringUpper contentUpper, int lineIndex) : base(contentUpper, WhiteSpaceBehaviourOptions.Disallow, lineIndex)
        {
            // Do all this validation (again) here in case this constructor wasn't called by the AtomToken.GetNewToken method
            if (!AtomToken.isOperatorUpper(contentUpper))
                throw new ArgumentException("Invalid content specified - not an Operator");
            if (AtomToken.isLogicalOperatorUpper(contentUpper) && (!(this is LogicalOperatorToken)))
                throw new ArgumentException("This content indicates a LogicalOperatorToken but this instance is not of that type");
            if (AtomToken.isComparisonUpper(contentUpper) && (!(this is ComparisonOperatorToken)))
                throw new ArgumentException("This content indicates a ComparisonOperatorToken but this instance is not of that type");
        }
        public OperatorToken(string content, int lineIndex) : this(content.ToUpperX(), lineIndex) { } // test
    }
}
