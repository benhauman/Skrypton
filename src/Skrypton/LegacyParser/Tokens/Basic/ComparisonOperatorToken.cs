using System;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public sealed class ComparisonOperatorToken : OperatorToken
    {
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the
        /// same token type while parsing the original content.
        /// </summary>
        public ComparisonOperatorToken(OperatorKind comparisonOperatorKind, string content, int lineIndex) : this(comparisonOperatorKind, content.ToUpperX(), lineIndex)
        {
        }
        public ComparisonOperatorToken(OperatorKind comparisonOperatorKind, StringUpper contentUpper, int lineIndex) : base(contentUpper, lineIndex)
        {
            if (contentUpper == null) throw new ArgumentNullException(nameof(contentUpper));
            // Do all this validation (again) here in case this constructor wasn't called
            // by the AtomToken.GetNewToken method
            if (contentUpper.Length == 0)
                throw new ArgumentException("Blank content specified for ComparisonToken - invalid", nameof(contentUpper));
            if (!AtomToken.isComparisonUpperX(contentUpper, out var cmpOpX))
                throw new ArgumentException("Invalid content specified - not a Comparison", nameof(contentUpper));

            if (comparisonOperatorKind != cmpOpX)
            {
                throw new ArgumentException("Mismatched operator and content specified", nameof(contentUpper));
            }
        }
    }
}
