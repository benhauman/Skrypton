using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public class BuiltInValueToken : AtomToken
    {
        public BuiltInValueToken(string content, int lineIndex) : this(content.ToUpperX(), lineIndex) { } // test
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the same token type while parsing the original content.
        /// </summary>
        public BuiltInValueToken(StringUpper contentUpper, int lineIndex) : base(contentUpper, WhiteSpaceBehaviourOptions.Disallow, lineIndex)
        {
            // Do all this validation (again) here in case this constructor wasn't called by the AtomToken.GetNewToken method
            if (contentUpper.IsNullOrWhiteSpace())
            {
                throw new ArgumentException("Null/blank content specified");
            }

            /// StringUpper contentUpper = content.ToUpperX();
            if (!AtomToken.isVBScriptValueUpper(contentUpper))
                throw new ArgumentException("Invalid content specified - not a VBScript value");

            IsAcceptableAsConstValue = _lowerCasedAcceptableBuiltinValues.Contains(Content.ToLower());
        }

        private readonly ReadOnlyCollection<string> _lowerCasedAcceptableBuiltinValues = new List<string> { "true", "false", "empty", "null", "nothing" }.AsReadOnly();

        /// <summary>
        /// The only acceptable values for a CONST value are  literals and a subset of the builtin VBScript values (such as true, false and empty but not including
        /// constants such as vbObjectError). If a token is not of a type that represents a numeric, string or date literal then it will be of this type if it is
        /// one of the other options that may be used as a CONST value - in which case this property will be useful if a caller needs to know whether it is
        /// acceptable-as-a-CONST-value or not
        /// </summary>
        [DataMember] public bool IsAcceptableAsConstValue { get; private set; }
    }
}
