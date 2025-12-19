using System;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public class CloseBrace : AtomToken
    {
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the
        /// same token type while parsing the original content.
        /// </summary>
        public CloseBrace(int lineIndex) : base(")".ToUpperX(), WhiteSpaceBehaviourOptions.Disallow, lineIndex) { }
    }
}
