using System;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This class is used by FunctionBlocks, ValueSettingStatements and other places where content is required that must represent a VBScript name
    /// </summary>
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public class NameToken : AtomToken
    {
        public NameToken(string content, int lineIndex) : this(false, content.ToUpperX(), lineIndex) { } // test

        public NameToken(bool alwaysFalse, StringUpper contentUpper, int lineIndex) : this(contentUpper, lineIndex)
        {
            if (alwaysFalse)
            {
                throw new InvalidOperationException();
            }

            if (this.GetType() != typeof(NameToken))
            {
                throw new InvalidOperationException();
            }

            // If this constructor is being called from a type derived from NameToken (eg. EscapedNameToken) then assume that all validation has been
            // performed in its constructor. If this constructor is being called to instantiate a new NameToken (and NOT a class derived from it) then
            // use the AtomToken's TryToGetAsRecognisedType method to try to ensure that this content is valid as a name and should not be for a token
            // of another type. This is process is kind of hokey but I'm trying to layer on a little additional type safety to some very old code so
            // I'm willing to live with this approach to it (the base class - the AtomToken - having knowledge of all of the derived types is not a
            // great design decision).
            /// if (this.GetType() == typeof(NameToken))
            {
                IToken recognisedType = TryToGetAsRecognisedType(contentUpper, lineIndex);
                if ((recognisedType != null) && !(recognisedType is NameToken))
                    throw new ArgumentException("Invalid content for a NameToken");
            }
        }

        protected NameToken(StringUpper contentUpper, int lineIndex)
            : this(contentUpper, WhiteSpaceBehaviourOptions.Disallow, lineIndex)
        {
        }

        protected NameToken(StringUpper contentUpper, WhiteSpaceBehaviourOptions whiteSpaceBehaviour, int lineIndex) : base(contentUpper, whiteSpaceBehaviour, lineIndex)
        {
        }

        public static int CompareNameTokens(NameToken x, NameToken y)
        {
            return CompareAtomTokens(x, y);
        }

    }
}
