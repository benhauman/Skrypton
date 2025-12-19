using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    /// <summary>
    /// This token represents a single unprocessed section of script content (not string or comment) - it is not initialised directly through a constructor,
    /// instead use the static GetNewToken method which try to will return an appropriate token type (an actual AtomToken, Operator Token or one of the
    /// AbstractEndOfStatement types)
    /// </summary>
    [Serializable]
    [DataContract(Namespace = "http://vbs")]
    public abstract class AtomToken : IToken
    {
        // =======================================================================================
        // CLASS INITIALISATION - INTERNAL
        // =======================================================================================
        protected AtomToken(StringUpper contentUpper, WhiteSpaceBehaviourOptions whiteSpaceBehaviour, int lineIndex)
        {
            // Do all this validation AGAIN because we may re-use this from inheriting classes (eg. OperatorToken)
            if (contentUpper == null)
                throw new ArgumentNullException("content");
            if (!Enum.IsDefined(typeof(WhiteSpaceBehaviourOptions), whiteSpaceBehaviour))
                throw new ArgumentOutOfRangeException("whiteSpaceBehaviour");
            if ((whiteSpaceBehaviour == WhiteSpaceBehaviourOptions.Disallow) && contentUpper.containsWhiteSpace())
                throw new ArgumentException("Whitespace encountered in AtomToken - invalid");
            if (contentUpper.Length == 0)
                throw new ArgumentException("Blank content specified for AtomToken - invalid");
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex", "must be zero or greater");

            Content = contentUpper.Original;
            LineIndex = lineIndex;
        }

        protected enum WhiteSpaceBehaviourOptions
        {
            Allow,
            Disallow
        }
        public static int CompareAtomTokens(AtomToken x, AtomToken y)
        {
            if (x.LineIndex != x.LineIndex)
                return 1;

            StringUpper x_Content_Upper = x.ContentUpperX();
            StringUpper y_Content_Upper = y.ContentUpperX();

            if (x_Content_Upper.Length != y_Content_Upper.Length)
                return 2;

            if (!string.Equals(x_Content_Upper.UpperText, y_Content_Upper.UpperText, StringComparison.OrdinalIgnoreCase))
                return 2;

            if (!(IsMustHandleKeyWord(x_Content_Upper) == IsMustHandleKeyWord(y_Content_Upper)))
                return 3;

            if (!(isVBScriptFunctionUpper(x_Content_Upper) == isVBScriptFunctionUpper(y_Content_Upper)))
                return 4;

            if (!(IsVBScriptSymbolUpper(x_Content_Upper) == IsVBScriptSymbolUpper(y_Content_Upper)))
                return 5;

            if (!(isVBScriptValueUpper(x_Content_Upper) == isVBScriptValueUpper(y_Content_Upper)))
                return 6;

            return 0; // equal
        }
        // =======================================================================================
        // CLASS INITIALISATION - PUBLIC
        // =======================================================================================
        /// <summary>
        /// This will return an AtomToken, OperatorToken, EndOfStatementNewLineToken or
        /// EndOfStatementSameLineToken if the content appears valid (it must be non-
        /// null, non-blank and contain no whitespace - unless it's a single line-
        /// return)
        /// </summary>
        public static IToken GetNewToken(StringUpper contentUpper, int lineIndex)
        {
            if (contentUpper.Length == 0)
                throw new ArgumentException("Blank content specified for AtomToken - invalid");

            return GetNewTokenCore(contentUpper, lineIndex);
        }
        public static IToken GetNewToken(KnownTextContent content, int lineIndex)
        {
            return GetNewTokenCore(content.TheContentUpper, lineIndex);
        }
        private static IToken GetNewTokenCore(StringUpper contentUpper, int lineIndex)
        {
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex", "must be zero or greater");

            var recognisedType = TryToGetAsRecognisedType(contentUpper, lineIndex);
            if (recognisedType != null)
                return recognisedType;

            if (contentUpper.Original.StartsWith("["))
            {
                if (!contentUpper.Original.EndsWith("]"))
                    throw new ArgumentException("If content starts with a square bracket then it must have a closing bracket to indicate an escaped-name variable");
                return new EscapedNameToken(contentUpper, lineIndex);
            }

            if (contentUpper.containsWhiteSpace())
                throw new ArgumentException("Whitespace in an AtomToken - invalid");

            return new NameToken(false, contentUpper, lineIndex);
        }

        /// <summary>
        /// This will try to identify the token content as a VBScript operator or comparison or built-in function or value or line return or statement
        /// separator or numeric value. If unable to match its type then it will return null - this should indicate the name of a function, property,
        /// variable, etc.. defined in the source code being processed.
        /// </summary>
        protected static IToken TryToGetAsRecognisedType(StringUpper contentUpper, int lineIndex)
        {
            if (lineIndex < 0)
                throw new ArgumentOutOfRangeException("lineIndex", "must be zero or greater");


            if (contentUpper.Length == 1)
            {
                if (contentUpper.Original == "\n")
                    return new EndOfStatementNewLineToken(lineIndex);
                if (contentUpper.Original == ":")
                    return new EndOfStatementSameLineToken(lineIndex);
            }


            if (isMustHandleKeyWordUpper(contentUpper) || isMiscKeyWordUpper(contentUpper))
                return new KeyWordToken(contentUpper, lineIndex);
            if (isContextDependentKeywordUpper(contentUpper))
                return new MayBeKeywordOrNameToken(contentUpper, lineIndex);
            if (isVBScriptFunctionUpper(contentUpper))
                return new BuiltInFunctionToken(contentUpper, lineIndex);
            if (isVBScriptValueUpper(contentUpper))
                return new BuiltInValueToken(contentUpper, lineIndex);
            if (isLogicalOperatorUpper(contentUpper))
                return new LogicalOperatorToken(contentUpper, lineIndex);
            if (isComparisonUpper(contentUpper))
                return new ComparisonOperatorToken(contentUpper, lineIndex);
            if (isOperatorUpper(contentUpper))
                return new OperatorToken(contentUpper, lineIndex);
            if (isMemberAccessorUpper(contentUpper))
                return new MemberAccessorOrDecimalPointToken(contentUpper, lineIndex);
            if (isArgumentSeparatorUpper(contentUpper))
                return new ArgumentSeparatorToken(lineIndex);
            if (isOpenBraceUpper(contentUpper))
                return new OpenBrace(lineIndex);
            if (isCloseBraceUpper(contentUpper))
                return new CloseBrace(lineIndex);
            if (isTargetCurrentClassTokenUpper(contentUpper))
                return new TargetCurrentClassToken(lineIndex);

            double numericValue;
            if (double.TryParse(contentUpper.Original, out numericValue))
                return new NumericValueToken(contentUpper, lineIndex);
            if (contentUpper.Original.StartsWith("&h", StringComparison.InvariantCultureIgnoreCase))
            {
                int numericHexValue;
                if (int.TryParse(contentUpper.Original.Substring(2), NumberStyles.HexNumber, null, out numericHexValue))
                    return new NumericValueToken(numericHexValue.ToString().ToUpperX(), lineIndex);
            }

            return null;
        }

        /// private static string WhiteSpaceChars = new string(
        /// Enumerable.Range((int)char.MinValue, (int)char.MaxValue).Select(v => (char)v).Where(c => char.IsWhiteSpace(c)).ToArray()
        /// );

        protected static bool containsWhiteSpace(string content)
        {
            if (content == null)
                throw new ArgumentNullException("token");

            /// return content.Any(c => WhiteSpaceChars.IndexOf(c) != -1);
            for (int ix = 0; ix < content.Length; ix++)
            {
                char ccc = content[ix];
                if (char.IsWhiteSpace(ccc))
                    return true;
            }
            return false;
        }

        // =======================================================================================
        // [PRIVATE] CONTENT DETERMINATION - eg. isOperator
        // =======================================================================================
        /// <summary>
        /// This will not be null, empty, contain any null or blank values, any duplicates or any content containing whitespace. These are ordered
        /// according to the precedence that the VBScript interpreter will give to them when multiple occurences are encountered within an expression
        /// (see http://msdn.microsoft.com/en-us/library/6s7zy3d1(v=vs.84).aspx).
        /// </summary>
        public static IEnumerable<string> ArithmeticAndStringOperatorTokenValues = new List<string>
        {
            "^", "/", "\\", "*", "\"", "MOD", "+", "-", "&" // Note: "\" is integer division (see the link above)
		}.AsReadOnly();

        /// This will not be null, empty, contain any null or blank values, any duplicates or any content containing whitespace. These are ordered
        /// according to the precedence that the VBScript interpreter will give to them when multiple occurences are encountered within an expression
        /// (see http://msdn.microsoft.com/en-us/library/6s7zy3d1(v=vs.84).aspx).
        public static IEnumerable<string> LogicalOperatorTokenValues = new List<string>
        {
            "NOT", "AND", "OR", "XOR"
        }.AsReadOnly();

        /// <summary>
        /// Does the content appear to represent a VBScript operator (eg. an arithermetic operator such as "*", a logical operator such as "AND" or
        /// a comparison operator such as ">")? An exception will be raised for null, blank or whitespace-containing input.
        /// </summary>
        internal static bool isOperatorUpper(StringUpper atomContent)
        {
            return KnownTextResolver.isOperatorUpper(atomContent).HasValue;
        }

        /// <summary>
        /// Does the content appear to represent a VBScript operator (eg. AND)? An exception will be raised
        /// for null, blank or whitespace-containing input.
        /// </summary>
        internal static bool isLogicalOperatorUpper(StringUpper atomContent)
        {
            return KnownTextResolver.isLogicalOperatorUpper(atomContent).HasValue;
        }

        /// <summary>
        /// This will not be null, empty, contain any null or blank values, any duplicates or any content containing whitespace. These are ordered
        /// according to the precedence that the VBScript interpreter will give to them when multiple occurences are encountered within an expression
        /// (see http://msdn.microsoft.com/en-us/library/6s7zy3d1(v=vs.84).aspx).
        /// </summary>
        public static IEnumerable<string> ComparisonTokenValues = new List<string>
        {
            "=", "<>", "<", ">", "<=", ">=", "IS",
            "EQV", "IMP"
        }.AsReadOnly();

        /// <summary>
        /// Does the content appear to represent a VBScript comparison? An exception will be raised
        /// for null, blank or whitespace-containing input.
        /// </summary>
        internal static bool isComparisonUpper(StringUpper atomContent)
        {
            return KnownTextResolver.isComparisonUpper(atomContent).HasValue;
        }

        internal static bool isMemberAccessorUpper(StringUpper atomContent)
        {
            return KnownTextResolver.isMemberAccessorUpper(atomContent);
        }

        private static bool isArgumentSeparatorUpper(StringUpper atomContent)
        {
            return KnownTextResolver.isArgumentSeparatorUpper(atomContent);
        }

        private static bool isOpenBraceUpper(StringUpper atomContent)
        {
            return KnownTextResolver.isOpenBraceUpper(atomContent);
        }

        private static bool isCloseBraceUpper(StringUpper atomContent)
        {
            return KnownTextResolver.isCloseBraceUpper(atomContent);
        }

        private static bool isTargetCurrentClassTokenUpper(StringUpper atomContent)
        {
            return KnownTextResolver.isTargetCurrentClassTokenUpper(atomContent);
        }

        /// <summary>
        /// Does the content appear to represent a VBScript keyword that will have to be handled by an
        /// AbstractBlockHandler - eg. a "FOR" in a loop, or the "OPTION" from "OPTION EXPLICIT" or the
        /// "RANDOMIZE" command? An exception will be raised for null, blank or whitespace-containing
        /// input.
        /// </summary>
        internal static bool isMustHandleKeyWordUpper(StringUpper atomContent)
        {
            return KnownTextResolver.isMustHandleKeyWordUpper(atomContent);
        }

        /// <summary>
        /// There are some keywords that can be used as variable names, deeper parsing work than looking
        /// solely at its content is required to determine if a token is a NameToken or KeywordToken
        /// </summary>
        internal static bool isContextDependentKeywordUpper(StringUpper atomContent)
        {
            return KnownTextResolver.isContextDependentKeywordUpper(atomContent);
        }

        /// <summary>
        /// Does the content appear to represent a VBScript keyword that may form part of a general
        /// statement and not have to be handled by a specific AbstractBlockHandler - eg. a "NEW"
        /// declaration for instantiating a class instance. An exception will be raised for null,
        /// blank or whitespace-containing input.
        /// </summary>
        internal static bool isMiscKeyWordUpper(StringUpper atomContent)
        {
            return KnownTextResolver.isMiscKeyWordUpper(atomContent);
        }

        /// <summary>
        /// Does the content appear to represent a VBScript expression - eg. "NOTHING". An exception will be raised for null, blank or whitespace-containing input.
        /// </summary>
        internal static bool isVBScriptValueUpper(StringUpper atomContent)
        {
            return KnownTextResolver.isVBScriptValueUpper(atomContent);
        }

        /// <summary>
        /// Does the content appear to represent a VBScript function - eg. the "ISNULL" method.
        /// An exception will be raised for null, blank or whitespace-containing input.
        /// </summary>
        internal static bool isVBScriptFunctionUpper(StringUpper atomContent)
        {
            return KnownTextResolver.isVBScriptFunctionUpper(atomContent);
        }

        /// <summary>
        /// Does the content appear to represent a VBScript function that is guaranteed to return a numeric value - eg. the "CDBL" method. This does
        /// not include functions that will ever return Null (such as ABS, which returns Null if Null is provided as the argument). An exception will
        /// be raised for null, blank or whitespace-containing input.
        /// </summary>
        internal static bool isVBScriptFunctionThatAlwaysReturnsNumericContentUpper(StringUpper atomContent)
        {
            return KnownTextResolver.isVBScriptFunctionThatAlwaysReturnsNumericContentUpper(atomContent);
            ///// These must ONLY include those that will never return null
            /// return isType(
			/// 	atomContent,
			/// 	new string[]
			/// 	{
			/// 		"LBOUND", "UBOUND",
			/// 		"CBYTE", "CCUR", "CINT", "CLNG", "CSNG", "CDBL", "CDATE",
			/// 		"DATEADD", "DATESERIAL", "TIMESERIAL",
			/// 		"NOW", "DATE", "TIME",
			/// 		"ATN", "COS", "SIN", "TAN", "EXP", "LOG", "SQR", "RND", "ROUND",
			/// 		"HEX", "OCT", "FIX", "INT", "SNG",
			/// 		"ASC", "ASCB", "ASCW",
			/// 		"SCRIPTENGINEBUILDVERSION", "SCRIPTENGINEMAJORVERSION", "SCRIPTENGINEMINORVERSION",
			/// 		"TIMER"
			/// 	}
			/// );
		}

        // =======================================================================================
        // PUBLIC DATA ACCESS
        // =======================================================================================
        /// <summary>
        /// This will never be blank or null
        /// </summary>
        [DataMember] public string Content { get; private set; }

        [NonSerialized] StringUpper contentUpper;
        public StringUpper ContentUpperX()

        {
            if (contentUpper == null)
                contentUpper = Content.ToUpperX();
            return contentUpper;
        }

        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        [DataMember] public int LineIndex { get; private set; }

        /// <summary>
        /// Does this AtomContent describe a reserved VBScript keyword or operator?
        /// </summary>
        internal static bool IsVBScriptSymbolUpper(StringUpper ContentUpper)
        {

            {
                //string ContentUpper = Content.ToUpperInvariant();
                // Note: isContextDependentKeyword is not consulted here since it the values there are not reserved in all cases
                return
                    isMustHandleKeyWordUpper(ContentUpper) ||
                    isMiscKeyWordUpper(ContentUpper) ||
                    isComparisonUpper(ContentUpper) ||
                    isOperatorUpper(ContentUpper) ||
                    isMemberAccessorUpper(ContentUpper) ||
                    isArgumentSeparatorUpper(ContentUpper) ||
                    isOpenBraceUpper(ContentUpper) ||
                    isCloseBraceUpper(ContentUpper) ||
                    isVBScriptFunctionUpper(ContentUpper) ||
                    isVBScriptValueUpper(ContentUpper);
            }
        }

        /// <summary>
        /// Does this AtomContent describe a reserved VBScript keyword that must be handled by
        /// a targeted AbstractCodeBlockHandler? (eg. "FOR", "DIM")
        /// </summary>
        internal static bool IsMustHandleKeyWord(StringUpper ContentUpper)
        {
            return isMustHandleKeyWordUpper(ContentUpper);
        }

        /// <summary>
        /// Does this AtomContent describe a VBScript (eg. "ABS")?
        /// </summary>
        internal bool IsVBScriptFunctionUpper(StringUpper ContentUpper)
        {
            return isVBScriptFunctionUpper(ContentUpper);
        }

        /// <summary>
        /// Does this AtomContent describe a VBScript value (eg. "NOTHING")?
        /// </summary>
        internal bool IsVBScriptValueUpper(StringUpper ContentUpper)
        {
            return KnownTextResolver.isVBScriptValueUpper(ContentUpper);
        }

        public override string ToString()
        {
            return base.ToString() + ":" + Content;
        }
    }
}
