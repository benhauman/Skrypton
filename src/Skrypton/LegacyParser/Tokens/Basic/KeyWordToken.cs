using System;
using System.Collections.Generic;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    [Serializable]
    public sealed class KeyWordToken : AtomToken
    {
        [NonSerialized] private readonly KnownKeyWordId _keywordId;
        /// <summary>
        /// This inherits from AtomToken since a lot of processing would consider them the
        /// same token type while parsing the original content.
        /// </summary>
        public KeyWordToken(StringUpper contentUpper, int lineIndex) : base(contentUpper, WhiteSpaceBehaviourOptions.Disallow, lineIndex)
        {
            // Do all this validation (again) here in case this constructor wasn't called by the AtomToken.GetNewToken method
            if (contentUpper.Length == 0)
                throw new ArgumentException("Null/blank content specified");
            if (!AtomToken.isMustHandleKeyWordUpper(contentUpper) && !AtomToken.isContextDependentKeywordUpper(contentUpper) && !AtomToken.isMiscKeyWordUpper(contentUpper))
                throw new ArgumentException("Invalid content specified - not a VBScript keyword");

            _keywordId = ParseKeyword(contentUpper);
        }
        public KeyWordToken(string content, int lineIndex) : this(content.ToUpperX(), lineIndex) { } // test

        internal KnownKeyWordId KeyWordId => _keywordId;

        internal static KnownKeyWordId ParseKeyword(StringUpper contentUpper)
        {
            return KeywordsTextIdMap.TryGetValue(contentUpper.UpperText, out var keywordId)
                ? keywordId
                : throw new ArgumentException("Unknown keyword:" + contentUpper.UpperText, nameof(contentUpper));
        }
        private static readonly Dictionary<string, KnownKeyWordId> KeywordsTextIdMap = new Dictionary<string, KnownKeyWordId>()
        {
            { "CALL", KnownKeyWordId.KeywordCall },
            { "CASE", KnownKeyWordId.KeywordCase },
            { "CLASS", KnownKeyWordId.KeywordClass },
            { "DIM", KnownKeyWordId.KeywordDim },
            { "DO", KnownKeyWordId.KeywordDo },
            { "EACH", KnownKeyWordId.KeywordEach },
            { "ELSE", KnownKeyWordId.KeywordElse },
            { "ELSEIF", KnownKeyWordId.KeywordElseIf },
            { "END", KnownKeyWordId.KeywordEnd },
            { "ERASE", KnownKeyWordId.KeywordErase },
            { "EXIT", KnownKeyWordId.KeywordExit },
            { "FOR", KnownKeyWordId.KeywordFor },
            { "FUNCTION", KnownKeyWordId.KeywordFunction },
            { "GET", KnownKeyWordId.KeywordGet },
            { "IF", KnownKeyWordId.KeywordIf },
            { "LET", KnownKeyWordId.KeywordLet },
            { "LOOP", KnownKeyWordId.KeywordLoop },
            { "NEXT", KnownKeyWordId.KeywordNext },
            { "NEW", KnownKeyWordId.KeywordNew },
            { "ON", KnownKeyWordId.KeywordOn },
            { "OPTION", KnownKeyWordId.KeywordOption },
            { "PUBLIC", KnownKeyWordId.KeywordPublic },
            { "PRESERVE", KnownKeyWordId.KeywordPreserve },
            { "PRIVATE", KnownKeyWordId.KeywordPrivate },
            { "REDIM", KnownKeyWordId.KeywordReDim },
            { "RESUME", KnownKeyWordId.KeywordResume },
            { "SELECT", KnownKeyWordId.KeywordSelect },
            { "SET", KnownKeyWordId.KeywordSet },
            { "STEP", KnownKeyWordId.KeywordStep },
            { "SUB", KnownKeyWordId.KeywordSub },
            { "THEN", KnownKeyWordId.KeywordThen },
            { "TO", KnownKeyWordId.KeywordTo },
            { "WHILE", KnownKeyWordId.KeywordWhile },
            { "WITH", KnownKeyWordId.KeywordWith },
            { "UNTIL", KnownKeyWordId.KeywordUntil },
        };
    }

    public enum KnownKeyWordId
    {
        Unknown = 0,
        KeywordCall,
        KeywordCase,
        KeywordClass,
        KeywordDim,
        KeywordDo,
        KeywordEach,
        KeywordElse,
        KeywordElseIf,
        KeywordEnd,
        KeywordErase,
        KeywordExit,
        KeywordFor,
        KeywordFunction,
        KeywordGet,
        KeywordIf,
        KeywordLet,
        KeywordLoop,
        KeywordNext,
        KeywordNew,
        KeywordOn,
        KeywordOption,
        KeywordPublic,
        KeywordPreserve,
        KeywordPrivate,
        KeywordReDim,
        KeywordResume,
        KeywordSelect,
        KeywordSet,
        KeywordStep,
        KeywordSub,
        KeywordTo,
        KeywordThen,
        KeywordWhile,
        KeywordWith,
        KeywordUntil,
    }
}
