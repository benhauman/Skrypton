using System;
using System.Collections.Generic;
using System.Text;
using Skrypton.LegacyParser.Tokens;

namespace Skrypton.LegacyParser
{
    public sealed class SourceProcessingError
    {
        internal readonly int LineNumber;
        internal readonly string Message;
        public SourceProcessingError(int lineNumber, string message)
        {
            this.LineNumber = lineNumber;
            this.Message = message;
        }
    }

    // see SyntaxError
    public sealed class SourceProcessingException : Exception // see SyntaxError
    {
        public int LineNumber { get; private set; }
        public SourceProcessingException(SourceProcessingError error)
            : base(error.Message)
        {
            this.LineNumber = error.LineNumber;
        }

        public SourceProcessingException(string message) : base(message)
        {
        }

        public SourceProcessingException(string message, Exception innerException) : base(message, innerException)
        {
        }

        internal static SourceProcessingException Create(IFragment fragment, int lineNumber, string message)
        {
            StringBuilder error_text_builder = new StringBuilder();
            error_text_builder.Append("Line " + lineNumber);
            error_text_builder.Append(" " + "fragment-id:" + fragment.FragmentId);
            error_text_builder.Append(" " + "fragment-type:" + fragment.GetType().Name);
            error_text_builder.Append(" " + message);
            return new SourceProcessingException(new SourceProcessingError(lineNumber, error_text_builder.ToString()));
        }
        internal static SourceProcessingException Create(IToken token, string message)
        {
            StringBuilder error_text_builder = new StringBuilder();
            error_text_builder.Append("Line " + token.LineIndex);
            //error_text_builder.Append(" " + "token-id:" + token.LineIndex);
            error_text_builder.Append(" " + "token-type:" + token.GetType().Name);
            error_text_builder.Append(" " + message);
            return new SourceProcessingException(new SourceProcessingError(token.LineIndex, error_text_builder.ToString()));
        }
    }

    internal interface IProcessingTokens : IEnumerable<IToken> // see List<IToken>
    {
        int Count { get; }

        //IList<IToken> d;
        IToken this[int index] { get; }

        void RemoveRange(int index, int length);
        void RemoveAt(int index);
        void Clear();
    }

}