using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using System;
using System.Collections.Generic;

namespace Skrypton.StageTwoParser.ExpressionParsing
{
    public class NewInstanceExpressionSegment : IExpressionSegment
    {
        public NewInstanceExpressionSegment(NameToken className)
        {
            ClassName = className ?? throw new ArgumentNullException(nameof(className));
        }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NameToken ClassName { get; private set; }

        /// <summary>
        /// This will never be null, empty or contain any null references
        /// </summary>
        IEnumerable<IToken> IExpressionSegment.AllTokens
        {
#pragma warning disable CA1033 // Interface methods should be callable by child types
            get
#pragma warning restore CA1033 // Interface methods should be callable by child types
            {
                return new IToken[]
                {
                    new KeyWordToken("new".ToUpperX(), ClassName.LineIndex),
                    ClassName
                };
            }
        }

        public string RenderedContent
        {
            get { return "new " + ClassName.Content; }
        }

        public override string ToString()
        {
            return base.ToString() + ":" + RenderedContent;
        }
    }
}
