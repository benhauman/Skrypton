using System;
using System.Collections.Generic;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.StageTwoParser.ExpressionParsing
{
    public class NewInstanceExpressionSegment : IExpressionSegment
    {
        public NewInstanceExpressionSegment(NameToken className)
        {
            if (className == null)
                throw new ArgumentNullException("className");

            ClassName = className;
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
            get
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
