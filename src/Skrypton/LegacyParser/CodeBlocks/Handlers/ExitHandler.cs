using Skrypton.LegacyParser.CodeBlocks.Basic;
using Skrypton.LegacyParser.Tokens;
using Skrypton.LegacyParser.Tokens.Basic;
using System;
using System.Collections.Generic;

namespace Skrypton.LegacyParser.CodeBlocks.Handlers
{
    public sealed class ExitHandler : AbstractBlockHandler // public due to tests
    {
        private static readonly (ExitStatement.ExitableStatementType ExitType, string Name)[] s_exitTypes = InitializeExitTypes(); // ToString() called once here
        private static (ExitStatement.ExitableStatementType, string)[] InitializeExitTypes()
        {
            var values = (ExitStatement.ExitableStatementType[])Enum.GetValues(typeof(ExitStatement.ExitableStatementType));
            var arr = new (ExitStatement.ExitableStatementType, string)[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                arr[i] = (values[i], values[i].ToString());
            }

            return arr;
        }
        /// <summary>
        /// The token list will be edited in-place as handlers are able to deal with the content, so the input list should expect to be mutated
        /// </summary>
        public override ICodeBlock? Process(List<IToken> tokens)
        {
            if (tokens == null)
                throw new ArgumentNullException(nameof(tokens));
            if (tokens.Count == 0)
                return null;

            for (var ixPair = 0; ixPair < s_exitTypes.Length; ixPair++)
            {
                var exitPair = s_exitTypes[ixPair];
                string[] matchPattern = new string[] { "EXIT", exitPair.Name };
                if (checkAtomTokenPattern(tokens, matchPattern, false))
                {
                    var lineIndexOfExit = tokens[0].LineIndex;
                    var requireAnEndOfStatementToken = (tokens.Count > matchPattern.Length);
                    if (requireAnEndOfStatementToken)
                    {
                        if (!(tokens[matchPattern.Length] is AbstractEndOfStatementToken))
                            throw new InvalidOperationException("EXIT statement wasn't followed by end-of-statement token");
                    }

                    tokens.RemoveRange(0, matchPattern.Length);
                    if (requireAnEndOfStatementToken)
                        tokens.RemoveRange(0, 1);
                    return new ExitStatement(exitPair.ExitType, lineIndexOfExit);
                }
            }

            return null;
        }
    }
}
