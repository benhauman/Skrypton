using Skrypton.CSharpWriter.CodeTranslation;
using Skrypton.LegacyParser.Tokens.Basic;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Skrypton.Tests.Shared.Comparers
{
	public class TranslatedStatementContentDetailsComparer : IEqualityComparer<TranslatedStatementContentDetails>
	{
        internal static readonly TranslatedStatementContentDetailsComparer Instance = new TranslatedStatementContentDetailsComparer();

        public TranslatedStatementContentDetailsComparer() // TODO: private
        {
        }
        public bool Equals(TranslatedStatementContentDetails x, TranslatedStatementContentDetails y)
		{
			if (x == null)
				throw new ArgumentNullException(nameof(x));
			if (y == null)
				throw new ArgumentNullException(nameof(y));
			return EqualsX(x.TranslatedContent, x.VariablesAccessed, y);
		}
        public static bool EqualsX(string exTranslatedContent, IReadOnlyCollection<NameToken> exVariablesAccessed, TranslatedStatementContentDetails y)
        {
            if (y == null)
                throw new ArgumentNullException(nameof(y));

            if (exTranslatedContent != y.TranslatedContent)
                return false;

            return TokenSetComparer.EqualsX(
                exVariablesAccessed.Distinct(TokenComparer.Instance),
                y.VariablesAccessed.Distinct(TokenComparer.Instance)
            );
        }

        public int GetHashCode(TranslatedStatementContentDetails obj)
		{
			if (obj == null)
				throw new ArgumentNullException(nameof(obj));

			return 0;
		}
	}
}
