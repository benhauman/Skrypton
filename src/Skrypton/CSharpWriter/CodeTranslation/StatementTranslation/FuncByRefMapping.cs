using System;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.CSharpWriter.CodeTranslation.StatementTranslation
{
    public class FuncByRefMapping
    {
        public FuncByRefMapping(NameToken from, CSharpName to, bool mappedValueIsReadOnly)
        {
            From = from ?? throw new ArgumentNullException(nameof(from));
            To = to ?? throw new ArgumentNullException(nameof(to));
            MappedValueIsReadOnly = mappedValueIsReadOnly;
        }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NameToken From { get; private set; }

        /// <summary>
        /// This will never be null
        /// </summary>
        public CSharpName To { get; private set; }

        /// <summary>
        /// Some mappings are only required to make the data available from a ByRef argument reference, and so creating an alias is all that is required. In some
        /// cases, however, any changes to the alias must be reflected back onto the original reference - this is only applicable when the alis is passed as a
        /// "ref" argument to another function. In the first case, this value will be true, the mapped value IS ready only, while in the second case this
        /// value will be false.
        /// true:Relevant for error handling and what happens on 'restore'
        /// </summary>
        public bool MappedValueIsReadOnly { get; private set; }
    }
}
