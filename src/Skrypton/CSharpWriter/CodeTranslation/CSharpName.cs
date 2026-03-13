using System;
using System.Diagnostics;

namespace Skrypton.CSharpWriter.CodeTranslation
{
    [DebuggerDisplay("{Name}")]
    public sealed class CSharpName
    {
        public CSharpName(string name)
        {
            if (name == null) throw new ArgumentNullException(nameof(name));
            if (name.Length == 0) throw new ArgumentException("Blank name specified", nameof(name));
            for (int ix = 0; ix < name.Length; ix++)// PERFORMANCE:(no linq) if (name.Any(c => char.IsWhiteSpace(c)))
            {
                if (char.IsWhiteSpace(name, ix))
                    throw new ArgumentException("Specified name contains Whitespace - invalid", nameof(name));
            }

            Name = name;
        }

        /// <summary>
        /// This will never be null, blank or contain any Whitespace
        /// </summary>
        public string Name { get; private set; }
    }
}
