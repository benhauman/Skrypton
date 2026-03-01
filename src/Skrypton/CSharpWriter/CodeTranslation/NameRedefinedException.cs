using System;
using System.Globalization;
using System.Runtime.Serialization;
using Skrypton.LegacyParser.Tokens.Basic;

namespace Skrypton.CSharpWriter.CodeTranslation
{
    public sealed class NameRedefinedException : Exception
    {
        public NameRedefinedException(NameToken name) : base(GetMessage(name))
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
        }

        private NameRedefinedException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        /// This will never be null
        /// </summary>
        public NameToken Name { get; private set; }

        private static string GetMessage(NameToken name)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));

            return string.Format(CultureInfo.InvariantCulture,
                "Name redefined at line {0}: \"{1}\"",
                name.LineIndex + 1,
                name.Content
            );
        }
    }
}
