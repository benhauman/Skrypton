using Skrypton.LegacyParser.CodeBlocks.Basic;
using System;
using System.Linq;

namespace Skrypton.CSharpWriter.CodeTranslation.Extensions
{
    internal static class PropertyBlockExtensions
    {
        public static bool IsIndexedProperty(this PropertyBlock source)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            if ((source.PropType == PropertyBlock.PropertyType.Get) && source.Parameters.Any())
                return true;

            return source.Parameters.Count() > 1;
        }
    }
}
