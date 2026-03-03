using System;
using System.Collections.Generic;
using System.Linq;

namespace Skrypton.RuntimeSupport
{
    public static class IBuildCallArgumentProvidersExtensions // public : used by generated code
    {
        /// <summary>
        /// TODO
        /// This should return a reference to itself to enable chaining when building up argument sets
        /// </summary>
        public static IBuildCallArgumentProviders RefIfArray(this IBuildCallArgumentProviders source, object target, params IBuildCallArgumentProviders[] argumentProviderBuilders)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (target == null)
                throw new ArgumentNullException(nameof(target));
            if (argumentProviderBuilders == null)
                throw new ArgumentNullException(nameof(argumentProviderBuilders));

            IProvideCallArguments[] argumentProviders = argumentProviderBuilders.Select(b => (b == null) ? throw new ArgumentException("Null reference encountered in argumentProviderBuilders set") : b.GetArgs()).ToArray();
            if (argumentProviders.Any(p => p == null))
                throw new ArgumentException("Null reference encountered in argumentProviderBuilders set");

            return source.RefIfArray(target, argumentProviders);
        }
    }
}
