using System;
using System.Globalization;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// Creates instances of Skrypton's built-in COM object replacements (which are internal to Skrypton) the way a script does: through a compat layer
    /// with CreateObject, which falls back to the built-in dictionary when no factory is registered (see issue #50)
    /// </summary>
    internal static class TestBuiltInComponents
    {
        private sealed class NullLogger : IRuntimeLogger
        {
            public void LogException(Exception exception)
            {
            }
        }

        private static readonly DefaultRuntimeFunctionalityProvider _compatLayer = CreateCompatLayer();

        private static DefaultRuntimeFunctionalityProvider CreateCompatLayer()
        {
            var compatLayer = (DefaultRuntimeFunctionalityProvider)DefaultRuntimeSupportClassFactory
                .Create(new TestRuntimeHost(new TestHostServices()), new NullLogger(), CultureInfo.InvariantCulture)
                .Get();
            return compatLayer;
        }

        /// <summary>
        /// The compat layer that the dictionaries are created by (it can be used to call members on them, eg. CALLm1v2(null, d, "Add", k, v))
        /// </summary>
        public static DefaultRuntimeFunctionalityProvider CompatLayer => _compatLayer;

        public static object Create() => _compatLayer.CREATEOBJECT("Scripting.Dictionary");

        public static object Create(string progId) => _compatLayer.CREATEOBJECT(progId);

        /// <summary>
        /// Create a built-in object through a compat layer whose host provides the specified service (eg. an offline IHostHttpClientHostService)
        /// </summary>
        public static object CreateWithHostService<TService>(string progId, TService service) where TService : class
        {
            var hostServices = new TestHostServices().RegisterHostService<TService>(() => service);
            var compatLayer = (DefaultRuntimeFunctionalityProvider)DefaultRuntimeSupportClassFactory
                .Create(new TestRuntimeHost(hostServices), new NullLogger(), CultureInfo.InvariantCulture)
                .Get();
            return compatLayer.CREATEOBJECT(progId);
        }
    }
}
