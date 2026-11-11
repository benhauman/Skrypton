using System;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    internal static class IHostProcessControlHostServiceExtensions
    {
        public static T GetRequiredService<T>(this IServiceProvider serviceProvider) where T : class
        {
            if (serviceProvider == null) throw new ArgumentNullException(nameof(serviceProvider));
            return ((T)serviceProvider.GetService(typeof(T))) ?? throw new InvalidOperationException($"Service '{typeof(T).FullName}' not registered.");
        }
    }
}
