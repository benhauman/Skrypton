using System;

namespace Skrypton.RuntimeSupport.Implementations
{
    public interface IRuntimeHost
    {
        object? TryGetRuntimeHostService(Type serviceType);
    }
}