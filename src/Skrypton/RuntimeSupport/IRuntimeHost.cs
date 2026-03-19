namespace Skrypton.RuntimeSupport.Implementations
{
    public interface IRuntimeHost
    {
        TService? TryGetRuntimeHostService<TService>() where TService : class;
    }
}