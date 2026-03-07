namespace Skrypton.RuntimeSupport.Implementations
{
    public interface IHostObjectFactoryHostService
    {
        System.Func<IRuntimeHost, object>? TryGetObjectFactoryRegistration(string progId);
    }
}