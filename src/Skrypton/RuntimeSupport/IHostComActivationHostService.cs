namespace Skrypton.RuntimeSupport.Implementations
{
    /// <summary>
    /// The last resort of CreateObject (issue #103): consulted only when neither RegisterObjectCreateFactory, the host's
    /// IHostObjectFactoryHostService nor a built-in object provides the ProgID - eg. to create registered COM classes (Skrypton.Windows:
    /// WindowsHostComActivationHostService, opt-in and limited to the ProgIDs the host allows). Null if it does not create the ProgID, then
    /// CreateObject raises error 429 (ActiveX component can't create object).
    /// </summary>
    public interface IHostComActivationHostService
    {
        object? TryCreateObject(string progId);
    }
}
