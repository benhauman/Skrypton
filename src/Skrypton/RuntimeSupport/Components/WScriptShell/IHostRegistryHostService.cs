namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The host service that WshShell.RegRead reads the registry through (issue #122) - reading the registry is a decision for the host, so
    /// RegRead raises an error if the host has not registered one (Skrypton.Windows: WindowsHostRegistryHostService, on Microsoft.Win32.Registry)
    /// </summary>
    public interface IHostRegistryHostService
    {
        /// <summary>
        /// The value (null if the key has no such value; keyExists is False if the key does not exist). root: HKEY_CURRENT_USER,
        /// HKEY_LOCAL_MACHINE, HKEY_CLASSES_ROOT, HKEY_USERS or HKEY_CURRENT_CONFIG; subKey: the path below it ("" for the root); valueName: ""
        /// for the default value of the key
        /// </summary>
        HostRegistryValue? ReadValue(string root, string subKey, string valueName, out bool keyExists);
    }

    public enum HostRegistryValueKind
    {
        Sz,
        ExpandSz,
        DWord,
        QWord,
        MultiSz,
        Binary,
        Other,
    }

    /// <summary>
    /// A registry value as stored: an Sz (REG_SZ) / ExpandSz (REG_EXPAND_SZ, not expanded) is a string, a DWord an int, a QWord a long, a MultiSz a string[], a
    /// Binary a byte[]
    /// </summary>
    public sealed class HostRegistryValue
    {
        public HostRegistryValue(HostRegistryValueKind kind, object data)
        {
            Kind = kind;
            Data = data ?? throw new System.ArgumentNullException(nameof(data));
        }

        public HostRegistryValueKind Kind { get; }

        public object Data { get; }
    }
}
