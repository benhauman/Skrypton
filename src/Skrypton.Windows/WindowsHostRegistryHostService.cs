using System;
using Microsoft.Win32;
using Skrypton.RuntimeSupport.Components;

namespace Skrypton.Windows
{
    /// <summary>
    /// WshShell.RegRead on the registry of Windows (issue #122): the values as stored (REG_EXPAND_SZ not expanded), as Windows Script Host reads
    /// them
    /// </summary>
    public sealed class WindowsHostRegistryHostService : IHostRegistryHostService
    {
        public HostRegistryValue? ReadValue(string root, string subKey, string valueName, out bool keyExists)
        {
            RegistryKey baseKey = root switch
            {
                "HKEY_CURRENT_USER" => Registry.CurrentUser,
                "HKEY_LOCAL_MACHINE" => Registry.LocalMachine,
                "HKEY_CLASSES_ROOT" => Registry.ClassesRoot,
                "HKEY_USERS" => Registry.Users,
                "HKEY_CURRENT_CONFIG" => Registry.CurrentConfig,
                _ => throw new ArgumentException($"Unknown registry root '{root}'", nameof(root)),
            };
            using RegistryKey? key = (subKey.Length == 0) ? baseKey : baseKey.OpenSubKey(subKey, writable: false);
            keyExists = key != null;
            if (key == null)
                return null;
            object? data = key.GetValue(valueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
            if (data == null)
                return null;
            HostRegistryValueKind kind = key.GetValueKind(valueName) switch
            {
                RegistryValueKind.String => HostRegistryValueKind.Sz,
                RegistryValueKind.ExpandString => HostRegistryValueKind.ExpandSz,
                RegistryValueKind.DWord => HostRegistryValueKind.DWord,
                RegistryValueKind.QWord => HostRegistryValueKind.QWord,
                RegistryValueKind.MultiString => HostRegistryValueKind.MultiSz,
                RegistryValueKind.Binary => HostRegistryValueKind.Binary,
                _ => HostRegistryValueKind.Other,
            };
            return new HostRegistryValue(kind, data);
        }
    }
}
