using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;
using Skrypton.RuntimeSupport.Components;

namespace Skrypton.Windows
{
    /// <summary>
    /// WshShell.Environment(type) as on Windows (issue #107): PROCESS is the environment of the process, SYSTEM, USER and VOLATILE are the
    /// variables in the registry as they are stored - not expanded ("%USERPROFILE%\AppData\Local\Temp"), as WshEnvironment returns them
    /// </summary>
    public sealed class WindowsHostEnvironmentVariablesHostService : DefaultHostEnvironmentVariablesHostService
    {
        public WindowsHostEnvironmentVariablesHostService()
        {
        }

        public override IReadOnlyList<KeyValuePair<string, string>> GetVariables(string type)
        {
            if (type == "PROCESS")
                return base.GetVariables(type);
            using RegistryKey? key = OpenKey(type, writable: false);
            if (key == null)
                return Array.Empty<KeyValuePair<string, string>>();
            return key.GetValueNames()
                .Where(name => name.Length > 0)
                .Select(name => new KeyValuePair<string, string>(name, ReadValue(key, name) ?? ""))
                .ToList();
        }

        public override string? GetVariable(string type, string name)
        {
            if (type == "PROCESS")
                return base.GetVariable(type, name);
            using RegistryKey? key = OpenKey(type, writable: false);
            return (key == null) ? null : ReadValue(key, name);
        }

        public override void SetVariable(string type, string name, string? value)
        {
            if (type == "PROCESS")
            {
                base.SetVariable(type, name, value);
                return;
            }
            using RegistryKey key = OpenKey(type, writable: true) ?? throw new InvalidOperationException($"WshEnvironment: the {type} environment is not available");
            if (value == null)
            {
                key.DeleteValue(name, throwOnMissingValue: false);
                return;
            }
            // a value that refers to other variables is stored as REG_EXPAND_SZ (as Windows does)
            RegistryValueKind kind = (key.GetValueNames().Contains(name, StringComparer.OrdinalIgnoreCase) ? key.GetValueKind(name) : RegistryValueKind.Unknown) == RegistryValueKind.ExpandString || value.Contains('%')
                ? RegistryValueKind.ExpandString
                : RegistryValueKind.String;
            key.SetValue(name, value, kind);
        }

        private static string? ReadValue(RegistryKey key, string name)
        {
            object? value = key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
            return (value == null) ? null : Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
        }

        private static RegistryKey? OpenKey(string type, bool writable)
        {
            switch (type)
            {
                case "SYSTEM": return Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Session Manager\Environment", writable);
                case "USER": return Registry.CurrentUser.OpenSubKey("Environment", writable);
                case "VOLATILE": return Registry.CurrentUser.OpenSubKey("Volatile Environment", writable);
                default: return null;
            }
        }
    }
}
