using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The host service that WshShell.Environment(type) reads and changes environment variables through (issue #107). The type is "PROCESS",
    /// "SYSTEM", "USER" or "VOLATILE" (upper case). If the host has not registered one: the process environment (SYSTEM and USER through .NET),
    /// unless the host has registered its own IHostEnvironmentHostService - then the variables are read through its ExpandEnvironmentStrings
    /// and cannot be changed, so that a sandboxed environment stays sandboxed.
    /// </summary>
    public interface IHostEnvironmentVariablesHostService
    {
        /// <summary>
        /// The variables of the type (names and values as stored - Windows does not expand the SYSTEM and USER values for WshEnvironment)
        /// </summary>
        IReadOnlyList<KeyValuePair<string, string>> GetVariables(string type);

        /// <summary>
        /// The value of the variable, null if there is none (the name is compared case-insensitively)
        /// </summary>
        string? GetVariable(string type, string name);

        /// <summary>
        /// Set the variable - a null value removes it
        /// </summary>
        void SetVariable(string type, string name, string? value);
    }

    /// <summary>
    /// The environment of the process (EnvironmentVariableTarget.Process / Machine / User - .NET expands the SYSTEM and USER values; the
    /// Windows host service reads them as stored). VOLATILE is the user's volatile environment, which only the Windows host service reads.
    /// </summary>
    public class DefaultHostEnvironmentVariablesHostService : IHostEnvironmentVariablesHostService
    {
        public static readonly DefaultHostEnvironmentVariablesHostService Instance = new DefaultHostEnvironmentVariablesHostService();

        protected DefaultHostEnvironmentVariablesHostService()
        {
        }

        public virtual IReadOnlyList<KeyValuePair<string, string>> GetVariables(string type)
        {
            if (!TryGetTarget(type, out EnvironmentVariableTarget target))
                return Array.Empty<KeyValuePair<string, string>>();
            return Environment.GetEnvironmentVariables(target).Cast<DictionaryEntry>()
                .Select(e => new KeyValuePair<string, string>((string)e.Key, (string?)e.Value ?? ""))
                .Where(v => !v.Key.StartsWith("=", StringComparison.Ordinal)) // the hidden "=C:=C:\.." variables are not listed by WshEnvironment
                .ToList();
        }

        public virtual string? GetVariable(string type, string name)
        {
            return TryGetTarget(type, out EnvironmentVariableTarget target) ? Environment.GetEnvironmentVariable(name, target) : null;
        }

        public virtual void SetVariable(string type, string name, string? value)
        {
            if (!TryGetTarget(type, out EnvironmentVariableTarget target))
                throw new NotSupportedException($"WshEnvironment: the {type} environment is only available through the Windows host service");
            Environment.SetEnvironmentVariable(name, value, target);
        }

        private static bool TryGetTarget(string type, out EnvironmentVariableTarget target)
        {
            switch (type)
            {
                case "PROCESS": target = EnvironmentVariableTarget.Process; return true;
                case "SYSTEM": target = EnvironmentVariableTarget.Machine; return true;
                case "USER": target = EnvironmentVariableTarget.User; return true;
                default: target = EnvironmentVariableTarget.Process; return false;
            }
        }
    }

    /// <summary>
    /// The variables of a host's own IHostEnvironmentHostService (a sandbox): read through ExpandEnvironmentStrings("%name%"), not listed and
    /// not changed
    /// </summary>
    internal sealed class ExpandingEnvironmentVariablesHostService : IHostEnvironmentVariablesHostService
    {
        private readonly IHostEnvironmentHostService _environment;

        public ExpandingEnvironmentVariablesHostService(IHostEnvironmentHostService environment)
        {
            _environment = environment ?? throw new ArgumentNullException(nameof(environment));
        }

        public IReadOnlyList<KeyValuePair<string, string>> GetVariables(string type)
        {
            throw new NotSupportedException($"WshEnvironment: listing the variables requires the host to register an {nameof(IHostEnvironmentVariablesHostService)}");
        }

        public string? GetVariable(string type, string name)
        {
            if (type != "PROCESS")
                return null;
            string reference = "%" + name + "%";
            string value = _environment.ExpandEnvironmentStrings(reference);
            return (value == reference) ? null : value;
        }

        public void SetVariable(string type, string name, string? value)
        {
            throw new NotSupportedException($"WshEnvironment: changing the variables requires the host to register an {nameof(IHostEnvironmentVariablesHostService)}");
        }
    }
}
