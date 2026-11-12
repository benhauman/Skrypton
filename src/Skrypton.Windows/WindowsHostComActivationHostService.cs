using System;
using System.Collections.Generic;
using System.Linq;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.Windows
{
    /// <summary>
    /// CreateObject of registered COM classes (issue #103) - opt-in and limited to the ProgIDs the host allows: in-house components
    /// ("hlStartITIL.Global"), "WScript.Network", ... The COM class is created only when no factory and no built-in object provides the ProgID;
    /// the script calls it through IDispatch. A ProgID that is not allowed or not registered raises error 429 (ActiveX component can't create
    /// object), as with Windows VBScript.
    /// <code>
    /// WindowsRuntimeHost host = WindowsRuntimeHost.CreateDefault().RegisterComActivation("hlStartITIL.Global", "WScript.Network");
    /// </code>
    /// </summary>
    public sealed class WindowsHostComActivationHostService : IHostComActivationHostService
    {
        private readonly Func<string, bool> _isAllowed;

        /// <summary>
        /// Only the ProgIDs in the list (compared case-insensitively)
        /// </summary>
        public WindowsHostComActivationHostService(IEnumerable<string> allowedProgIds)
        {
            var allowed = new HashSet<string>((allowedProgIds ?? throw new ArgumentNullException(nameof(allowedProgIds))).Where(p => !string.IsNullOrWhiteSpace(p)), StringComparer.OrdinalIgnoreCase);
            _isAllowed = allowed.Contains;
        }

        /// <summary>
        /// The ProgIDs that the predicate allows (eg. p => p.StartsWith("hlStart", StringComparison.OrdinalIgnoreCase))
        /// </summary>
        public WindowsHostComActivationHostService(Func<string, bool> isAllowed)
        {
            _isAllowed = isAllowed ?? throw new ArgumentNullException(nameof(isAllowed));
        }

        public object? TryCreateObject(string progId)
        {
            if (string.IsNullOrWhiteSpace(progId) || !_isAllowed(progId))
                return null;
            Type? comType = Type.GetTypeFromProgID(progId, throwOnError: false);
            return (comType == null) ? null : Activator.CreateInstance(comType);
        }
    }
}
