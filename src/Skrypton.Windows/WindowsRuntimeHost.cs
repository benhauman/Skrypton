using System;
using System.Collections.Generic;
using Skrypton.RuntimeSupport.Components;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.Windows
{
    /// <summary>
    /// An IRuntimeHost with the host services of Windows - scripts behave as with cscript.exe: WScript.Shell starts processes, Shell.Application and
    /// InternetExplorer.Application open files and URLs, htmlfile uses the clipboard, MsgBox / InputBox show dialogs, ADODB connects through OLE
    /// DB providers and Outlook.Application stores, sends and shows items with Outlook. Every service can be replaced (Register).
    /// </summary>
    public sealed class WindowsRuntimeHost : IRuntimeHost
    {
        private readonly Dictionary<Type, object> _services = new Dictionary<Type, object>();

        /// <summary>
        /// A host without services (the built-in components use their defaults or raise errors) - see CreateDefault
        /// </summary>
        public WindowsRuntimeHost()
        {
        }

        /// <summary>
        /// A host with all Windows host services
        /// </summary>
        public static WindowsRuntimeHost CreateDefault()
        {
            return new WindowsRuntimeHost()
                .Register<IHostProcessControlHostService>(new WindowsHostProcessControlHostService())
                .Register<IHostClipboardHostService>(new WindowsHostClipboardHostService())
                .Register<IHostMessageBoxHostService>(new WindowsHostMessageBoxHostService())
                .Register<IHostInputBoxHostService>(new WindowsHostInputBoxHostService())
                .Register<IHostDatabaseConnectionFactoryHostService>(new WindowsHostOleDbConnectionFactoryHostService())
                .Register<IHostOutlookHostService>(new WindowsHostOutlookHostService())
                .Register<IHostEnvironmentVariablesHostService>(new WindowsHostEnvironmentVariablesHostService())
                .Register<IHostEventLogHostService>(new WindowsHostEventLogHostService());
        }

        /// <summary>
        /// Register (or replace) a host service
        /// </summary>
        public WindowsRuntimeHost Register<TService>(TService service) where TService : class
        {
            _services[typeof(TService)] = service ?? throw new ArgumentNullException(nameof(service));
            return this;
        }

        /// <summary>
        /// Remove a host service (the built-in components then use their default or raise an error)
        /// </summary>
        public WindowsRuntimeHost Unregister<TService>() where TService : class
        {
            _services.Remove(typeof(TService));
            return this;
        }

        public TService? TryGetRuntimeHostService<TService>() where TService : class
        {
            return _services.TryGetValue(typeof(TService), out object? service) ? (TService)service : null;
        }
    }
}
