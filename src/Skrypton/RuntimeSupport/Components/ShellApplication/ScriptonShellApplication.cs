using System;
using System.Globalization;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The built-in Shell.Application (IShellDispatch). ShellExecute goes through the host's IHostProcessControlHostService - opening files and
    /// URLs from a script is a decision for the host, so it raises an error if the host has not registered one.
    /// </summary>
    [SourceClassName("IShellDispatch6")] // TypeName(CreateObject("Shell.Application")) (as with Windows VBScript)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonShellApplication : ScriptonIReflectOnClrType
    {
        public const string ProgId = "Shell.Application";

        private readonly IHostProcessControlHostService? _processControl;

        public ScriptonShellApplication(IRuntimeHost runtimeHost) : base(typeof(ScriptonShellApplication))
        {
            if (runtimeHost == null)
                throw new ArgumentNullException(nameof(runtimeHost));
            _processControl = runtimeHost.TryGetRuntimeHostService<IHostProcessControlHostService>();
        }

        /// <summary>
        /// ShellExecute(File, [Arguments], [Directory], [Operation], [Show]) - fire and forget, as with Windows
        /// </summary>
        [DispId(0x60040000)]
        public void ShellExecute(
            [In, MarshalAs(UnmanagedType.BStr)] string File,
            [In, Optional, DefaultParameterValue("")] object? vArgs,
            [In, Optional, DefaultParameterValue("")] object? vDir,
            [In, Optional, DefaultParameterValue("")] object? vOperation,
            [In, Optional, DefaultParameterValue(1)] object? vShow)
        {
            if (string.IsNullOrWhiteSpace(File))
                throw new ArgumentException("Shell.ShellExecute: a file is required", nameof(File));

            IHostProcessControlHostService processControl = _processControl ?? throw new InvalidOperationException(
                $"Shell.ShellExecute: opening files and URLs requires the host to register an {nameof(IHostProcessControlHostService)}"
            );
            processControl.ShellExecute(File, ToText(vArgs), ToText(vDir), ToText(vOperation), IsMissing(vShow) ? 1 : Convert.ToInt32(vShow, CultureInfo.InvariantCulture));
        }

        private static bool IsMissing(object? value) => (value == null) || (value is System.Reflection.Missing);

        private static string ToText(object? value) => IsMissing(value) ? "" : (Convert.ToString(value, CultureInfo.InvariantCulture) ?? "");
    }
}
