// The members are bound by name from scripts, so they must be instance members
#pragma warning disable CA1822
using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The built-in htmlfile (an MSHTML HTMLDocument without a page) - scripts create it for parentWindow.clipboardData, which goes through the
    /// host's IHostClipboardHostService (there is no default, so the clipboard members raise an error if the host has not registered one)
    /// </summary>
    [DefaultMember("toString")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("HTMLDocument")] // TypeName(CreateObject("htmlfile")) (as with Windows VBScript)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonHtmlFile : ScriptonIReflectOnClrType
    {
        public const string ProgId = "htmlfile";

        private readonly ScriptonHtmlWindow _parentWindow;

        public ScriptonHtmlFile(IRuntimeHost runtimeHost) : base(typeof(ScriptonHtmlFile))
        {
            _parentWindow = new ScriptonHtmlWindow(runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost)));
        }

        /// <summary>
        /// The default member ("" &amp; CreateObject("htmlfile") is "[object]", VarType 8)
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public string toString() => "[object]";

        [DispId(1034)]
        public object parentWindow => _parentWindow;
    }

    /// <summary>
    /// document.parentWindow (VarType 9: it has no default value)
    /// </summary>
    [SourceClassName("HTMLWindow2")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonHtmlWindow : ScriptonIReflectOnClrType
    {
        private readonly ScriptonHtmlDataTransfer _clipboardData;

        public ScriptonHtmlWindow(IRuntimeHost runtimeHost) : base(typeof(ScriptonHtmlWindow))
        {
            _clipboardData = new ScriptonHtmlDataTransfer(runtimeHost);
        }

        [DispId(1159)]
        public object clipboardData => _clipboardData;
    }

    /// <summary>
    /// window.clipboardData (IHTMLDataTransfer) - setData / getData / clearData with the formats "Text" and "URL"
    /// </summary>
    [DefaultMember("toString")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("IHTMLDataTransfer")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonHtmlDataTransfer : ScriptonIReflectOnClrType
    {
        private readonly IRuntimeHost _runtimeHost;

        public ScriptonHtmlDataTransfer(IRuntimeHost runtimeHost) : base(typeof(ScriptonHtmlDataTransfer))
        {
            _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
        }

        /// <summary>
        /// The default member (VarType(clipboardData) is 8)
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public string toString() => "[object]";

        [DispId(1001)]
        public bool setData([In, MarshalAs(UnmanagedType.BStr)] string format, [In, MarshalAs(UnmanagedType.Struct)] object? data)
        {
            return Clipboard().SetText(Format(format), ToText(data));
        }

        /// <summary>
        /// getData(format) - the text on the clipboard ("" if there is none)
        /// </summary>
        [DispId(1002)]
        public string getData([In, MarshalAs(UnmanagedType.BStr)] string format)
        {
            return Clipboard().GetText(Format(format)) ?? "";
        }

        /// <summary>
        /// clearData([format]) - all formats if there is none
        /// </summary>
        [DispId(1003)]
        public bool clearData([In, Optional, DefaultParameterValue(null)] object? format)
        {
            return Clipboard().Clear(((format == null) || (format is System.Reflection.Missing)) ? "" : Format(ToText(format)));
        }

        private IHostClipboardHostService Clipboard()
        {
            return _runtimeHost.TryGetRuntimeHostService<IHostClipboardHostService>() ?? throw new InvalidOperationException(
                $"clipboardData: using the clipboard requires the host to register an {nameof(IHostClipboardHostService)}"
            );
        }

        private static string Format(string format)
        {
            if (string.Equals(format, "Text", StringComparison.OrdinalIgnoreCase))
                return "Text";
            if (string.Equals(format, "URL", StringComparison.OrdinalIgnoreCase))
                return "URL";
            throw new ArgumentException($"clipboardData: the format '{format}' is not supported (Text and URL are)", nameof(format));
        }

        private static string ToText(object? value) => ((value == null) || (value is DBNull)) ? "" : (Convert.ToString(value, CultureInfo.InvariantCulture) ?? "");
    }
}
