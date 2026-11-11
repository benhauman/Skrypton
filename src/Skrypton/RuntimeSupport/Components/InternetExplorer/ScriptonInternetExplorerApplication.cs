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
    /// The built-in InternetExplorer.Application (IWebBrowser2). There is no browser window: a page that is navigated to is opened with the
    /// associated application (the default browser) through the host's IHostProcessControlHostService as soon as the browser is visible - an
    /// invisible browser only records the location (as a hidden Internet Explorer loads it without showing it). The window properties (size,
    /// bars, ..) are kept, with the defaults of Internet Explorer.
    /// </summary>
    [DefaultMember("Name")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("IWebBrowser2")] // TypeName(CreateObject("InternetExplorer.Application")) (as with Windows VBScript)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonInternetExplorerApplication : ScriptonIReflectOnClrType
    {
        public const string ProgId = "InternetExplorer.Application";

        // READYSTATE_UNINITIALIZED / READYSTATE_COMPLETE
        private const int ReadyStateUninitialized = 0;
        private const int ReadyStateComplete = 4;

        private readonly IRuntimeHost _runtimeHost;
        private bool _visible;
        private string _locationUrl = "";
        private bool _locationShown = true;
        private bool _quit;

        public ScriptonInternetExplorerApplication(IRuntimeHost runtimeHost) : base(typeof(ScriptonInternetExplorerApplication))
        {
            _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
        }

        /// <summary>
        /// The default member (VarType 8, "" &amp; ie is "Internet Explorer")
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public string Name => "Internet Explorer";

        [DispId(400)] public string FullName => "IEXPLORE.EXE";

        [DispId(402)]
        public bool Visible
        {
            get => _visible;
            set
            {
                EnsureNotQuit();
                _visible = value;
                ShowLocation();
            }
        }

        [DispId(212)] public bool Busy => false;

        [DispId(-525)] public int ReadyState => (_locationUrl.Length == 0) ? ReadyStateUninitialized : ReadyStateComplete;

        [DispId(211)] public string LocationURL => _locationUrl;

        [DispId(210)] public string LocationName => _locationUrl;

        [DispId(551)] public bool Silent { get; set; }
        [DispId(550)] public bool Offline { get; set; }
        [DispId(552)] public bool RegisterAsBrowser { get; set; }
        [DispId(553)] public bool RegisterAsDropTarget { get; set; } = true;
        [DispId(554)] public bool TheaterMode { get; set; }
        [DispId(555)] public bool AddressBar { get; set; } = true;
        [DispId(556)] public bool Resizable { get; set; } = true;
        [DispId(406)] public bool MenuBar { get; set; } = true;
        [DispId(407)] public bool FullScreen { get; set; }
        [DispId(403)] public bool StatusBar { get; set; }
        [DispId(404)] public string StatusText { get; set; } = "";
        [DispId(405)] public int ToolBar { get; set; } = 1;
        [DispId(206)] public int Left { get; set; }
        [DispId(207)] public int Top { get; set; }
        [DispId(208)] public int Width { get; set; } = 1024;
        [DispId(209)] public int Height { get; set; } = 768;
        [DispId(-515)] public int HWND => 0;

        /// <summary>
        /// Navigate(URL, [Flags], [TargetFrameName], [PostData], [Headers]) - posting data or headers is not supported
        /// </summary>
        [DispId(104)]
        public void Navigate(
            [In, MarshalAs(UnmanagedType.BStr)] string URL,
            [In, Optional, DefaultParameterValue(null)] object? Flags,
            [In, Optional, DefaultParameterValue(null)] object? TargetFrameName,
            [In, Optional, DefaultParameterValue(null)] object? PostData,
            [In, Optional, DefaultParameterValue(null)] object? Headers)
        {
            NavigateCore("Navigate", URL, PostData, Headers);
        }

        /// <summary>
        /// Navigate2(URL, [Flags], [TargetFrameName], [PostData], [Headers])
        /// </summary>
        [DispId(500)]
        public void Navigate2(
            [In, MarshalAs(UnmanagedType.Struct)] object URL,
            [In, Optional, DefaultParameterValue(null)] object? Flags,
            [In, Optional, DefaultParameterValue(null)] object? TargetFrameName,
            [In, Optional, DefaultParameterValue(null)] object? PostData,
            [In, Optional, DefaultParameterValue(null)] object? Headers)
        {
            NavigateCore("Navigate2", Convert.ToString(URL, CultureInfo.InvariantCulture) ?? "", PostData, Headers);
        }

        [DispId(-550)]
        public void Refresh()
        {
            EnsureNotQuit();
        }

        [DispId(106)]
        public void Stop()
        {
            EnsureNotQuit();
        }

        [DispId(300)]
        public void Quit()
        {
            _quit = true;
        }

        private void NavigateCore(string member, string url, object? postData, object? headers)
        {
            EnsureNotQuit();
            if (string.IsNullOrWhiteSpace(url))
                throw new ArgumentException($"InternetExplorer.{member}: a URL is required", nameof(url));
            if (!IsMissing(postData) || !IsMissing(headers))
                throw new NotSupportedException($"InternetExplorer.{member}: posting data or sending headers is not supported (the page is opened in the default browser)");
            _locationUrl = url;
            _locationShown = false;
            ShowLocation();
        }

        private void ShowLocation()
        {
            if (!_visible || _locationShown || (_locationUrl.Length == 0))
                return;
            _locationShown = true;
            IHostProcessControlHostService processControl = _runtimeHost.TryGetRuntimeHostService<IHostProcessControlHostService>() ?? throw new InvalidOperationException(
                $"InternetExplorer: showing a page requires the host to register an {nameof(IHostProcessControlHostService)}"
            );
            processControl.ShellExecute(_locationUrl, "", "", "open", 1);
        }

        private void EnsureNotQuit()
        {
            // As with a real browser that was quit: the object is disconnected from its server
            if (_quit)
                throw new InvalidOperationException("InternetExplorer: the browser has been quit (the object invoked has disconnected from its clients)");
        }

        private static bool IsMissing(object? value) => (value == null) || (value is System.Reflection.Missing) || (value is DBNull);
    }
}
