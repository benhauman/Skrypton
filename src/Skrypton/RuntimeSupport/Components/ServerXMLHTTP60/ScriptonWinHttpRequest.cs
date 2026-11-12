using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The built-in WinHttp.WinHttpRequest.5.1 (issue #102), on the same transport as Msxml2.ServerXMLHTTP.6.0: requests are sent through the
    /// host's IHostHttpClientHostService (or with HttpClient if the host has not registered one), synchronously. Option(n) keeps the options;
    /// Option(1) (WinHttpRequestOption_URL) is the URL of the response (after redirects)
    /// </summary>
    [SourceClassName("WinHttpRequest")] // TypeName(CreateObject("WinHttp.WinHttpRequest.5.1")) (as with Windows VBScript)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonWinHttpRequest : ScriptonIReflectOnClrType
    {
        public const string ProgId = "WinHttp.WinHttpRequest.5.1";

        private const int OptionUrl = 1; // WinHttpRequestOption_URL
        private const int OptionUserAgentString = 0; // WinHttpRequestOption_UserAgentString

        /// <summary>
        /// The User-Agent that WinHttpRequest sends unless the script sets one (as received from Windows' WinHttp.WinHttpRequest.5.1)
        /// </summary>
        internal const string DefaultUserAgent = "Mozilla/4.0 (compatible; Win32; WinHttp.WinHttpRequest.5)";

        private readonly ScriptonServerXMLHTTP60 _request;
        private readonly Dictionary<int, object?> _options = new Dictionary<int, object?>();

        public ScriptonWinHttpRequest(IRuntimeHost runtimeHost) : base(typeof(ScriptonWinHttpRequest))
        {
            _request = new ScriptonServerXMLHTTP60(runtimeHost);
        }

        [DispId(0x00000006)]
        public void Open(string Method, string Url, object? Async = null)
        {
            _request.open(Method, Url, Async);
        }

        [DispId(0x00000005)]
        public void SetRequestHeader(string Header, string Value) => _request.setRequestHeader(Header, Value);

        [DispId(0x00000003)]
        public void SetCredentials(string UserName, string Password, int Flags) => _request.SetCredentials(UserName, Password);

        /// <summary>
        /// The proxy settings are up to the IHostHttpClientHostService - they are accepted and ignored
        /// </summary>
        [DispId(0x0000000D)]
#pragma warning disable CA1822 // Bound by name from scripts, so it must be an instance member
        public void SetProxy(int ProxySetting, object? ProxyServer = null, object? BypassList = null)
        {
        }
#pragma warning restore CA1822

        [DispId(0x0000000C)]
        public void Send() => Send(null);

        /// <summary>
        /// Send the request - with WinHttp's defaults: its User-Agent (or Option(0)) and, for a text body, "text/plain; Charset=UTF-8"
        /// </summary>
        [DispId(0x0000000C)]
        public void Send(object? Body)
        {
            object? body = (Body is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : Body;
            if (!_request.HasRequestHeader("User-Agent"))
                _request.setRequestHeader("User-Agent", (_options.TryGetValue(OptionUserAgentString, out object? userAgent) && (userAgent is string text)) ? text : DefaultUserAgent);
            if ((body is string) && !_request.HasRequestHeader("Content-Type"))
                _request.setRequestHeader("Content-Type", "text/plain; Charset=UTF-8");
            _request.send(body);
        }

        [DispId(0x00000007)]
        public int Status => _request.status;

        [DispId(0x00000008)]
        public string StatusText => _request.statusText;

        [DispId(0x00000009)]
        public string ResponseText => _request.responseText;

        [DispId(0x0000000A)]
        public byte[] ResponseBody => _request.responseBody;

        [DispId(0x00000001)]
        public string GetResponseHeader(string Header) => _request.getResponseHeader(Header);

        [DispId(0x00000002)]
        public string GetAllResponseHeaders() => _request.getAllResponseHeaders();

        [DispId(0x00000010)]
        public void SetTimeouts(int ResolveTimeout, int ConnectTimeout, int SendTimeout, int ReceiveTimeout) => _request.setTimeouts(ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout);

        /// <summary>
        /// Requests are sent synchronously, so this is always True
        /// </summary>
        [DispId(0x0000000F)]
        public bool WaitForResponse(object? Timeout = null) => _request.waitForResponse(Timeout);

        [DispId(0x0000000E)]
        public void Abort() => _request.abort();

        /// <summary>
        /// Option(n) - the options set by the script; Option(1) (WinHttpRequestOption_URL) is the URL of the response after redirects (the
        /// opened URL before Send)
        /// </summary>
        [DispId(0x00000006 + 0x100)]
        [IndexerName("Option")]
        public object? this[int Option]
        {
            get
            {
                if (Option == OptionUrl)
                    return _request.ResponseUrlOrOpenedUrl?.AbsoluteUri ?? "";
                if ((Option == OptionUserAgentString) && !_options.ContainsKey(Option))
                    return DefaultUserAgent;
                return _options.TryGetValue(Option, out object? value) ? value : null;
            }
            set
            {
                _options[Option] = (value is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : value;
            }
        }
    }
}
