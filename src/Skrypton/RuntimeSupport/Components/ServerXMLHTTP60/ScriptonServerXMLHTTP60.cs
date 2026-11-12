using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The built-in Msxml2.ServerXMLHTTP.6.0 - also for Msxml2.XMLHTTP (and its versions; TypeName is IServerXMLHTTPRequest2 for them all with
    /// Windows VBScript - issue #102). Requests are sent through the host's IHostHttpClientHostService (or with HttpClient if the host has not
    /// registered one), synchronously - a request opened with async = True is also completed by send (readyState is then 4).
    /// </summary>
    [SourceClassName("IServerXMLHTTPRequest2")] // TypeName(CreateObject("Msxml2.ServerXMLHTTP.6.0")) (as with Windows VBScript)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonServerXMLHTTP60 : ScriptonIReflectOnClrType
    {
        public const string ProgId = "Msxml2.ServerXMLHTTP.6.0";

        /// <summary>
        /// The ProgIDs of this built-in (issue #102)
        /// </summary>
        public static readonly IReadOnlyCollection<string> ProgIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ProgId, "Msxml2.ServerXMLHTTP", "Msxml2.ServerXMLHTTP.3.0", "Msxml2.XMLHTTP", "Msxml2.XMLHTTP.3.0", "Msxml2.XMLHTTP.6.0", "Microsoft.XMLHTTP",
        };

        private readonly IHostHttpClientHostService _httpClient;
        private readonly List<KeyValuePair<string, string>> _requestHeaders = new List<KeyValuePair<string, string>>();
        private readonly Dictionary<int, object?> _options = new Dictionary<int, object?> { { 2, 0 } }; // SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS
        private string? _method;
        private Uri? _url;
        private string? _user;
        private string? _password;
        private TimeSpan? _timeout;
        private HostHttpResponse? _response;
        private readonly IRuntimeHost _runtimeHost;

        public ScriptonServerXMLHTTP60(IRuntimeHost runtimeHost) : base(typeof(ScriptonServerXMLHTTP60))
        {
            if (runtimeHost == null)
                throw new ArgumentNullException(nameof(runtimeHost));
            _httpClient = runtimeHost.TryGetRuntimeHostService<IHostHttpClientHostService>() ?? DefaultHostHttpClientHostService.Instance;
            _runtimeHost = runtimeHost;
        }

        /// <summary>
        /// 0 (unsent), 1 (opened) or 4 (completed)
        /// </summary>
        [DispId(0x0000000E)]
        public int readyState { get; private set; }

        [DispId(0x00000001)]
        public void open(string bstrMethod, string bstrUrl, object? varAsync = null, object? bstrUser = null, object? bstrPassword = null)
        {
            if (string.IsNullOrWhiteSpace(bstrMethod))
                throw new ArgumentException("ServerXMLHTTP.open: a method (eg. GET) is required", nameof(bstrMethod));
            if (string.IsNullOrWhiteSpace(bstrUrl))
                throw new ArgumentException("ServerXMLHTTP.open: a URL is required", nameof(bstrUrl));

            _method = bstrMethod;
            _url = new Uri(bstrUrl, UriKind.Absolute);
            _user = ToOptionalString(bstrUser);
            _password = ToOptionalString(bstrPassword);
            _requestHeaders.Clear();
            _response = null;
            readyState = 1;
        }

        [DispId(0x00000002)]
        public void setRequestHeader(string bstrHeader, string bstrValue)
        {
            if (readyState != 1)
                throw new InvalidOperationException("ServerXMLHTTP.setRequestHeader: open must be called first (and send not yet)");
            if (string.IsNullOrWhiteSpace(bstrHeader))
                throw new ArgumentException("ServerXMLHTTP.setRequestHeader: a header name is required", nameof(bstrHeader));
            _requestHeaders.Add(new KeyValuePair<string, string>(bstrHeader, bstrValue ?? ""));
        }

        [DispId(0x00000005)]
        public void send() => send(null);

        [DispId(0x00000005)]
        public void send(object? varBody)
        {
            if ((readyState != 1) || (_method == null) || (_url == null))
                throw new InvalidOperationException("ServerXMLHTTP.send: open must be called first");

            byte[]? body = ToRequestBody(varBody is ScriptDispatchWrapper wrapper ? wrapper.WrappedObject : varBody);
            _response = _httpClient.Send(new HostHttpRequest(_method, _url, _user, _password, _requestHeaders.ToArray(), body, _timeout));
            readyState = 4;
        }

        [DispId(0x00000008)]
        public int status => Response.Status;

        [DispId(0x00000009)]
        public string statusText => Response.StatusText;

        /// <summary>
        /// The response body as a Byte array
        /// </summary>
        [DispId(0x0000000F)]
        public byte[] responseBody => Response.Body;

        /// <summary>
        /// The response body decoded with the charset of its Content-Type (UTF-8 if there is none)
        /// </summary>
        [DispId(0x00000007)]
        public string responseText => GetEncoding().GetString(Response.Body);

        /// <summary>
        /// The response body parsed as XML (an empty document if it is not XML, as with MSXML)
        /// </summary>
        [DispId(0x00000006)]
        public object responseXML
        {
            get
            {
                var document = new ScriptonMsxml2DOMDocument(_runtimeHost);
                document.setProperty("SelectionLanguage", MsxmlSelection.XPath); // an MSXML 6 document: XPath is its selection language
                document.load(Response.Body);
                return document;
            }
        }

        [DispId(0x00000003)]
        public string getResponseHeader(string bstrHeader)
        {
            return string.Join(", ", Response.Headers.Where(h => h.Key.Equals(bstrHeader, StringComparison.OrdinalIgnoreCase)).Select(h => h.Value));
        }

        [DispId(0x00000004)]
        public string getAllResponseHeaders()
        {
            return string.Concat(Response.Headers.Select(h => h.Key + ": " + h.Value + "\r\n"));
        }

        /// <summary>
        /// Requests are sent synchronously, so there is nothing in flight to abort - this resets the object to its unsent state
        /// </summary>
        [DispId(0x0000000A)]
        public void abort()
        {
            _response = null;
            _requestHeaders.Clear();
            readyState = 0;
        }

        /// <summary>
        /// The timeouts in milliseconds - the request uses the longest of the connect, send and receive timeouts as its overall timeout
        /// (zero or a negative value means no timeout, as with MSXML)
        /// </summary>
        [DispId(0x0000000B)]
        public void setTimeouts(int resolveTimeout, int connectTimeout, int sendTimeout, int receiveTimeout)
        {
            int milliseconds = new[] { resolveTimeout, connectTimeout, sendTimeout, receiveTimeout }.Max();
            _timeout = (milliseconds > 0) ? TimeSpan.FromMilliseconds(milliseconds) : System.Threading.Timeout.InfiniteTimeSpan;
        }

        /// <summary>
        /// Wait for an asynchronous request - requests are sent synchronously, so this is always True
        /// </summary>
        [DispId(0x00000010)]
#pragma warning disable CA1822 // Bound by name from scripts, so it must be an instance member
        public bool waitForResponse(object? timeoutInSeconds = null) => true;
#pragma warning restore CA1822

        /// <summary>
        /// The options are stored and returned by getOption (eg. SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS = 2), but certificate handling is up
        /// to the IHostHttpClientHostService
        /// </summary>
        [DispId(0x0000000C)]
        public void setOption(int option, object? value)
        {
            _options[option] = value;
        }

        [DispId(0x0000000D)]
        public object? getOption(int option)
        {
            return _options.TryGetValue(option, out object? value) ? value : null;
        }

        /// <summary>
        /// Has the script set the request header (before send)
        /// </summary>
        internal bool HasRequestHeader(string header) => _requestHeaders.Any(h => h.Key.Equals(header, StringComparison.OrdinalIgnoreCase));

        /// <summary>
        /// The credentials of WinHttpRequest.SetCredentials (sent with the request, as the user and password of open)
        /// </summary>
        internal void SetCredentials(string user, string password)
        {
            _user = user;
            _password = password;
        }

        /// <summary>
        /// The URL of the response (after redirects), or the opened URL before send
        /// </summary>
        internal Uri? ResponseUrlOrOpenedUrl => _response?.Url ?? _url;

        private HostHttpResponse Response => _response ?? throw new InvalidOperationException("ServerXMLHTTP: the response is only available after send has completed");

        private Encoding GetEncoding()
        {
            string contentType = getResponseHeader("Content-Type");
            int charsetIndex = contentType.IndexOf("charset=", StringComparison.OrdinalIgnoreCase);
            if (charsetIndex >= 0)
            {
                string charset = contentType.Substring(charsetIndex + "charset=".Length).Split(';')[0].Trim().Trim('"');
                try
                {
                    return Encoding.GetEncoding(charset);
                }
                catch (ArgumentException)
                {
                    // Unknown charset - fall back to UTF-8
                }
            }
            return Encoding.UTF8;
        }

        private static string? ToOptionalString(object? value)
        {
            if ((value == null) || (value is ScriptDispatchWrapper) || (value == DBNull.Value) || (value is System.Reflection.Missing))
                return null;
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static byte[]? ToRequestBody(object? body)
        {
            switch (body)
            {
                case null:
                case System.Reflection.Missing _:
                    return null;
                case byte[] bytes:
                    return bytes;
                case string text:
                    return Encoding.UTF8.GetBytes(text);
                case ScriptonMsxml2DOMDocument document:
                    return Encoding.UTF8.GetBytes(document.xml);
                default:
                    return Encoding.UTF8.GetBytes(Convert.ToString(body, CultureInfo.InvariantCulture) ?? "");
            }
        }
    }
}
