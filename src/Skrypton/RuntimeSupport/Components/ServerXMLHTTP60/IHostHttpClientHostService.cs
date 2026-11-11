using System;
using System.Collections.Generic;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The host service that the built-in Msxml2.ServerXMLHTTP.6.0 sends its requests through (like IHostFileSystemHostService for the file system).
    /// A host registers its own implementation (IRuntimeHost.TryGetRuntimeHostService) to control, proxy, log or fake HTTP traffic - if none is
    /// registered, requests are sent with System.Net.Http.HttpClient.
    /// </summary>
    public interface IHostHttpClientHostService
    {
        HostHttpResponse Send(HostHttpRequest request);
    }

    public sealed class HostHttpRequest
    {
        public HostHttpRequest(string method, Uri url, string? user, string? password, IReadOnlyList<KeyValuePair<string, string>> headers, byte[]? body, TimeSpan? timeout)
        {
            if (string.IsNullOrWhiteSpace(method))
                throw new ArgumentException("A request method (eg. GET) is required", nameof(method));
            if (url == null)
                throw new ArgumentNullException(nameof(url), "A request URL is required");

            Method = method;
            Url = url;
            User = user;
            Password = password;
            Headers = headers ?? throw new ArgumentNullException(nameof(headers));
            Body = body;
            Timeout = timeout;
        }

        public string Method { get; }
        public Uri Url { get; }
        public string? User { get; }
        public string? Password { get; }
        public IReadOnlyList<KeyValuePair<string, string>> Headers { get; }

        /// <summary>
        /// Null if the request has no body
        /// </summary>
#pragma warning disable CA1819 // The body is handed over as is (a copy per access would be wasteful)
        public byte[]? Body { get; }
#pragma warning restore CA1819

        /// <summary>
        /// Null for the service's default timeout
        /// </summary>
        public TimeSpan? Timeout { get; }
    }

    public sealed class HostHttpResponse
    {
        public HostHttpResponse(int status, string statusText, IReadOnlyList<KeyValuePair<string, string>> headers, byte[] body)
        {
            Status = status;
            StatusText = statusText ?? "";
            Headers = headers ?? throw new ArgumentNullException(nameof(headers));
            Body = body ?? throw new ArgumentNullException(nameof(body));
        }

        public int Status { get; }
        public string StatusText { get; }
        public IReadOnlyList<KeyValuePair<string, string>> Headers { get; }
#pragma warning disable CA1819 // The body is handed over as is (a copy per access would be wasteful)
        public byte[] Body { get; }
#pragma warning restore CA1819
    }
}
