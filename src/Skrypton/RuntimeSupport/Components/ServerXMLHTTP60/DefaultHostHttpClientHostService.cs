using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// Sends the requests of the built-in Msxml2.ServerXMLHTTP.6.0 with System.Net.Http.HttpClient (used when the host has not registered an
    /// IHostHttpClientHostService of its own). Requests are sent synchronously, as ServerXMLHTTP does with async = False.
    /// </summary>
    internal sealed class DefaultHostHttpClientHostService : IHostHttpClientHostService
    {
        public static readonly DefaultHostHttpClientHostService Instance = new DefaultHostHttpClientHostService();

        private static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(100);

        private static readonly HttpClient SharedClient = CreateClient(credentials: null);

        private DefaultHostHttpClientHostService()
        {
        }

        private static HttpClient CreateClient(ICredentials? credentials)
        {
#pragma warning disable CA2000 // The handler is owned (and disposed) by the HttpClient (disposeHandler: true)
            var handler = new HttpClientHandler { CheckCertificateRevocationList = true };
#pragma warning restore CA2000
            if (credentials != null)
                handler.Credentials = credentials;
            return new HttpClient(handler, disposeHandler: true) { Timeout = Timeout.InfiniteTimeSpan };
        }

        public HostHttpResponse Send(HostHttpRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            if (string.IsNullOrEmpty(request.User))
                return Send(SharedClient, request);

            // Credentials need a handler of their own (the shared client is used for anonymous requests)
            using HttpClient clientWithCredentials = CreateClient(new NetworkCredential(request.User, request.Password ?? ""));
            return Send(clientWithCredentials, request);
        }

        private static HostHttpResponse Send(HttpClient client, HostHttpRequest request)
        {
            using var message = new HttpRequestMessage(new HttpMethod(request.Method), request.Url);
            if (request.Body != null)
                message.Content = new ByteArrayContent(request.Body);
            foreach (KeyValuePair<string, string> header in request.Headers)
            {
                if (message.Headers.TryAddWithoutValidation(header.Key, header.Value))
                    continue;

                // Content headers (eg. Content-Type) belong to the content
                message.Content ??= new ByteArrayContent(Array.Empty<byte>());
                if (header.Key.Equals("Content-Type", StringComparison.OrdinalIgnoreCase))
                    message.Content.Headers.ContentType = MediaTypeHeaderValue.Parse(header.Value);
                else
                    message.Content.Headers.TryAddWithoutValidation(header.Key, header.Value);
            }

            using var cancellation = new CancellationTokenSource(request.Timeout ?? DefaultTimeout);
            // The request is completed (synchronously) before the message and the cancellation source are disposed
#pragma warning disable CA2025 // Ensure tasks using 'IDisposable' instances complete before the instances are disposed
            using HttpResponseMessage response = client.SendAsync(message, cancellation.Token).ConfigureAwait(false).GetAwaiter().GetResult();
#pragma warning restore CA2025
            byte[] body = response.Content.ReadAsByteArrayAsync().ConfigureAwait(false).GetAwaiter().GetResult();

            List<KeyValuePair<string, string>> headers = response.Headers
                .Concat(response.Content.Headers)
                .SelectMany(h => h.Value.Select(v => new KeyValuePair<string, string>(h.Key, v)))
                .ToList();
            return new HostHttpResponse((int)response.StatusCode, response.ReasonPhrase ?? "", headers, body, response.RequestMessage?.RequestUri); // the URL after redirects
        }
    }
}
