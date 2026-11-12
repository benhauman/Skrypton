using System;
using System.Collections.Generic;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The host service that the built-in mail components send their messages through (CDO.Message - issue #149). A host registers its own
    /// implementation (IRuntimeHost.TryGetRuntimeHostService) to redirect, restrict, log or switch off the sending of mail - if none is registered,
    /// messages are sent over SMTP with System.Net.Mail.SmtpClient.
    /// </summary>
    public interface IHostMailHostService
    {
        /// <summary>
        /// Send the message - smtp: the SMTP settings of the script (CDO.Configuration), null if the script has none (the host decides how to send)
        /// </summary>
        void Send(HostMailMessage message, HostSmtpSettings? smtp);
    }

    /// <summary>
    /// A mail message: the addresses as the script gives them ("Name &lt;address&gt;" or "address", several separated by "," or ";")
    /// </summary>
    public sealed class HostMailMessage
    {
        public HostMailMessage(string from, IReadOnlyList<string> to, IReadOnlyList<string> cc, IReadOnlyList<string> bcc, string subject, string textBody, string? htmlBody)
            : this(from, to, cc, bcc, subject, textBody, htmlBody, Array.Empty<HostMailAttachment>())
        {
        }

        public HostMailMessage(string from, IReadOnlyList<string> to, IReadOnlyList<string> cc, IReadOnlyList<string> bcc, string subject, string textBody, string? htmlBody, IReadOnlyList<HostMailAttachment> attachments)
        {
            Attachments = attachments ?? throw new ArgumentNullException(nameof(attachments));
            From = from ?? throw new ArgumentNullException(nameof(from));
            To = to ?? throw new ArgumentNullException(nameof(to));
            Cc = cc ?? throw new ArgumentNullException(nameof(cc));
            Bcc = bcc ?? throw new ArgumentNullException(nameof(bcc));
            Subject = subject ?? "";
            TextBody = textBody ?? "";
            HtmlBody = htmlBody;
        }

        public string From { get; }
        public IReadOnlyList<string> To { get; }
        public IReadOnlyList<string> Cc { get; }
        public IReadOnlyList<string> Bcc { get; }
        public string Subject { get; }

        /// <summary>
        /// The plain text body - with an HTML body, the text that CDO makes of it (unless the script set its own)
        /// </summary>
        public string TextBody { get; }

        /// <summary>
        /// Null if the message has no HTML body
        /// </summary>
        public string? HtmlBody { get; }

        /// <summary>
        /// The files attached to the message (CDO.Message.AddAttachment - issue #175)
        /// </summary>
        public IReadOnlyList<HostMailAttachment> Attachments { get; }
    }

    /// <summary>
    /// A file attached to a mail message: its name, media type (eg. text/plain) and content
    /// </summary>
    public sealed class HostMailAttachment
    {
        public HostMailAttachment(string fileName, string mediaType, byte[] content)
        {
            FileName = fileName ?? throw new ArgumentNullException(nameof(fileName));
            MediaType = mediaType ?? throw new ArgumentNullException(nameof(mediaType));
            Content = content ?? throw new ArgumentNullException(nameof(content));
        }

        public string FileName { get; }
        public string MediaType { get; }

#pragma warning disable CA1819 // The content is handed over as is (a copy per access would be wasteful)
        public byte[] Content { get; }
#pragma warning restore CA1819
    }

    /// <summary>
    /// The SMTP settings of a message (the CDO configuration fields smtpserver, smtpserverport, smtpauthenticate, sendusername, sendpassword,
    /// smtpusessl, smtpconnectiontimeout)
    /// </summary>
    public sealed class HostSmtpSettings
    {
        public HostSmtpSettings(string server, int port, string? userName, string? password, bool useSsl, TimeSpan timeout)
        {
            if (string.IsNullOrWhiteSpace(server))
                throw new ArgumentException("An SMTP server is required", nameof(server));
            Server = server;
            Port = port;
            UserName = userName;
            Password = password;
            UseSsl = useSsl;
            Timeout = timeout;
        }

        public string Server { get; }
        public int Port { get; }

        /// <summary>
        /// Null without authentication (smtpauthenticate 0)
        /// </summary>
        public string? UserName { get; }
        public string? Password { get; }
        public bool UseSsl { get; }
        public TimeSpan Timeout { get; }
    }
}
