using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// Sends the messages of the built-in CDO.Message over SMTP with System.Net.Mail.SmtpClient (used when the host has not registered an
    /// IHostMailHostService of its own - issue #149): an HTML body is sent as multipart/alternative with its text, as CDO does
    /// </summary>
    internal sealed class DefaultHostMailHostService : IHostMailHostService
    {
        public static readonly DefaultHostMailHostService Instance = new DefaultHostMailHostService();

        private DefaultHostMailHostService()
        {
        }

        public void Send(HostMailMessage message, HostSmtpSettings? smtp)
        {
            if (message == null)
                throw new ArgumentNullException(nameof(message));
            if (smtp == null)
                throw new InvalidOperationException("The message has no SMTP settings - register an IHostMailHostService to send it otherwise");

            using var mail = new MailMessage { From = new MailAddress(message.From), Subject = message.Subject, SubjectEncoding = Encoding.UTF8 };
            AddAll(mail.To, message.To);
            AddAll(mail.CC, message.Cc);
            AddAll(mail.Bcc, message.Bcc);
            if (message.HtmlBody == null)
            {
                mail.Body = message.TextBody;
                mail.BodyEncoding = Encoding.UTF8;
                mail.BodyTransferEncoding = TransferEncodingOf(message.TextBody);
            }
            else
            {
#pragma warning disable CA2000 // The views are owned (and disposed) by the message
                mail.AlternateViews.Add(View(message.TextBody, MediaTypeNames.Text.Plain));
                mail.AlternateViews.Add(View(message.HtmlBody, MediaTypeNames.Text.Html));
#pragma warning restore CA2000
            }

            // the attached files - the message becomes multipart/mixed, each file a base64 part with its file name (issue #175)
            foreach (HostMailAttachment file in message.Attachments)
            {
#pragma warning disable CA2000 // The attachment (and its stream) is owned (and disposed) by the message
                var attachment = new Attachment(new System.IO.MemoryStream(file.Content), file.FileName, file.MediaType);
#pragma warning restore CA2000
                attachment.ContentDisposition!.FileName = file.FileName;
                mail.Attachments.Add(attachment);
            }

            using var client = new SmtpClient(smtp.Server, smtp.Port)
            {
                EnableSsl = smtp.UseSsl,
                Timeout = (int)Math.Min(int.MaxValue, smtp.Timeout.TotalMilliseconds),
                DeliveryMethod = SmtpDeliveryMethod.Network,
            };
            if (smtp.UserName != null)
                client.Credentials = new NetworkCredential(smtp.UserName, smtp.Password ?? "");
            client.Send(mail);
        }

        private static AlternateView View(string text, string mediaType)
        {
            AlternateView view = AlternateView.CreateAlternateViewFromString(text, Encoding.UTF8, mediaType);
            view.TransferEncoding = TransferEncodingOf(text);
            return view;
        }

        /// <summary>
        /// 7bit for ASCII text (as CDO sends it - readable), quoted-printable otherwise
        /// </summary>
        private static TransferEncoding TransferEncodingOf(string text) => text.All(c => c < 128) ? TransferEncoding.SevenBit : TransferEncoding.QuotedPrintable;

        private static void AddAll(MailAddressCollection target, IReadOnlyList<string> addresses)
        {
            foreach (string address in addresses)
                target.Add(new MailAddress(address));
        }
    }
}
