using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The built-in CDO.Configuration (CDOSYS - issue #149): its Fields are the settings of the messages that use it
    /// (http://schemas.microsoft.com/cdo/configuration/sendusing, smtpserver, smtpserverport, smtpauthenticate, sendusername, sendpassword,
    /// smtpusessl, smtpconnectiontimeout, sendemailaddress); the names are not case-sensitive
    /// </summary>
    [SourceClassName("IConfiguration")] // TypeName(CreateObject("CDO.Configuration")) (as with Windows VBScript)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonCdoConfiguration : ScriptonIReflectOnClrType
    {
        public const string ProgId = "CDO.Configuration";

        internal const string FieldNamespace = "http://schemas.microsoft.com/cdo/configuration/";

        public ScriptonCdoConfiguration() : base(typeof(ScriptonCdoConfiguration))
        {
        }

        [DispId(0x00000001)]
        public ScriptonCdoFields Fields { get; } = new ScriptonCdoFields();

        internal string? GetText(string name)
        {
            object? value = Fields.ValueOf(FieldNamespace + name);
            return ((value == null) || (value is DBNull)) ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        internal int GetNumber(string name, int defaultValue)
        {
            string? text = GetText(name);
            return ((text != null) && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) ? (int)Math.Round(number) : defaultValue;
        }

        internal bool GetBoolean(string name)
        {
            object? value = Fields.ValueOf(FieldNamespace + name);
            return (value is bool b) ? b : ((value is string text) ? (text.Trim().Equals("True", StringComparison.OrdinalIgnoreCase) || (GetNumber(name, 0) != 0)) : (GetNumber(name, 0) != 0));
        }
    }

    /// <summary>
    /// The Fields of a CDO.Configuration: Fields(name) and Fields.Item(name) are the Field of the name (one with the value Empty if it has not been
    /// set), Fields(name) = value sets its value (through the default property of the Field - issue #147), Update keeps the values
    /// </summary>
    [SourceClassName("Fields")]
    [System.Reflection.DefaultMember("Item")] // Fields(name) is Fields.Item(name)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonCdoFields : ScriptonIReflectOnClrType
    {
        private readonly Dictionary<string, ScriptonCdoField> _fields = new Dictionary<string, ScriptonCdoField>(StringComparer.OrdinalIgnoreCase);

        public ScriptonCdoFields() : base(typeof(ScriptonCdoFields))
        {
            // The fields that a new CDO.Configuration has (as with Windows CDOSYS; its urn:schemas:calendar:timezoneid depends on the machine)
            Item(ScriptonCdoConfiguration.FieldNamespace + "languagecode").Value = "en-us";
            Item(ScriptonCdoConfiguration.FieldNamespace + "postusing").Value = 0;
            Item(ScriptonCdoConfiguration.FieldNamespace + "sendusing").Value = 0;
            Item(ScriptonCdoConfiguration.FieldNamespace + "usemessageresponsetext").Value = true;
        }

        [DispId(0)]
        [IsDefault]
        public ScriptonCdoField Item(object Index)
        {
            string name = Convert.ToString((Index is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : Index, CultureInfo.InvariantCulture) ?? "";
            if (!_fields.TryGetValue(name, out ScriptonCdoField? field))
            {
                field = new ScriptonCdoField(name);
                _fields.Add(name, field);
            }
            return field;
        }

        /// <summary>
        /// The configuration is used as it is - Update is accepted (CDO writes the values to the configuration source with it)
        /// </summary>
        [DispId(0x00000005)]
#pragma warning disable CA1822 // Bound by name from scripts, so it must be an instance member
        public void Update()
        {
        }
#pragma warning restore CA1822

        internal object? ValueOf(string name) => _fields.TryGetValue(name, out ScriptonCdoField? field) ? field.Value : null;
    }

    /// <summary>
    /// A field of a CDO.Configuration: its Name and its Value (the default property)
    /// </summary>
    [SourceClassName("Field")]
    [System.Reflection.DefaultMember("Value")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonCdoField : ScriptonIReflectOnClrType
    {
        private object? _value;

        public ScriptonCdoField(string name) : base(typeof(ScriptonCdoField))
        {
            Name = name;
        }

        [DispId(0x0000044C)]
        public string Name { get; }

        [DispId(0)]
        [IsDefault]
        public object? Value
        {
            get => _value;
            set => _value = (value is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : value;
        }
    }

    /// <summary>
    /// The built-in CDO.Message (CDOSYS - issue #149): Send hands the message to the host's IHostMailHostService (or sends it over SMTP with
    /// SmtpClient if the host has not registered one) with the SMTP settings of its Configuration - sendusing must be 2 (cdoSendUsingPort). The
    /// errors are the ones of CDOSYS (number, description and source "CDO.Message.1").
    /// </summary>
    [SourceClassName("IMessage")] // TypeName(CreateObject("CDO.Message")) (as with Windows VBScript)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonCdoMessage : ScriptonIReflectOnClrType
    {
        public const string ProgId = "CDO.Message";

        private const string ErrorSource = "CDO.Message.1";
        private const int SendUsingPort = 2; // cdoSendUsingPort
        private const int DefaultTimeoutSeconds = 30; // smtpconnectiontimeout

        private readonly IHostMailHostService _mail;
        private readonly IHostFileSystemHostService _fileSystem;
        private ScriptonCdoConfiguration _configuration = new ScriptonCdoConfiguration();
        private string _htmlBody = "";
        private string _textBody = "";

        public ScriptonCdoMessage(IRuntimeHost runtimeHost) : base(typeof(ScriptonCdoMessage))
        {
            if (runtimeHost == null)
                throw new ArgumentNullException(nameof(runtimeHost));
            _mail = runtimeHost.TryGetRuntimeHostService<IHostMailHostService>() ?? DefaultHostMailHostService.Instance;
            _fileSystem = runtimeHost.TryGetRuntimeHostService<IHostFileSystemHostService>() ?? DefaultHostFileSystemHostService.Instance;
        }

        /// <summary>
        /// Set Message.Configuration = configuration - a message has a configuration of its own until then
        /// </summary>
        [DispId(0x00000067)]
        public object Configuration
        {
            get => _configuration;
            set => _configuration = (((value is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : value) as ScriptonCdoConfiguration)
                ?? throw new CustomException(13, ErrorSource, "Type mismatch");
        }

        [DispId(0x00000068)]
        public string From { get; set; } = "";

        [DispId(0x00000069)]
        public string To { get; set; } = "";

        [DispId(0x0000006A)]
        public string CC { get; set; } = "";

        [DispId(0x0000006B)]
        public string BCC { get; set; } = "";

        [DispId(0x0000006C)]
        public string Subject { get; set; } = "";

        /// <summary>
        /// The plain text body - setting HTMLBody makes it the text of the HTML (as CDO does), setting it keeps the HTML body
        /// </summary>
        [DispId(0x0000006D)]
        public string TextBody
        {
            get => _textBody;
            set => _textBody = value ?? "";
        }

        [DispId(0x0000006E)]
        public string HTMLBody
        {
            get => _htmlBody;
            set
            {
                _htmlBody = value ?? "";
                _textBody = CdoHtmlText.ToText(_htmlBody);
            }
        }

        [DispId(0x0000006F)]
        public void Send()
        {
            if (_configuration.GetNumber("sendusing", 0) != SendUsingPort)
                throw CdoError(unchecked((int)0x80040220), "The \"SendUsing\" configuration value is invalid.");
            string server = _configuration.GetText("smtpserver") ?? "";
            if (server.Trim().Length == 0)
                throw CdoError(unchecked((int)0x8004020A), "The SMTP server name is required, and was not found in the configuration source.");
            if (From.Trim().Length == 0)
            {
                string sendEmailAddress = _configuration.GetText("sendemailaddress") ?? "";
                if (sendEmailAddress.Trim().Length == 0)
                    throw CdoError(unchecked((int)0x8004020D), "At least one of the From or Sender fields is required, and neither was found.");
                From = "\"" + sendEmailAddress.Trim() + "\" <" + sendEmailAddress.Trim() + ">"; // as CDO sets it
            }
            IReadOnlyList<string> to = Addresses(To), cc = Addresses(CC), bcc = Addresses(BCC);
            if ((to.Count + cc.Count + bcc.Count) == 0)
                throw CdoError(unchecked((int)0x8004020C), "At least one recipient is required, but none were found.");

            bool authenticate = _configuration.GetNumber("smtpauthenticate", 0) != 0;
            var smtp = new HostSmtpSettings(
                server.Trim(),
                _configuration.GetNumber("smtpserverport", 25),
                authenticate ? (_configuration.GetText("sendusername") ?? "") : null,
                authenticate ? (_configuration.GetText("sendpassword") ?? "") : null,
                _configuration.GetBoolean("smtpusessl"),
                TimeSpan.FromSeconds(_configuration.GetNumber("smtpconnectiontimeout", DefaultTimeoutSeconds)));
            var message = new HostMailMessage(From, to, cc, bcc, Subject, TextBody, (_htmlBody.Length == 0) ? null : _htmlBody, Attachments.Files());
            try
            {
                _mail.Send(message, smtp);
            }
            catch (System.Net.Mail.SmtpException e) when ((e.StatusCode == System.Net.Mail.SmtpStatusCode.GeneralFailure) || (e.InnerException is System.Net.Sockets.SocketException) || (e.InnerException is System.IO.IOException))
            {
                throw CdoError(unchecked((int)0x80040213), "The transport failed to connect to the server.");
            }
        }

        private static CustomException CdoError(int number, string description) => new CustomException(number, ErrorSource, description + "\r\n"); // CDO's descriptions end with a line break

        /// <summary>
        /// The attachments of the message (issue #175)
        /// </summary>
        [DispId(0x00000070)]
        public ScriptonCdoBodyParts Attachments { get; } = new ScriptonCdoBodyParts();

        /// <summary>
        /// AddAttachment(URL, [UserName], [Password]) - attaches the file (a path or a file: URL) and returns its body part (issue #175); a missing
        /// file is error 80070002, as with CDOSYS
        /// </summary>
        [DispId(0x00000071)]
        public ScriptonCdoBodyPart AddAttachment(string URL, [Optional, DefaultParameterValue(null)] object? UserName, [Optional, DefaultParameterValue(null)] object? Password)
        {
            string path = URL ?? "";
            if (path.StartsWith("file:", StringComparison.OrdinalIgnoreCase) && Uri.TryCreate(path, UriKind.Absolute, out Uri? uri))
                path = uri.LocalPath;
            if (!_fileSystem.FileExists(path))
                throw CdoError(unchecked((int)0x80070002), "The system cannot find the file specified.");
            byte[] content;
            using (System.IO.StreamReader reader = _fileSystem.OpenTextFileRead(path))
            using (var buffer = new System.IO.MemoryStream())
            {
                reader.BaseStream.CopyTo(buffer);
                content = buffer.ToArray();
            }
            var part = new ScriptonCdoBodyPart(System.IO.Path.GetFileName(path), content);
            Attachments.Add(part);
            return part;
        }

        /// <summary>
        /// The addresses of To, CC or BCC - separated by "," or ";" (not within quotes or angle brackets)
        /// </summary>
        internal static IReadOnlyList<string> Addresses(string value)
        {
            var addresses = new List<string>();
            var current = new StringBuilder();
            bool quoted = false, bracketed = false;
            foreach (char c in value ?? "")
            {
                if ((c == '"') && !bracketed)
                    quoted = !quoted;
                else if ((c == '<') && !quoted)
                    bracketed = true;
                else if ((c == '>') && !quoted)
                    bracketed = false;
                if (((c == ',') || (c == ';')) && !quoted && !bracketed)
                {
                    Add(addresses, current);
                    continue;
                }
                current.Append(c);
            }
            Add(addresses, current);
            return addresses;
        }

        private static void Add(List<string> addresses, StringBuilder current)
        {
            string address = current.ToString().Trim();
            if (address.Length > 0)
                addresses.Add(address);
            current.Clear();
        }
    }

    /// <summary>
    /// The attachments of a CDO.Message (IBodyParts - issue #175): Count, Item(index) from 1, For Each, Delete(index or part), DeleteAll; an index
    /// out of range is error 8002000B (Invalid index.), as with CDOSYS
    /// </summary>
    [SourceClassName("IBodyParts")]
    [System.Reflection.DefaultMember("Item")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
#pragma warning disable CA1010 // Generic IEnumerable<T> - For Each of a script enumerates the parts as variants
    internal sealed class ScriptonCdoBodyParts : ScriptonIReflectOnClrType, System.Collections.IEnumerable
#pragma warning restore CA1010
    {
        private readonly List<ScriptonCdoBodyPart> _parts = new List<ScriptonCdoBodyPart>();

        public ScriptonCdoBodyParts() : base(typeof(ScriptonCdoBodyParts))
        {
        }

        [DispId(1)]
        public int Count => _parts.Count;

        [DispId(0)]
        [IsDefault]
        public ScriptonCdoBodyPart Item(object Index) => _parts[IndexOf(Index)];

        [DispId(2)]
        public void Delete(object Varbodypart)
        {
            object? value = (Varbodypart is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : Varbodypart;
            _parts.RemoveAt((value is ScriptonCdoBodyPart part) ? _parts.IndexOf(part) is int index && (index >= 0) ? index : throw InvalidIndex() : IndexOf(value));
        }

        [DispId(3)]
        public void DeleteAll() => _parts.Clear();

        [DispId(-4)] // DISPID_NEWENUM
#pragma warning disable CA1707 // The name is the COM member name
        public System.Collections.IEnumerator _NewEnum() => GetEnumerator();
#pragma warning restore CA1707

        public System.Collections.IEnumerator GetEnumerator() => _parts.ToArray().GetEnumerator();

        internal void Add(ScriptonCdoBodyPart part) => _parts.Add(part);

        internal IReadOnlyList<HostMailAttachment> Files() => _parts.Select(p => new HostMailAttachment(p.FileName, p.ContentMediaType, p.Content)).ToList();

        private int IndexOf(object? index)
        {
            double number;
            try
            {
                number = Convert.ToDouble((index is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : index, CultureInfo.InvariantCulture);
            }
            catch (Exception e) when ((e is FormatException) || (e is InvalidCastException))
            {
                throw InvalidIndex();
            }
            int position = (int)Math.Round(number);
            if ((position < 1) || (position > _parts.Count))
                throw InvalidIndex();
            return position - 1;
        }

        private static CustomException InvalidIndex() => new CustomException(unchecked((int)0x8002000B), "CDO.Message.1", "Invalid index.\r\n");
    }

    /// <summary>
    /// An attachment of a CDO.Message (IBodyPart - issue #175): FileName, ContentMediaType (from the extension) and ContentTransferEncoding
    /// </summary>
    [SourceClassName("IBodyPart")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonCdoBodyPart : ScriptonIReflectOnClrType
    {
        // the media types that Windows (the registry) gives the common extensions - others are application/octet-stream
        private static readonly Dictionary<string, string> MediaTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { ".txt", "text/plain" }, { ".log", "text/plain" }, { ".xml", "text/xml" }, { ".htm", "text/html" }, { ".html", "text/html" },
            { ".csv", "application/vnd.ms-excel" }, { ".pdf", "application/pdf" }, { ".zip", "application/x-zip-compressed" },
            { ".png", "image/png" }, { ".jpg", "image/jpeg" }, { ".jpeg", "image/jpeg" }, { ".gif", "image/gif" }, { ".bmp", "image/bmp" },
            { ".doc", "application/msword" }, { ".xls", "application/vnd.ms-excel" },
            { ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
            { ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        };

        public ScriptonCdoBodyPart(string fileName, byte[] content) : base(typeof(ScriptonCdoBodyPart))
        {
            FileName = fileName;
            Content = content;
            ContentMediaType = MediaTypes.TryGetValue(System.IO.Path.GetExtension(fileName), out string? mediaType) ? mediaType : "application/octet-stream";
        }

        [DispId(0x00000068)]
        public string FileName { get; }

        [DispId(0x00000066)]
        public string ContentMediaType { get; set; }

        /// <summary>
        /// 7bit, as CDOSYS reports it before it sends the message (the default mail service sends the files as base64)
        /// </summary>
        [DispId(0x00000067)]
        public string ContentTransferEncoding { get; set; } = "7bit";

#pragma warning disable CA1819 // The content is handed over as is
        internal byte[] Content { get; }
#pragma warning restore CA1819
    }

    /// <summary>
    /// The text of an HTML body, as CDO makes it (MimeOLE): the text of the elements with the white space collapsed, a line break for br, div and
    /// a table row, an empty line after a paragraph, a heading and a list, "*" and a tab before a list item, a tab after a table cell; the content
    /// of head, script and style is left out and the character references are decoded (a non-breaking space is a space)
    /// </summary>
    internal static class CdoHtmlText
    {
        private static readonly Regex Skipped = new Regex(@"<(head|script|style|title)\b[^>]*>.*?</\1\s*>", RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant);
        private static readonly Regex Token = new Regex(@"<(/?)([a-zA-Z][a-zA-Z0-9]*)[^>]*>|<!--.*?-->|[^<]+|<", RegexOptions.Singleline | RegexOptions.CultureInvariant);
        private static readonly Regex WhiteSpace = new Regex(@"\s+", RegexOptions.CultureInvariant);

        public static string ToText(string html)
        {
            var text = new StringBuilder();
            foreach (Match token in Token.Matches(Skipped.Replace(html ?? "", "")))
            {
                if (!token.Groups[2].Success)
                {
                    if (token.Value.StartsWith("<!--", StringComparison.Ordinal))
                        continue;
                    string content = WhiteSpace.Replace(WebUtility.HtmlDecode(token.Value).Replace(' ', ' ').Replace("\r", " ").Replace("\n", " "), " ");
                    if (AtLineStart(text))
                        content = content.TrimStart(' ');
                    text.Append(content);
                    continue;
                }
                bool end = token.Groups[1].Value == "/";
                switch (token.Groups[2].Value.ToUpperInvariant())
                {
                    case "BR":
                        text.Append("\r\n");
                        break;
                    case "P":
                        if (end)
                            text.Append("\r\n\r\n");
                        break;
                    case "DIV":
                    case "TR":
                        if (end)
                            text.Append("\r\n");
                        break;
                    case "H1":
                    case "H2":
                    case "H3":
                    case "H4":
                    case "H5":
                    case "H6":
                        text.Append(end ? "\r\n\r\n" : "\r\n");
                        break;
                    case "UL":
                    case "OL":
                        if (end)
                            text.Append("\r\n");
                        else if (!AtLineStart(text))
                            text.Append("\r\n\r\n");
                        break;
                    case "LI":
                        text.Append(end ? "\r\n" : "*\t");
                        break;
                    case "TD":
                    case "TH":
                        if (end)
                            text.Append('\t');
                        break;
                }
            }
            return text.ToString();
        }

        private static bool AtLineStart(StringBuilder text) => (text.Length == 0) || (text[text.Length - 1] == '\n') || (text[text.Length - 1] == '\t');
    }
}
