using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The built-in MAPI.Session (CDO 1.21 - issue #150) for scripts that send mail through a MAPI profile: CDO 1.21 is not installed on current
    /// Windows, so the members that these scripts use (Logon, Outbox.Messages.Add, Recipients.Add / Resolve, Subject, Text, Send, Logoff) send
    /// through the host's IHostMailHostService - there is no SMTP (or MAPI) fallback, the host must register one. The mailbox of the ProfileInfo
    /// of Logon ("server" &amp; vbLf &amp; "mailbox") is the sender.
    /// </summary>
    [SourceClassName("Session")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonMapiSession : ScriptonIReflectOnClrType
    {
        public const string ProgId = "MAPI.Session";

        internal const string ErrorSource = "Collaboration Data Objects";

        private readonly IHostMailHostService? _mail;
        private ScriptonMapiFolder? _outbox;

        public ScriptonMapiSession(IRuntimeHost runtimeHost) : base(typeof(ScriptonMapiSession))
        {
            if (runtimeHost == null)
                throw new ArgumentNullException(nameof(runtimeHost));
            _mail = runtimeHost.TryGetRuntimeHostService<IHostMailHostService>();
        }

        /// <summary>
        /// Logon([ProfileName], [ProfilePassword], [ShowDialog], [NewSession], [ParentWindow], [NoMail], [ProfileInfo]) - ProfileInfo is the server
        /// and the mailbox separated by vbLf, the mailbox is the sender of the messages
        /// </summary>
        [DispId(0x00000077)]
        public void Logon(
            [Optional, DefaultParameterValue(null)] object? ProfileName,
            [Optional, DefaultParameterValue(null)] object? ProfilePassword,
            [Optional, DefaultParameterValue(null)] object? ShowDialog,
            [Optional, DefaultParameterValue(null)] object? NewSession,
            [Optional, DefaultParameterValue(null)] object? ParentWindow,
            [Optional, DefaultParameterValue(null)] object? NoMail,
            [Optional, DefaultParameterValue(null)] object? ProfileInfo)
        {
            if (_mail == null)
                throw Error(unchecked((int)0x80040111), "MAPI_E_LOGON_FAILED", "no mail host service (IHostMailHostService) is registered to send the mail of MAPI.Session");
            string profile = Text(ProfileInfo) ?? Text(ProfileName) ?? "";
            int separator = profile.IndexOf('\n');
            string mailbox = ((separator < 0) ? profile : profile.Substring(separator + 1)).Trim();
            _outbox = new ScriptonMapiFolder(_mail, mailbox);
        }

        [DispId(0x00000078)]
        public void Logoff() => _outbox = null;

        [DispId(0x00000079)]
        public ScriptonMapiFolder Outbox => _outbox ?? throw Error(unchecked((int)0x80040605), "MAPI_E_NOT_INITIALIZED", "Logon first");

        /// <summary>
        /// The error of CDO 1.21: "[Collaboration Data Objects - [MAPI_E_..(hresult)]]", with what Skrypton knows about it
        /// </summary>
        internal static CustomException Error(int hresult, string name, string? detail)
        {
            string description = "[" + ErrorSource + " - [" + name + "(" + hresult.ToString("X8", CultureInfo.InvariantCulture) + ")]]";
            return new CustomException(hresult, ErrorSource, (detail == null) ? description : (description + ": " + detail));
        }

        internal static string? Text(object? value)
        {
            object? unwrapped = (value is ScriptDispatchWrapper wrapper) ? wrapper.WrappedObject : value;
            return ((unwrapped == null) || (unwrapped is DBNull) || (unwrapped == Type.Missing)) ? null : Convert.ToString(unwrapped, CultureInfo.InvariantCulture);
        }
    }

    [SourceClassName("Folder")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonMapiFolder : ScriptonIReflectOnClrType
    {
        public ScriptonMapiFolder(IHostMailHostService mail, string mailbox) : base(typeof(ScriptonMapiFolder))
        {
            Messages = new ScriptonMapiMessages(mail, mailbox);
        }

        [DispId(0x00000082)]
        public ScriptonMapiMessages Messages { get; }
    }

    [SourceClassName("Messages")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonMapiMessages : ScriptonIReflectOnClrType
    {
        private readonly IHostMailHostService _mail;
        private readonly string _mailbox;

        public ScriptonMapiMessages(IHostMailHostService mail, string mailbox) : base(typeof(ScriptonMapiMessages))
        {
            _mail = mail;
            _mailbox = mailbox;
        }

        /// <summary>
        /// Add([Subject], [Text], [Type], [Importance]) - a new message
        /// </summary>
        [DispId(0x0000008C)]
        public ScriptonMapiMessage Add(
            [Optional, DefaultParameterValue(null)] object? Subject,
            [Optional, DefaultParameterValue(null)] object? Text,
            [Optional, DefaultParameterValue(null)] object? Type,
            [Optional, DefaultParameterValue(null)] object? Importance)
        {
            return new ScriptonMapiMessage(_mail, _mailbox) { Subject = ScriptonMapiSession.Text(Subject) ?? "", Text = ScriptonMapiSession.Text(Text) ?? "" };
        }
    }

    [SourceClassName("Message")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonMapiMessage : ScriptonIReflectOnClrType
    {
        private readonly IHostMailHostService _mail;
        private readonly string _mailbox;

        public ScriptonMapiMessage(IHostMailHostService mail, string mailbox) : base(typeof(ScriptonMapiMessage))
        {
            _mail = mail;
            _mailbox = mailbox;
        }

        [DispId(0x00000096)]
        public string Subject { get; set; } = "";

        [DispId(0x00000097)]
        public string Text { get; set; } = "";

        [DispId(0x00000098)]
        public ScriptonMapiRecipients Recipients { get; } = new ScriptonMapiRecipients();

        /// <summary>
        /// Send([SaveCopy], [ShowDialog], [ParentWindow]) - the message goes to the mail host service (without SMTP settings)
        /// </summary>
        [DispId(0x00000099)]
        public void Send(
            [Optional, DefaultParameterValue(null)] object? SaveCopy,
            [Optional, DefaultParameterValue(null)] object? ShowDialog,
            [Optional, DefaultParameterValue(null)] object? ParentWindow)
        {
            if (Recipients.Count == 0)
                throw ScriptonMapiSession.Error(unchecked((int)0x80040607), "MAPI_E_NO_RECIPIENTS", null);
            _mail.Send(
                new HostMailMessage(_mailbox, Recipients.Addresses(1), Recipients.Addresses(2), Recipients.Addresses(3), Subject, Text, htmlBody: null),
                smtp: null);
        }
    }

    [SourceClassName("Recipients")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonMapiRecipients : ScriptonIReflectOnClrType
    {
        private readonly List<ScriptonMapiRecipient> _recipients = new List<ScriptonMapiRecipient>();

        public ScriptonMapiRecipients() : base(typeof(ScriptonMapiRecipients))
        {
        }

        [DispId(0x000000A0)]
        public int Count => _recipients.Count;

        /// <summary>
        /// Add([Name], [Address], [Type], [EntryID]) - Type 1 To (the default), 2 CC, 3 BCC; the address is Address (without "SMTP:") or Name
        /// </summary>
        [DispId(0x000000A1)]
        public ScriptonMapiRecipient Add(
            [Optional, DefaultParameterValue(null)] object? Name,
            [Optional, DefaultParameterValue(null)] object? Address,
            [Optional, DefaultParameterValue(null)] object? Type,
            [Optional, DefaultParameterValue(null)] object? EntryID)
        {
            string? type = ScriptonMapiSession.Text(Type);
            var recipient = new ScriptonMapiRecipient(ScriptonMapiSession.Text(Name) ?? "", ScriptonMapiSession.Text(Address) ?? "",
                ((type != null) && int.TryParse(type, NumberStyles.Integer, CultureInfo.InvariantCulture, out int number)) ? number : 1);
            _recipients.Add(recipient);
            return recipient;
        }

        /// <summary>
        /// The addresses are used as they are (there is no address book) - Resolve is accepted
        /// </summary>
        [DispId(0x000000A2)]
#pragma warning disable CA1822 // Bound by name from scripts, so it must be an instance member
        public void Resolve([Optional, DefaultParameterValue(null)] object? ShowDialog)
        {
        }
#pragma warning restore CA1822

        internal IReadOnlyList<string> Addresses(int type) => _recipients.Where(r => r.Type == type).Select(r => r.EmailAddress).Where(a => a.Length > 0).ToList();
    }

    [SourceClassName("Recipient")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonMapiRecipient : ScriptonIReflectOnClrType
    {
        public ScriptonMapiRecipient(string name, string address, int type) : base(typeof(ScriptonMapiRecipient))
        {
            Name = name;
            Address = address;
            Type = type;
        }

        [DispId(0x000000AA)]
        public string Name { get; set; }

        [DispId(0x000000AB)]
        public string Address { get; set; }

        [DispId(0x000000AC)]
        public int Type { get; set; }

        internal string EmailAddress
        {
            get
            {
                string address = (Address.Length > 0) ? Address : Name;
                return address.StartsWith("SMTP:", StringComparison.OrdinalIgnoreCase) ? address.Substring(5) : address.Trim();
            }
        }
    }
}
