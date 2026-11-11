// The members are bound by name from scripts, so they must be instance members
#pragma warning disable CA1822
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;
using Skrypton.RuntimeSupport.Implementations;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The built-in Outlook.Application - mail, contact and appointment items are created and filled in as with Outlook; saving, sending,
    /// displaying, the items of folders and resolving recipients go through the host's IHostOutlookHostService (there is no default, those
    /// members raise an error if the host has not registered one).
    /// </summary>
    [DefaultMember("Name")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Application")] // TypeName(CreateObject("Outlook.Application")) (as with Windows VBScript)
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonOutlookApplication : ScriptonIReflectOnClrType
    {
        public const string ProgId = "Outlook.Application";
        internal const string ErrorSource = "Microsoft Outlook";

        private readonly IRuntimeHost _runtimeHost;
        private readonly ScriptonOutlookNameSpace _session;

        public ScriptonOutlookApplication(IRuntimeHost runtimeHost) : base(typeof(ScriptonOutlookApplication))
        {
            _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
            _session = new ScriptonOutlookNameSpace(runtimeHost);
        }

        /// <summary>
        /// The default member ("" &amp; CreateObject("Outlook.Application") is "Outlook", VarType 8)
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public string Name => "Outlook";

        [DispId(278)] public string Version => "16.0.0.0";

        [DispId(61451)] public object Session => _session;

        /// <summary>
        /// CreateItem(ItemType) - 0 olMailItem, 1 olAppointmentItem, 2 olContactItem
        /// </summary>
        [DispId(266)]
        public object CreateItem([In] object ItemType) => CreateItem(_runtimeHost, ItemType);

        internal static ScriptonOutlookItem CreateItem(IRuntimeHost runtimeHost, object? itemType)
        {
            int type = IsMissing(itemType) ? 0 : Convert.ToInt32(itemType, CultureInfo.InvariantCulture);
            return type switch
            {
                (int)HostOutlookItemType.MailItem => new ScriptonOutlookMailItem(runtimeHost, new HostOutlookItem(HostOutlookItemType.MailItem)),
                (int)HostOutlookItemType.AppointmentItem => new ScriptonOutlookAppointmentItem(runtimeHost, new HostOutlookItem(HostOutlookItemType.AppointmentItem)),
                (int)HostOutlookItemType.ContactItem => new ScriptonOutlookContactItem(runtimeHost, new HostOutlookItem(HostOutlookItemType.ContactItem)),
                _ => throw InvalidProcedureCall(),
            };
        }

        internal static ScriptonOutlookItem Wrap(IRuntimeHost runtimeHost, HostOutlookItem item)
        {
            return item.ItemType switch
            {
                HostOutlookItemType.AppointmentItem => new ScriptonOutlookAppointmentItem(runtimeHost, item),
                HostOutlookItemType.ContactItem => new ScriptonOutlookContactItem(runtimeHost, item),
                _ => new ScriptonOutlookMailItem(runtimeHost, item),
            };
        }

        /// <summary>
        /// GetNamespace("MAPI")
        /// </summary>
        [DispId(272)]
        public object GetNamespace([In, MarshalAs(UnmanagedType.BStr)] string Type)
        {
            if (!string.Equals(Type, "MAPI", StringComparison.OrdinalIgnoreCase))
                throw InvalidProcedureCall();
            return _session;
        }

        [DispId(275)]
        public void Quit()
        {
        }

        internal static CustomException InvalidProcedureCall() => new CustomException(5, ErrorSource, "Invalid procedure call or argument");

        internal static bool IsMissing(object? value) => (value == null) || (value is Missing);

        internal static IHostOutlookHostService Host(IRuntimeHost runtimeHost, string member)
        {
            return runtimeHost.TryGetRuntimeHostService<IHostOutlookHostService>() ?? throw new InvalidOperationException(
                $"Outlook {member}: Outlook items are stored, sent and shown by the host, which requires the host to register an {nameof(IHostOutlookHostService)}"
            );
        }
    }

    [DefaultMember("Type")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("NameSpace")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonOutlookNameSpace : ScriptonIReflectOnClrType
    {
        private readonly IRuntimeHost _runtimeHost;

        public ScriptonOutlookNameSpace(IRuntimeHost runtimeHost) : base(typeof(ScriptonOutlookNameSpace))
        {
            _runtimeHost = runtimeHost;
        }

        /// <summary>
        /// The default member ("Mapi", as with Outlook)
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public string Type => "Mapi";

        /// <summary>
        /// GetDefaultFolder(FolderType) - OlDefaultFolders (6 inbox, 9 calendar, 10 contacts, ..)
        /// </summary>
        [DispId(8459)]
        public object GetDefaultFolder([In] object FolderType) => new ScriptonOutlookFolder(_runtimeHost, new HostOutlookFolder(ScriptonOutlookFolder.ToFolderType(FolderType), null));

        [DispId(8458)]
        public object CreateRecipient([In, MarshalAs(UnmanagedType.BStr)] string RecipientName) => new ScriptonOutlookRecipient(_runtimeHost, RecipientName ?? "", 1);

        /// <summary>
        /// GetSharedDefaultFolder(Recipient, FolderType) - a default folder of another user
        /// </summary>
        [DispId(8460)]
        public object GetSharedDefaultFolder([In] object Recipient, [In] object FolderType)
        {
            string owner = (Recipient as ScriptonOutlookRecipient)?.Name ?? Convert.ToString(Recipient, CultureInfo.InvariantCulture) ?? "";
            return new ScriptonOutlookFolder(_runtimeHost, new HostOutlookFolder(ScriptonOutlookFolder.ToFolderType(FolderType), owner));
        }

        [DispId(8449)]
        public void Logon([In, Optional, DefaultParameterValue(null)] object? Profile, [In, Optional, DefaultParameterValue(null)] object? Password, [In, Optional, DefaultParameterValue(null)] object? ShowDialog, [In, Optional, DefaultParameterValue(null)] object? NewSession)
        {
        }

        [DispId(8450)]
        public void Logoff()
        {
        }
    }

    /// <summary>
    /// A default folder - its name is the English name of the folder (Outlook shows the name in the language of the mailbox)
    /// </summary>
    [DefaultMember("Name")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("MAPIFolder")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonOutlookFolder : ScriptonIReflectOnClrType
    {
        private static readonly Dictionary<int, (string Name, HostOutlookItemType ItemType)> Folders = new Dictionary<int, (string, HostOutlookItemType)>
        {
            { 3, ("Deleted Items", HostOutlookItemType.MailItem) },
            { 4, ("Outbox", HostOutlookItemType.MailItem) },
            { 5, ("Sent Items", HostOutlookItemType.MailItem) },
            { 6, ("Inbox", HostOutlookItemType.MailItem) },
            { 9, ("Calendar", HostOutlookItemType.AppointmentItem) },
            { 10, ("Contacts", HostOutlookItemType.ContactItem) },
            { 16, ("Drafts", HostOutlookItemType.MailItem) },
            { 23, ("Junk Email", HostOutlookItemType.MailItem) },
        };

        private readonly IRuntimeHost _runtimeHost;

        public ScriptonOutlookFolder(IRuntimeHost runtimeHost, HostOutlookFolder folder) : base(typeof(ScriptonOutlookFolder))
        {
            _runtimeHost = runtimeHost;
            Folder = folder;
        }

        internal HostOutlookFolder Folder { get; }

        internal static int ToFolderType(object? folderType)
        {
            int type = Convert.ToInt32(folderType, CultureInfo.InvariantCulture);
            if (!Folders.ContainsKey(type))
                throw ScriptonOutlookApplication.InvalidProcedureCall();
            return type;
        }

        [DispId(0)]
        [IsDefault]
        public string Name => Folders[Folder.DefaultFolderType].Name;

        [DispId(12550)] public int DefaultItemType => (int)Folders[Folder.DefaultFolderType].ItemType;

        [DispId(12544)] public object Items => new ScriptonOutlookItems(_runtimeHost, this);
    }

    /// <summary>
    /// The items of a folder - it has no default member (VarType 9)
    /// </summary>
    [DefaultMember("Item")] // +[DispId(0)] +[IsDefault] (a parameter, so VarType stays 9)
    [SourceClassName("Items")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonOutlookItems : ScriptonIReflectOnClrType, IEnumerable
    {
        private static readonly Regex Condition = new Regex(@"^\s*\[(?<property>[^\]]+)\]\s*=\s*(?:'(?<value>[^']*)'|""(?<value>[^""]*)"")\s*$", RegexOptions.CultureInvariant);
        private static readonly Regex And = new Regex(@"\s+and\s+", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        private readonly IRuntimeHost _runtimeHost;
        private readonly ScriptonOutlookFolder _folder;

        public ScriptonOutlookItems(IRuntimeHost runtimeHost, ScriptonOutlookFolder folder) : base(typeof(ScriptonOutlookItems))
        {
            _runtimeHost = runtimeHost;
            _folder = folder;
        }

        private List<HostOutlookItem> All(string member) => ScriptonOutlookApplication.Host(_runtimeHost, member).FolderItems(_folder.Folder).ToList();

        [DispId(80)] public int Count => All("Items.Count").Count;

        [DispId(0)]
        [IsDefault]
        public object Item([In] object Index)
        {
            List<HostOutlookItem> items = All("Items.Item");
            int index = Convert.ToInt32(Index, CultureInfo.InvariantCulture);
            if ((index < 1) || (index > items.Count))
                throw new CustomException(-2147352567, ScriptonOutlookApplication.ErrorSource, "Array index out of bounds.");
            return ScriptonOutlookApplication.Wrap(_runtimeHost, items[index - 1]);
        }

        /// <summary>
        /// Find(Filter) - the first item that matches, or Nothing; the filter is "[Property] = 'value'" conditions joined with And
        /// </summary>
        [DispId(98)]
        public object Find([In, MarshalAs(UnmanagedType.BStr)] string Filter)
        {
            var conditions = new List<(string Property, string Value)>();
            foreach (string part in And.Split(Filter ?? ""))
            {
                Match match = Condition.Match(part);
                if (!match.Success)
                    throw new NotSupportedException($"Items.Find: the filter '{Filter}' is not supported (only \"[Property] = 'value'\" conditions joined with And are)");
                conditions.Add((match.Groups["property"].Value, match.Groups["value"].Value));
            }
            HostOutlookItem? found = All("Items.Find").FirstOrDefault(item => conditions.All(c =>
                string.Equals(Convert.ToString(ScriptonOutlookItem.GetValue(item, c.Property), CultureInfo.InvariantCulture) ?? "", c.Value, StringComparison.OrdinalIgnoreCase)));
            return (found == null) ? VBScriptConstants.Nothing : ScriptonOutlookApplication.Wrap(_runtimeHost, found);
        }

        /// <summary>
        /// Add([Type]) - a new item of the folder's item type (not saved yet)
        /// </summary>
        [DispId(95)]
        public object Add([In, Optional, DefaultParameterValue(null)] object? Type)
        {
            return ScriptonOutlookApplication.CreateItem(_runtimeHost, ScriptonOutlookApplication.IsMissing(Type) ? _folder.DefaultItemType : Type);
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator() => All("Items").Select(i => ScriptonOutlookApplication.Wrap(_runtimeHost, i)).ToList().GetEnumerator();
    }

    /// <summary>
    /// The base of the items: the properties are kept in the HostOutlookItem (with Outlook's defaults), Save / Display go through the host
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal abstract class ScriptonOutlookItem : ScriptonIReflectOnClrType
    {
        protected ScriptonOutlookItem(Type type, IRuntimeHost runtimeHost, HostOutlookItem item) : base(type)
        {
            RuntimeHost = runtimeHost;
            Data = item;
        }

        protected IRuntimeHost RuntimeHost { get; }

        internal HostOutlookItem Data { get; }

        internal static object? GetValue(HostOutlookItem item, string property)
        {
            if (string.Equals(property, "FullName", StringComparison.OrdinalIgnoreCase))
                return ScriptonOutlookContactItem.FullNameOf(item);
            return item.GetProperty(property);
        }

        protected string GetText(string property) => Convert.ToString(Data.GetProperty(property), CultureInfo.CurrentCulture) ?? "";

        protected void SetValue(string property, object? value)
        {
            Data.Properties[property] = value;
            Saved = false;
        }

        [DispId(61603)] public bool Saved { get; private set; } = true; // a new item is "saved" until it is changed (as with Outlook)

        [DispId(61728)] public string EntryID => Data.EntryId ?? "";

        [DispId(12289)] public object Attachments => new ScriptonOutlookAttachments(RuntimeHost, this);

        [DispId(61512)]
        public void Save()
        {
            Data.EntryId = ScriptonOutlookApplication.Host(RuntimeHost, "Save").ItemSave(Data);
            Saved = true;
        }

        [DispId(61606)]
        public void Display([In, Optional, DefaultParameterValue(null)] object? Modal)
        {
            bool modal = !ScriptonOutlookApplication.IsMissing(Modal) && Convert.ToBoolean(Modal, CultureInfo.InvariantCulture);
            ScriptonOutlookApplication.Host(RuntimeHost, "Display").ItemDisplay(Data, modal);
        }

        [DispId(61475)]
        public void Close([In] object SaveMode)
        {
            // 0 olSave, 1 olDiscard, 2 olPromptForSave (nobody to ask: discard)
            if (Convert.ToInt32(SaveMode, CultureInfo.InvariantCulture) == 0)
                Save();
        }
    }

    [DefaultMember("Subject")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("MailItem")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonOutlookMailItem : ScriptonOutlookItem
    {
        private readonly ScriptonOutlookRecipients _recipients;

        public ScriptonOutlookMailItem(IRuntimeHost runtimeHost, HostOutlookItem item) : base(typeof(ScriptonOutlookMailItem), runtimeHost, item)
        {
            _recipients = new ScriptonOutlookRecipients(runtimeHost, this);
        }

        [DispId(61450)] public int Class => 43; // olMail

        [DispId(0)]
        [IsDefault]
        public object Subject { get => GetText("Subject"); set => SetValue("Subject", value); }

        [DispId(37120)] public object Body { get => GetText("Body"); set => SetValue("Body", value); }

        /// <summary>
        /// The HTML body - setting it makes the mail an HTML mail (BodyFormat 2)
        /// </summary>
        [DispId(62468)]
        public object HTMLBody
        {
            get => GetText("HTMLBody");
            set
            {
                SetValue("HTMLBody", value);
                SetValue("BodyFormat", 2);
            }
        }

        [DispId(64073)] public object BodyFormat { get => Data.GetProperty("BodyFormat") ?? 2; set => SetValue("BodyFormat", value); }

        [DispId(23)] public object Importance { get => Data.GetProperty("Importance") ?? 1; set => SetValue("Importance", value); }

        [DispId(15937)] public object SentOnBehalfOfName { get => GetText("SentOnBehalfOfName"); set => SetValue("SentOnBehalfOfName", value); }

        [DispId(3588)] public object To { get => _recipients.Names(1); set => _recipients.Replace(1, Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""); }

        [DispId(3587)] public object CC { get => _recipients.Names(2); set => _recipients.Replace(2, Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""); }

        [DispId(3586)] public object BCC { get => _recipients.Names(3); set => _recipients.Replace(3, Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""); }

        [DispId(63508)] public object Recipients => _recipients;

        [DispId(62466)] public bool Sent => (Data.GetProperty("Sent") as bool?) ?? false;

        [DispId(61557)]
        public void Send()
        {
            if (_recipients.Count == 0)
                throw new CustomException(-2147467259, ScriptonOutlookApplication.ErrorSource, "Outlook does not recognize one or more names.");
            ScriptonOutlookApplication.Host(RuntimeHost, "Send").ItemSend(Data);
            Data.Properties["Sent"] = true;
        }

        internal void RecipientsChanged() => SetValue("To", Data.GetProperty("To"));
    }

    [DefaultMember("FileAs")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("ContactItem")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonOutlookContactItem : ScriptonOutlookItem
    {
        public ScriptonOutlookContactItem(IRuntimeHost runtimeHost, HostOutlookItem item) : base(typeof(ScriptonOutlookContactItem), runtimeHost, item)
        {
        }

        internal static string FullNameOf(HostOutlookItem item)
        {
            string first = Convert.ToString(item.GetProperty("FirstName"), CultureInfo.InvariantCulture) ?? "";
            string last = Convert.ToString(item.GetProperty("LastName"), CultureInfo.InvariantCulture) ?? "";
            return (first + " " + last).Trim();
        }

        [DispId(61450)] public int Class => 40; // olContact

        /// <summary>
        /// The default member: "LastName, FirstName" (unless it was set)
        /// </summary>
        [DispId(0)]
        [IsDefault]
        public object FileAs
        {
            get
            {
                if (Data.GetProperty("FileAs") is object fileAs)
                    return fileAs;
                string first = GetText("FirstName");
                string last = GetText("LastName");
                return ((last.Length > 0) && (first.Length > 0)) ? last + ", " + first : last + first;
            }
            set => SetValue("FileAs", value);
        }

        [DispId(12289 + 1)] public object FullName => FullNameOf(Data);

        [DispId(14854)] public object FirstName { get => GetText("FirstName"); set => SetValue("FirstName", value); }
        [DispId(14865)] public object LastName { get => GetText("LastName"); set => SetValue("LastName", value); }
        [DispId(14870)] public object CompanyName { get => GetText("CompanyName"); set => SetValue("CompanyName", value); }
        [DispId(14871)] public object JobTitle { get => GetText("JobTitle"); set => SetValue("JobTitle", value); }
        [DispId(32837)] public object BusinessAddressStreet { get => GetText("BusinessAddressStreet"); set => SetValue("BusinessAddressStreet", value); }
        [DispId(32838)] public object BusinessAddressCity { get => GetText("BusinessAddressCity"); set => SetValue("BusinessAddressCity", value); }
        [DispId(32839)] public object BusinessAddressState { get => GetText("BusinessAddressState"); set => SetValue("BusinessAddressState", value); }
        [DispId(32840)] public object BusinessAddressPostalCode { get => GetText("BusinessAddressPostalCode"); set => SetValue("BusinessAddressPostalCode", value); }
        [DispId(32841)] public object BusinessAddressCountry { get => GetText("BusinessAddressCountry"); set => SetValue("BusinessAddressCountry", value); }
        [DispId(14856)] public object BusinessTelephoneNumber { get => GetText("BusinessTelephoneNumber"); set => SetValue("BusinessTelephoneNumber", value); }
        [DispId(14884)] public object BusinessFaxNumber { get => GetText("BusinessFaxNumber"); set => SetValue("BusinessFaxNumber", value); }
        [DispId(14876)] public object MobileTelephoneNumber { get => GetText("MobileTelephoneNumber"); set => SetValue("MobileTelephoneNumber", value); }
        [DispId(32899)] public object Email1Address { get => GetText("Email1Address"); set => SetValue("Email1Address", value); }
        [DispId(32915)] public object Email2Address { get => GetText("Email2Address"); set => SetValue("Email2Address", value); }

        /// <summary>
        /// The birthday - "none" is 1/1/4501, as with Outlook
        /// </summary>
        [DispId(14914)] public object Birthday { get => Data.GetProperty("Birthday") ?? new DateTime(4501, 1, 1); set => SetValue("Birthday", value); }
    }

    /// <summary>
    /// An appointment - it has no default member (VarType 9); a new appointment starts at the next half hour and lasts 30 minutes, with a
    /// reminder 15 minutes before
    /// </summary>
    [SourceClassName("AppointmentItem")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonOutlookAppointmentItem : ScriptonOutlookItem
    {
        public ScriptonOutlookAppointmentItem(IRuntimeHost runtimeHost, HostOutlookItem item) : base(typeof(ScriptonOutlookAppointmentItem), runtimeHost, item)
        {
            if (item.GetProperty("Start") == null)
            {
                DateTime now = DateTime.Now;
                var start = new DateTime(now.Year, now.Month, now.Day, now.Hour, (now.Minute < 30) ? 30 : 0, 0);
                if (now.Minute >= 30)
                    start = start.AddHours(1);
                item.Properties["Start"] = start;
            }
        }

        [DispId(61450)] public int Class => 26; // olAppointment

        [DispId(55)] public object Subject { get => GetText("Subject"); set => SetValue("Subject", value); }
        [DispId(37120)] public object Body { get => GetText("Body"); set => SetValue("Body", value); }
        [DispId(33288)] public object Location { get => GetText("Location"); set => SetValue("Location", value); }

        [DispId(33293)]
        public object Start
        {
            get => Data.GetProperty("Start")!;
            set => SetValue("Start", Convert.ToDateTime(value, CultureInfo.CurrentCulture));
        }

        /// <summary>
        /// The end - setting it changes the duration
        /// </summary>
        [DispId(33294)]
        public object End
        {
            get => ((DateTime)Start).AddMinutes(Convert.ToInt32(Duration, CultureInfo.InvariantCulture));
            set => Duration = (int)Math.Round((Convert.ToDateTime(value, CultureInfo.CurrentCulture) - (DateTime)Start).TotalMinutes);
        }

        [DispId(33299)] public object Duration { get => Data.GetProperty("Duration") ?? 30; set => SetValue("Duration", Convert.ToInt32(value, CultureInfo.InvariantCulture)); }
        [DispId(33301)] public object AllDayEvent { get => Data.GetProperty("AllDayEvent") ?? false; set => SetValue("AllDayEvent", value); }
        [DispId(33285)] public object BusyStatus { get => Data.GetProperty("BusyStatus") ?? 2; set => SetValue("BusyStatus", value); }
        [DispId(34049)] public object ReminderMinutesBeforeStart { get => Data.GetProperty("ReminderMinutesBeforeStart") ?? 15; set => SetValue("ReminderMinutesBeforeStart", value); }
        [DispId(34051)] public object ReminderSet { get => Data.GetProperty("ReminderSet") ?? true; set => SetValue("ReminderSet", value); }
        [DispId(34078)] public object ReminderPlaySound { get => Data.GetProperty("ReminderPlaySound") ?? true; set => SetValue("ReminderPlaySound", value); }
    }

    [DefaultMember("Item")] // +[DispId(0)] +[IsDefault] (a parameter, so VarType stays 9)
    [SourceClassName("Recipients")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonOutlookRecipients : ScriptonIReflectOnClrType, IEnumerable
    {
        private readonly IRuntimeHost _runtimeHost;
        private readonly ScriptonOutlookMailItem _mail;
        private readonly List<ScriptonOutlookRecipient> _recipients = new List<ScriptonOutlookRecipient>();

        public ScriptonOutlookRecipients(IRuntimeHost runtimeHost, ScriptonOutlookMailItem mail) : base(typeof(ScriptonOutlookRecipients))
        {
            _runtimeHost = runtimeHost;
            _mail = mail;
        }

        [DispId(80)] public int Count => _recipients.Count;

        [DispId(0)]
        [IsDefault]
        public object Item([In] object Index)
        {
            int index = Convert.ToInt32(Index, CultureInfo.InvariantCulture);
            if ((index < 1) || (index > _recipients.Count))
                throw new CustomException(-2147352567, ScriptonOutlookApplication.ErrorSource, "Array index out of bounds.");
            return _recipients[index - 1];
        }

        /// <summary>
        /// Add(Name) - a To recipient
        /// </summary>
        [DispId(111)]
        public object Add([In, MarshalAs(UnmanagedType.BStr)] string Name)
        {
            var recipient = new ScriptonOutlookRecipient(_runtimeHost, Name ?? "", 1);
            _recipients.Add(recipient);
            Changed();
            return recipient;
        }

        internal string Names(int type) => string.Join("; ", _recipients.Where(r => r.Type == type).Select(r => r.Name));

        /// <summary>
        /// Set To / CC / BCC: the recipients of the type are replaced by the names (separated with ';')
        /// </summary>
        internal void Replace(int type, string names)
        {
            _recipients.RemoveAll(r => r.Type == type);
            foreach (string name in names.Split(';').Select(n => n.Trim()).Where(n => n.Length > 0))
                _recipients.Add(new ScriptonOutlookRecipient(_runtimeHost, name, type));
            Changed();
        }

        private void Changed()
        {
            _mail.Data.Properties["To"] = Names(1);
            _mail.Data.Properties["CC"] = Names(2);
            _mail.Data.Properties["BCC"] = Names(3);
            _mail.RecipientsChanged();
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator() => _recipients.ToList().GetEnumerator();
    }

    [DefaultMember("Name")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Recipient")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonOutlookRecipient : ScriptonIReflectOnClrType
    {
        private readonly IRuntimeHost _runtimeHost;

        public ScriptonOutlookRecipient(IRuntimeHost runtimeHost, string name, int type) : base(typeof(ScriptonOutlookRecipient))
        {
            _runtimeHost = runtimeHost;
            Name = name;
            Type = type;
        }

        [DispId(0)]
        [IsDefault]
        public string Name { get; }

        /// <summary>
        /// 1 olTo, 2 olCC, 3 olBCC
        /// </summary>
        [DispId(3093)] public int Type { get; set; }

        [DispId(100)] public bool Resolved { get; private set; }

        [DispId(113)]
        public bool Resolve()
        {
            Resolved = ScriptonOutlookApplication.Host(_runtimeHost, "Recipient.Resolve").RecipientResolve(Name);
            return Resolved;
        }
    }

    [DefaultMember("Item")] // +[DispId(0)] +[IsDefault] (a parameter, so VarType stays 9)
    [SourceClassName("Attachments")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonOutlookAttachments : ScriptonIReflectOnClrType, IEnumerable
    {
        private readonly IRuntimeHost _runtimeHost;
        private readonly ScriptonOutlookItem _item;

        public ScriptonOutlookAttachments(IRuntimeHost runtimeHost, ScriptonOutlookItem item) : base(typeof(ScriptonOutlookAttachments))
        {
            _runtimeHost = runtimeHost;
            _item = item;
        }

        [DispId(80)] public int Count => _item.Data.Attachments.Count;

        [DispId(0)]
        [IsDefault]
        public object Item([In] object Index)
        {
            int index = Convert.ToInt32(Index, CultureInfo.InvariantCulture);
            if ((index < 1) || (index > _item.Data.Attachments.Count))
                throw new CustomException(-2147352567, ScriptonOutlookApplication.ErrorSource, "Array index out of bounds.");
            string path = _item.Data.Attachments[index - 1];
            return new ScriptonOutlookAttachment(path, ScriptonFileSystemObject.FileNameOf(path));
        }

        /// <summary>
        /// Add(Source, [Type], [Position], [DisplayName]) - a file (it must exist)
        /// </summary>
        [DispId(101)]
        public object Add([In] object Source, [In, Optional, DefaultParameterValue(null)] object? Type, [In, Optional, DefaultParameterValue(null)] object? Position, [In, Optional, DefaultParameterValue(null)] object? DisplayName)
        {
            string path = Convert.ToString(Source, CultureInfo.InvariantCulture) ?? "";
            IHostFileSystemHostService fileSystem = _runtimeHost.TryGetRuntimeHostService<IHostFileSystemHostService>() ?? DefaultHostFileSystemHostService.Instance;
            if (!fileSystem.FileExists(path))
                throw new CustomException(-2147024894, ScriptonOutlookApplication.ErrorSource, "Cannot find this file. Verify the path and file name are correct.");
            _item.Data.Attachments.Add(path);
            return new ScriptonOutlookAttachment(path, ScriptonOutlookApplication.IsMissing(DisplayName) ? ScriptonFileSystemObject.FileNameOf(path) : (Convert.ToString(DisplayName, CultureInfo.InvariantCulture) ?? ""));
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator() => _item.Data.Attachments.Select(a => new ScriptonOutlookAttachment(a, ScriptonFileSystemObject.FileNameOf(a))).ToList().GetEnumerator();
    }

    [DefaultMember("DisplayName")] // +[DispId(0)] +[IsDefault]
    [SourceClassName("Attachment")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    internal sealed class ScriptonOutlookAttachment : ScriptonIReflectOnClrType
    {
        public ScriptonOutlookAttachment(string path, string displayName) : base(typeof(ScriptonOutlookAttachment))
        {
            PathName = path;
            DisplayName = displayName;
        }

        [DispId(0)]
        [IsDefault]
        public string DisplayName { get; }

        [DispId(14084)] public string FileName => ScriptonFileSystemObject.FileNameOf(PathName);

        [DispId(14088)] public string PathName { get; }
    }
}
