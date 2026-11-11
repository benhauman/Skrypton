using System;
using System.Collections.Generic;

namespace Skrypton.RuntimeSupport.Components
{
    /// <summary>
    /// The host service that the built-in Outlook.Application stores, sends and shows its items through. Mail, contacts and appointments belong
    /// to the host (a mail server, a CRM, the user's mailbox, ..), so there is no default: creating items and setting their properties works
    /// without one, but saving, sending, displaying, finding items and resolving recipients raise an error if the host has not registered one.
    /// </summary>
    public interface IHostOutlookHostService
    {
        /// <summary>
        /// Save the item (a mail as a draft, a contact or an appointment in its folder) and return its EntryID
        /// </summary>
        string ItemSave(HostOutlookItem item);

        /// <summary>
        /// Send the mail (or the meeting request)
        /// </summary>
        void ItemSend(HostOutlookItem item);

        /// <summary>
        /// Show the item to the user (Display) - modal: wait until the user has closed it
        /// </summary>
        void ItemDisplay(HostOutlookItem item, bool modal);

        /// <summary>
        /// The items of the folder (for Items.Find, Items.Count, For Each)
        /// </summary>
        IEnumerable<HostOutlookItem> FolderItems(HostOutlookFolder folder);

        /// <summary>
        /// Whether the recipient (a name or an address) is known (Recipient.Resolve)
        /// </summary>
        bool RecipientResolve(string name);
    }

    /// <summary>
    /// The Outlook item types (OlItemType)
    /// </summary>
    public enum HostOutlookItemType
    {
        MailItem = 0,
        AppointmentItem = 1,
        ContactItem = 2,
    }

    /// <summary>
    /// An Outlook item: its type, its properties by their Outlook names ("Subject", "To", "FirstName", "Start", ..), its attachments (file paths)
    /// and its EntryID once it is saved
    /// </summary>
    public sealed class HostOutlookItem
    {
        public HostOutlookItem(HostOutlookItemType itemType)
        {
            ItemType = itemType;
        }

        public HostOutlookItemType ItemType { get; }

        public IDictionary<string, object?> Properties { get; } = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

        public IList<string> Attachments { get; } = new List<string>();

        public string? EntryId { get; set; }

        public object? GetProperty(string name) => Properties.TryGetValue(name, out object? value) ? value : null;
    }

    /// <summary>
    /// A default folder (OlDefaultFolders: 3 deleted items, 4 outbox, 5 sent mail, 6 inbox, 9 calendar, 10 contacts, 13 tasks, 16 drafts, ..) -
    /// of the current user, or of another user (GetSharedDefaultFolder)
    /// </summary>
    public sealed class HostOutlookFolder
    {
        public HostOutlookFolder(int defaultFolderType, string? sharedOwner)
        {
            DefaultFolderType = defaultFolderType;
            SharedOwner = sharedOwner;
        }

        public int DefaultFolderType { get; }

        public string? SharedOwner { get; }
    }
}
