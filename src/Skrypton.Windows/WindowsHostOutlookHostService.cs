using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Components;

namespace Skrypton.Windows
{
    /// <summary>
    /// Stores, sends and shows Outlook items with the installed Microsoft Outlook (its COM automation, late bound): the items of the built-in
    /// Outlook.Application become real Outlook items in the user's mailbox
    /// </summary>
    public sealed class WindowsHostOutlookHostService : IHostOutlookHostService
    {
        /// <summary>
        /// The properties that are read back from Outlook items (Items.Find, For Each)
        /// </summary>
        private static readonly Dictionary<HostOutlookItemType, string[]> ReadProperties = new Dictionary<HostOutlookItemType, string[]>
        {
            { HostOutlookItemType.MailItem, new[] { "Subject", "To", "CC", "BCC", "Body", "HTMLBody", "SentOnBehalfOfName" } },
            { HostOutlookItemType.AppointmentItem, new[] { "Subject", "Body", "Location", "Start", "Duration", "AllDayEvent", "BusyStatus", "ReminderSet", "ReminderMinutesBeforeStart" } },
            { HostOutlookItemType.ContactItem, new[] { "FirstName", "LastName", "FileAs", "CompanyName", "JobTitle", "Email1Address", "Email2Address", "BusinessTelephoneNumber", "BusinessFaxNumber", "MobileTelephoneNumber", "BusinessAddressStreet", "BusinessAddressCity", "BusinessAddressState", "BusinessAddressPostalCode", "BusinessAddressCountry", "Birthday" } },
        };

        public string ItemSave(HostOutlookItem item)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(item));
            object outlookItem = ToOutlookItem(item);
            try
            {
                Call(outlookItem, "Save");
                return Convert.ToString(Get(outlookItem, "EntryID"), CultureInfo.InvariantCulture) ?? "";
            }
            finally
            {
                Release(outlookItem);
            }
        }

        public void ItemSend(HostOutlookItem item)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(item));
            object outlookItem = ToOutlookItem(item);
            try
            {
                Call(outlookItem, "Send");
            }
            finally
            {
                Release(outlookItem);
            }
        }

        public void ItemDisplay(HostOutlookItem item, bool modal)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(item));
            object outlookItem = ToOutlookItem(item);
            try
            {
                Call(outlookItem, "Display", modal);
            }
            finally
            {
                Release(outlookItem);
            }
        }

        public IEnumerable<HostOutlookItem> FolderItems(HostOutlookFolder folder)
        {
            if (folder == null)
                throw new ArgumentNullException(nameof(folder));
            var result = new List<HostOutlookItem>();
            object session = Get(Outlook(), "Session")!;
            object outlookFolder = (folder.SharedOwner == null)
                ? Call(session, "GetDefaultFolder", folder.DefaultFolderType)!
                : Call(session, "GetSharedDefaultFolder", Call(session, "CreateRecipient", folder.SharedOwner)!, folder.DefaultFolderType)!;
            object items = Get(outlookFolder, "Items")!;
            int count = Convert.ToInt32(Get(items, "Count"), CultureInfo.InvariantCulture);
            for (int index = 1; index <= count; index++)
            {
                object outlookItem = Call(items, "Item", index)!;
                try
                {
                    HostOutlookItem? item = FromOutlookItem(outlookItem);
                    if (item != null)
                        result.Add(item);
                }
                finally
                {
                    Release(outlookItem);
                }
            }
            Release(items);
            Release(outlookFolder);
            return result;
        }

        public bool RecipientResolve(string name)
        {
            object session = Get(Outlook(), "Session")!;
            object recipient = Call(session, "CreateRecipient", name ?? "")!;
            try
            {
                return Convert.ToBoolean(Call(recipient, "Resolve"), CultureInfo.InvariantCulture);
            }
            finally
            {
                Release(recipient);
            }
        }

        /// <summary>
        /// The Outlook item for the item: the saved one (EntryID) with the properties set again, or a new one
        /// </summary>
        private static object ToOutlookItem(HostOutlookItem item)
        {
            object outlook = Outlook();
            object outlookItem = (item.EntryId != null)
                ? Call(Get(outlook, "Session")!, "GetItemFromID", item.EntryId)!
                : Call(outlook, "CreateItem", (int)item.ItemType)!;
            foreach (KeyValuePair<string, object?> property in item.Properties)
            {
                if (string.Equals(property.Key, "Sent", StringComparison.OrdinalIgnoreCase))
                    continue;
                Set(outlookItem, property.Key, property.Value);
            }
            if (item.EntryId == null)
            {
                object attachments = Get(outlookItem, "Attachments")!;
                foreach (string attachment in item.Attachments)
                    Call(attachments, "Add", attachment);
                Release(attachments);
            }
            return outlookItem;
        }

        private static HostOutlookItem? FromOutlookItem(object outlookItem)
        {
            // OlObjectClass: 43 olMail, 26 olAppointment, 40 olContact (other items, eg. distribution lists, are not listed)
            int itemClass = Convert.ToInt32(Get(outlookItem, "Class"), CultureInfo.InvariantCulture);
            HostOutlookItemType type;
            switch (itemClass)
            {
                case 43: type = HostOutlookItemType.MailItem; break;
                case 26: type = HostOutlookItemType.AppointmentItem; break;
                case 40: type = HostOutlookItemType.ContactItem; break;
                default: return null;
            }
            var item = new HostOutlookItem(type) { EntryId = Convert.ToString(Get(outlookItem, "EntryID"), CultureInfo.InvariantCulture) };
            foreach (string property in ReadProperties[type])
                item.Properties[property] = Get(outlookItem, property);
            return item;
        }

        private static object Outlook()
        {
            Type type = Type.GetTypeFromProgID("Outlook.Application")
                ?? throw new InvalidOperationException("Outlook.Application: Microsoft Outlook is not installed");
            return Activator.CreateInstance(type) ?? throw new InvalidOperationException("Outlook.Application: Microsoft Outlook could not be started");
        }

        private static object? Get(object target, string name) => target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, null, CultureInfo.InvariantCulture);

        private static void Set(object target, string name, object? value) => target.GetType().InvokeMember(name, BindingFlags.SetProperty, null, target, new[] { value }, CultureInfo.InvariantCulture);

        private static object? Call(object target, string name, params object?[] args) => target.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, target, args, CultureInfo.InvariantCulture);

        private static void Release(object? comObject)
        {
            if ((comObject != null) && Marshal.IsComObject(comObject))
                Marshal.ReleaseComObject(comObject);
        }
    }
}
