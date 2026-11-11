#nullable enable
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Skrypton.RuntimeSupport.Components;

namespace Skrypton.Tests.RuntimeSupport.Implementations
{
    /// <summary>
    /// An in-memory IHostOutlookHostService for the tests: saved items go to the folder of their type (contacts, calendar, drafts), sent mails
    /// and displayed items are recorded, recipients with an "@" are resolved
    /// </summary>
    internal sealed class TestOutlookHostService : IHostOutlookHostService
    {
        public List<HostOutlookItem> Saved { get; } = new List<HostOutlookItem>();
        public List<HostOutlookItem> Sent { get; } = new List<HostOutlookItem>();
        public List<(HostOutlookItem Item, bool Modal)> Displayed { get; } = new List<(HostOutlookItem, bool)>();

        public string ItemSave(HostOutlookItem item)
        {
            if (!Saved.Contains(item))
                Saved.Add(item);
            return item.EntryId ?? ("ENTRY" + Saved.Count.ToString(CultureInfo.InvariantCulture));
        }

        public void ItemSend(HostOutlookItem item) => Sent.Add(item);

        public void ItemDisplay(HostOutlookItem item, bool modal) => Displayed.Add((item, modal));

        public IEnumerable<HostOutlookItem> FolderItems(HostOutlookFolder folder)
        {
            HostOutlookItemType type = folder.DefaultFolderType switch
            {
                9 => HostOutlookItemType.AppointmentItem,
                10 => HostOutlookItemType.ContactItem,
                _ => HostOutlookItemType.MailItem,
            };
            return Saved.Where(i => i.ItemType == type).ToList();
        }

        public bool RecipientResolve(string name) => name.Contains('@');
    }
}
