using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Implementations;
using Skrypton.Tests.RuntimeSupport.Implementations;

namespace Skrypton.Tests.RuntimeSupport.Components.OutlookApplication;

[SourceClassName("Application")] // "Outlook.Application"
[ComVisible(true)]
public sealed class MyOutlookApplicationClass : IReflectOnClrType
{
    // "Interop.Microsoft.Office.Interop.Outlook.dll"  <Guid>00062fff-0000-0000-c000-000000000046</Guid>
    // In the Outlook COM type library (OUTL.TLB)
    // use VisualStudio [command prompt] or
    // c:\Program Files\Microsoft Visual Studio\18\Enterprise\Common7\Tools>OleView.exe
    // => Go to 'Type Libraries', Scroll down to 'Microsoft Outlook x.x Object Library' - DoubleClick
    // As Administrator : "C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\oleview.exe"
    private readonly IRuntimeHost _runtimeHost;

    public MyOutlookApplicationClass(IRuntimeHost runtimeHost)
    {
        _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
    }

    //[DispId(266)]
    //public object CreateItem()
    //{
    //    return CreateItem(OlItemType.olMailItem);
    //}

    [DispId(266)]
    public object CreateItem([In] object item)
    {
        return CreateItemCore(item);
    }
    internal static object CreateItemCore([In] object item)
    {
        if (item == null) throw new ArgumentNullException(nameof(item), "Parameter must be 0:olMailItem, 1:olAppointmentItem or 2:olContactItem");
        OlItemType ItemType = (OlItemType)Enum.ToObject(typeof(OlItemType), item);
        //OlItemType ItemType = (OlItemType)item;
        if (ItemType == OlItemType.olContactItem)
        {
            return new MyOutlookContactItemClass();
        }
        if (ItemType == OlItemType.olAppointmentItem)
        {
            return new MyOutlookAppointmentItemClass();
        }
        if (ItemType == OlItemType.olMailItem)
        {
            return new MyOutlookMailItemClass();
        }
        throw new NotImplementedException($"ItemType:{ItemType}");
    }

    [DispId(272)]
    public _NameSpace GetNamespace([In] string Type) // Type:MAPI
    {
        if (Type == "MAPI")
        {
            return new MyOutlookMAPISession();
        }
        throw new NotImplementedException($"Type:{Type}");
    }
}

internal sealed class MyOutlookMailItemClass : IReflectOnClrType // see interface _MailItem
{
    internal MyOutlookMailItemClass()
    {
    }
    [DispId(55)]
    public string Subject { get; set; }

    [DispId(3587)]
    public string CC { get; set; }
    [DispId(3588)]
    public string To { get; set; }

    [DispId(62468)]
    public string HTMLBody { get; set; }

    [DispId(61606)]
    public void Display()
    {

    }

    private MyOutlookAttachmentsClass _attachments;

    [DispId(63509)]
    public Attachments Attachments
    {
        get
        {
            if (_attachments == null)
            {
                _attachments = new MyOutlookAttachmentsClass();
            }
            return _attachments;
        }
    }
}

internal interface Attachments : IEnumerable
{

}

internal interface Attachment
{

}

internal sealed class MyOutlookAttachmentsClass : IReflectOnClrType, Attachments
{
    private List<MyOutlookAttachmentClass> _attachments = new List<MyOutlookAttachmentClass>();
    public MyOutlookAttachmentsClass()
    {

    }
    public IEnumerator GetEnumerator()
    {
        return _attachments.GetEnumerator();
    }

    public Attachment Add(object source, object attachmentType, object attachmentPosition, object attachmentDisplayName)
    {
        string sourceS = (string)source;
        var attachment = new MyOutlookAttachmentClass();
        _attachments.Add(attachment);
        return attachment;
    }
}

internal sealed class MyOutlookAttachmentClass : IReflectOnClrType, Attachment
{
    public MyOutlookAttachmentClass()
    {

    }
}

public enum OlItemType
{
    olMailItem = 0,
    olAppointmentItem = 1,
    olContactItem = 2,
    olTaskItem = 3,
    olJournalItem = 4,
    olNoteItem = 5,
    olPostItem = 6,
    olDistributionListItem = 7,
    olMobileItemSMS = 11,
    olMobileItemMMS = 12
}

//[ComVisible(true)]
internal sealed class MyOutlookContactItemClass : IReflectOnClrType
{
    public MyOutlookContactItemClass()
    {

    }
    [DispId(14854)]
    public string FirstName
    {
        get;
        set;
    }
    [DispId(14865)]
    public string LastName
    {
        get;
        set;
    }
    [DispId(14870)] public string CompanyName { get; set; }
    [DispId(14871)] public string JobTitle { get; set; }
    [DispId(32837)] public string BusinessAddressStreet { get; set; }
    [DispId(32838)] public string BusinessAddressCity { get; set; }
    [DispId(32839)] public string BusinessAddressState { get; set; }
    [DispId(32841)] public string BusinessAddressCountry { get; set; }
    [DispId(32840)] public string BusinessAddressPostalCode { get; set; }
    [DispId(14856)] public string BusinessTelephoneNumber { get; set; }
    [DispId(14884)] public string BusinessFaxNumber { get; set; }
    [DispId(32899)] public string Email1Address { get; set; }
    [DispId(14876)] public string MobileTelephoneNumber { get; set; }
    [DispId(61512)]
    public void Save()
    {

    }

    [DispId(61606)]
    public void Display()//[Optional][In] object Modal)
    {
        //throw new NotImplementedException($"Modal:{Modal}");
    }
}

public enum OlDefaultFolders
{
    olFolderDeletedItems = 3,
    olFolderOutbox = 4,
    olFolderSentMail = 5,
    olFolderInbox = 6,
    olFolderCalendar = 9,
    olFolderContacts = 10,
    olFolderJournal = 11,
    olFolderNotes = 12,
    olFolderTasks = 13,
    olFolderDrafts = 16,
    olPublicFoldersAllPublicFolders = 18,
    olFolderConflicts = 19,
    olFolderSyncIssues = 20,
    olFolderLocalFailures = 21,
    olFolderServerFailures = 22,
    olFolderJunk = 23,
    olFolderRssFeeds = 25,
    olFolderToDo = 28,
    olFolderManagedEmail = 29,
    olFolderSuggestedContacts = 30
}
public interface _NameSpace
{
}
internal sealed class MyOutlookMAPISession : IReflectOnClrType, _NameSpace
{
    [DispId(8459)]
    public MAPIFolder GetDefaultFolder([In] object folderType)
    {
        OlDefaultFolders xFolderType = (OlDefaultFolders)Enum.ToObject(typeof(OlDefaultFolders), folderType);
        if (xFolderType == OlDefaultFolders.olFolderContacts)
        {
            return new MyOutlookMAPIFolderContacts();
        }
        throw new NotImplementedException($"FolderType:{folderType}");
    }

    [DispId(8458)]
    public object CreateRecipient(string recipientName)
    {
        return new MyMAPIRecipientClass(recipientName);
    }

    [DispId(8460)]
    public MAPIFolder GetSharedDefaultFolder([In] object Recipient, [In] object folderType) // OlDefaultFolders
    {
        OlDefaultFolders xFolderType = (OlDefaultFolders)Enum.ToObject(typeof(OlDefaultFolders), folderType);
        if (xFolderType == OlDefaultFolders.olFolderCalendar)
        {
            return new MyOutlookMAPIFolderCalendars();
        }
        throw new NotImplementedException($"Recipient:{Recipient}, FolderType:{folderType}");
    }
}

internal sealed class MyMAPIRecipientClass : IReflectOnClrType
{
    private readonly string _recipientName;

    public MyMAPIRecipientClass(string recipientName)
    {
        _recipientName = recipientName;
    }

    private bool _resolved;
    [DispId(100)]
    public bool Resolved => _resolved;


    [DispId(113)]
    public bool Resolve()
    {
        _resolved = true;
        return true;
    }
}

internal interface MAPIFolder
{

}
interface _Items
{
}
internal sealed class MyOutlookMAPIFolderContacts : IReflectOnClrType, MAPIFolder
{
    private readonly ContactsItems _items = new ContactsItems();
    public MyOutlookMAPIFolderContacts()
    {
    }

    [DispId(12544)]
    public _Items Items
    {
        get
        {
            return _items;
        }
    }

    private sealed class ContactsItems : IReflectOnClrType, _Items
    {
        public ContactsItems()
        {
        }

        [DispId(98)]
        public object Find([In] string Filter) // Filter:[lastname] = ''  And [firstname] = ''
        {
            //throw new NotImplementedException($"Filter:{Filter}");
            return null;
        }
    }
}

internal sealed class MyOutlookMAPIFolderCalendars : IReflectOnClrType, MAPIFolder
{
    private readonly CalendarsItems _items = new CalendarsItems();
    public MyOutlookMAPIFolderCalendars()
    {
    }

    [DispId(12544)]
    public _Items Items
    {
        get
        {
            return _items;
        }
    }

    private sealed class CalendarsItems : IReflectOnClrType, _Items
    {
        public CalendarsItems()
        {
        }

        //[DispId(98)]
        //public object Find([In] string Filter) // Filter:[lastname] = ''  And [firstname] = ''
        //{
        //    //throw new NotImplementedException($"Filter:{Filter}");
        //    return null;
        //}

        [DispId(95)]
        public object Add([Optional][In] object itemType) // itemType = 1
        {
            return MyOutlookApplicationClass.CreateItemCore(itemType);
        }

    }
}

internal sealed class MyOutlookCalendarItemClass : IReflectOnClrType
{
}

internal sealed class MyOutlookAppointmentItemClass : IReflectOnClrType
{
    public MyOutlookAppointmentItemClass()
    {

    }

    [DispId(61606)]
    public void Display() // see CT98_dialog287_ButtonOutlook_Click
    {

    }


    [DispId(33293)] public object Start { get; set; }
    [DispId(55)] public string Subject { get; set; }
    [DispId(37120)] public string Body { get; set; }
    [DispId(33288)] public string Location { get; set; }
    [DispId(33299)] public object Duration { get; set; }
    [DispId(34049)] public int ReminderMinutesBeforeStart { get; set; }
    [DispId(34078)] public bool ReminderPlaySound { get; set; }
    [DispId(34051)] public bool ReminderSet { get; set; }
    //[DispId(61512)] public void Save() { }
}